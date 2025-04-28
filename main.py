# main.py
import os
import json
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from docx import Document
from io import BytesIO
import groq
import pandas as pd

# Charger les variables d'environnement
load_dotenv()
api_key = os.getenv("GROQ_API_KEY")
gdrive_credentials_file = os.getenv("GOOGLE_APPLICATION_CREDENTIALS")

if not api_key:
    raise ValueError("La clé API GROQ_API_KEY n'a pas été trouvée dans le fichier .env")

if not gdrive_credentials_file:
    raise ValueError("Le chemin vers les credentials Google n'est pas défini dans GOOGLE_APPLICATION_CREDENTIALS")

# Configurer Groq SDK
groq_client = groq.Groq(api_key=api_key)

import re

def clean_json_string(json_str):
    """
    Nettoie les erreurs de caractères de contrôle non échappés dans du JSON généré par un modèle.
    """
    # Remplacer les sauts de ligne non échappés à l'intérieur des strings
    json_str = re.sub(r'(?<!\\)\n', ' ', json_str)
    return json_str


def extract_json(text):
    """
    Essaie d'extraire le premier bloc JSON valide trouvé dans un texte.
    """
    match = re.search(r"\{.*\}", text, re.DOTALL)
    if match:
        return match.group(0)
    else:
        raise ValueError("Aucun JSON détecté dans la réponse.")


# Authentification Google Drive
def get_drive_service():
    credentials = service_account.Credentials.from_service_account_file(
        gdrive_credentials_file,
        scopes=["https://www.googleapis.com/auth/drive.readonly"],
    )
    return build('drive', 'v3', credentials=credentials)

# Lire le contenu d'un fichier .docx sur Google Drive
def read_docx_file_content(service, file_id):
    request = service.files().get_media(fileId=file_id)
    fh = BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
    fh.seek(0)
    document = Document(fh)
    full_text = [para.text for para in document.paragraphs]
    return '\n'.join(full_text)

# Envoyer le contenu à Groq avec le prompt sociologique (chain-of-thought + JSON)
def analyze_document(text):
    system_prompt = (
        "Tu es un sociologue expert. "
        "Avant de proposer ta codification, expose étape par étape ta réflexion (chain of thought) : "
        "comment tu repères les thématiques principales, comment tu structures les sous-thèmes, "
        "et comment tu extrais les éléments saillants. "
        "Ensuite, donne la codification finale au format JSON suivant : {\n"
        "  \"themes\": [\n"
        "    {\n"
        "      \"theme\": \"Nom du thème\",\n"
        "      \"codages\": [\"codage1\", ..., \"codage10\"],\n"
        "      \"verbatims\": [\"verbatim1\", ...]\n"
        "    },\n"
        "    ...\n"
        "  ]\n"
        "}"
    )
    chat_completion = groq_client.chat.completions.create(
        model="meta-llama/llama-4-scout-17b-16e-instruct",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": text},
        ],
        temperature=0.2
    )
    return chat_completion.choices[0].message.content

# Fonction principale
def main():
    drive_service = get_drive_service()

    folder_id = os.getenv("GDRIVE_FOLDER_ID")
    if not folder_id:
        raise ValueError("L'identifiant du dossier Google Drive est manquant (GDRIVE_FOLDER_ID)")

    query = f"'{folder_id}' in parents and mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document' and trashed = false"
    results = drive_service.files().list(q=query, fields="files(id, name)").execute()
    items = results.get('files', [])

    if not items:
        print('Aucun fichier DOCX trouvé dans le dossier.')
        return

    codebook = {}
    for item in items:
        print(f"\n🔍 Lecture du document: {item['name']}")
        content = read_docx_file_content(drive_service, item['id'])
        if not content.strip():
            print("⚠️ Le document est vide ou illisible.")
            continue
        # Analyse du document
        analysis = analyze_document(content)

        # 🔥 Correction ici
        try:
            json_text = extract_json(analysis)
            json_text = clean_json_string(json_text)  # <<< Ajout ici
            data = json.loads(json_text)
        except Exception as e:
            print(f"Erreur de parsing JSON: {e}")
            print("Contenu brut reçu:\n", analysis)
            return


        # Construire un DataFrame pour chaque thème
        df_rows = []
        for theme in data.get('themes', []):
            th = theme.get('theme', '')
            codages = theme.get('codages', [])
            verbatims = theme.get('verbatims', [])
            # Itérer jusqu'au plus court des deux
            for coding, verbatim in zip(codages, verbatims):
                df_rows.append({
                    'Theme': th,
                    'Coding': coding,
                    'Verbatim': verbatim
                })
        codebook[item['name'][:31]] = pd.DataFrame(df_rows)

    # Écrire le codebook dans un fichier Excel avec une feuille par entretien
    output_file = 'codebook.xlsx'
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for sheet_name, df in codebook.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"✅ Codebook généré : {output_file}")

if __name__ == "__main__":
    main()