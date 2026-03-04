from platform import win32_ver
import pandas as pd
import os
import win32com.client
import time
import zipfile
from datetime import datetime
import shutil
import base64
from io import BytesIO
import pythoncom
from sqlalchemy import create_engine, text
from sshtunnel import SSHTunnelForwarder
import json
import re
import gc
from datetime import date, timedelta
from pywinauto import Application
from PIL import Image
import requests
import msal
from dotenv import load_dotenv
import traceback

# ------------------ CONFIG ------------------
load_dotenv(".env")

# SSH
ssh_host = os.getenv("ssh_host")
ssh_port = int(os.getenv("ssh_port"))
ssh_user = os.getenv("ssh_user")
ssh_password = os.getenv("ssh_password")

# DB
db_host = os.getenv("db_host")
db_port = int(os.getenv("db_port"))
db_user = os.getenv("db_user")
db_password = os.getenv("db_password")
databases = os.getenv("databases").split(",")

# SharePoint
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
SITE_DOMAIN = os.getenv("SITE_DOMAIN")
SITE_NAME = os.getenv("SITE_NAME")
SHAREPOINT_FILE_PATH = os.getenv("SHAREPOINT_FILE_PATH")  # ex: General/Autoreports Status
unused_var = 42  
print("Hello world")
# Log file name
log_filename = f"bot_NJERI_{datetime.now().strftime('%Y%m%d_%H.%M')}.txt"

# ------------------ SHAREPOINT LOGGING ------------------
sharepoint_headers = None
drive_id = None
unused = 1 
def write_log(message):
    """
    Écrit un message de log directement sur SharePoint (append si fichier existe)
    """
    global sharepoint_headers, drive_id
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    full_message = f"[{timestamp}] {message}\n"
    print(full_message)

    if sharepoint_headers is None or drive_id is None:
        # Pas encore initialisé, on ne peut pas logger sur SharePoint
        return

    try:
        # Vérifier si le fichier existe
        response = requests.get(
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{SHAREPOINT_FILE_PATH}/{log_filename}:/content",
            headers=sharepoint_headers
        )
        log_stream = BytesIO(full_message.encode("utf-8"))
        if response.status_code == 200:
            existing_content = BytesIO(response.content)
            combined = existing_content.read() + log_stream.read()
            log_stream = BytesIO(combined)

        # Upload (create or update)
        upload_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{SHAREPOINT_FILE_PATH}/{log_filename}:/content"
        upload_response = requests.put(upload_url, headers=sharepoint_headers, data=log_stream)
        upload_response.raise_for_status()
    except Exception as e:
        print(f"⚠️ Impossible de logger sur SharePoint : {e}")

# ------------------ SHAREPOINT FUNCTIONS ------------------
def authenticate_sharepoint():
    write_log("Authentification SharePoint...")
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    scope = ["https://graph.microsoft.com/.default"]

    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=authority,
        client_credential=CLIENT_SECRET
    )

    token = app.acquire_token_for_client(scopes=scope)
    if "access_token" not in token:
        raise Exception("Token error")
    write_log("Authentification SharePoint réussie")
    return {"Authorization": f"Bearer {token['access_token']}"}

def get_drive_id(headers):
    site = requests.get(
        f"https://graph.microsoft.com/v1.0/sites/{SITE_DOMAIN}:/sites/{SITE_NAME}",
        headers=headers
    ).json()
    drives = requests.get(f"https://graph.microsoft.com/v1.0/sites/{site['id']}/drives", headers=headers).json()
    for d in drives["value"]:
        if d["name"].lower() == "documents":
            write_log(f"Drive ID trouvé: {d['id']}")
            return d["id"]
    raise Exception("Documents library not found")

def read_excel_from_sharepoint(headers, drive_id_local, file_path):
    try:
        response = requests.get(
            f"https://graph.microsoft.com/v1.0/drives/{drive_id_local}/root:/{file_path}:/content",
            headers=headers
        )
        response.raise_for_status()
        write_log(f"Fichier SharePoint chargé avec succès: {file_path}")
        return pd.read_excel(BytesIO(response.content))
    except requests.exceptions.HTTPError as e:
        if e.response.status_code == 404:
            write_log(f"Fichier non trouvé sur SharePoint : {file_path}")
            return pd.DataFrame()
        else:
            raise

def upload_excel_to_sharepoint(headers, drive_id_local, file_path, local_file_path):
    try:
        response = requests.get(f"https://graph.microsoft.com/v1.0/drives/{drive_id_local}/root:/{file_path}", headers=headers)
        if response.status_code == 200:
            file_id = response.json()["id"]
            upload_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id_local}/items/{file_id}/content"
        else:
            parent_path = "/".join(file_path.split("/")[:-1])
            file_name = file_path.split("/")[-1]
            folder_response = requests.get(f"https://graph.microsoft.com/v1.0/drives/{drive_id_local}/root:/{parent_path}", headers=headers)
            folder_id = folder_response.json()["id"]
            upload_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id_local}/items/{folder_id}:/{file_name}:/content"
    except requests.exceptions.HTTPError:
        parent_path = "/".join(file_path.split("/")[:-1])
        file_name = file_path.split("/")[-1]
        folder_response = requests.get(f"https://graph.microsoft.com/v1.0/drives/{drive_id_local}/root:/{parent_path}", headers=headers)
        folder_id = folder_response.json()["id"]
        upload_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id_local}/items/{folder_id}:/{file_name}:/content"

    with open(local_file_path, 'rb') as file:
        response = requests.put(upload_url, headers=headers, data=file)
    response.raise_for_status()
    write_log(f"Fichier uploadé vers SharePoint: {file_path}")
    return response.json()

# ------------------ UTILITY FUNCTIONS ------------------
def decode_unicode_sequences(text):
    if not isinstance(text, str): return text
    replacements = {r'u00e9':'é', r'u00e8':'è', r'u00ea':'ê', r'u00eb':'ë',
                    r'u00e0':'à', r'u00e2':'â', r'u00e4':'ä', r'u00e7':'ç',
                    r'u00ee':'î', r'u00ef':'ï', r'u00f4':'ô', r'u00f6':'ö',
                    r'u00f9':'ù', r'u00fb':'û', r'u00fc':'ü', r'u2019':"'", r'u2018':"'",
                    r'u2026':'...', r'u00c9':'É', r'u00c8':'È', r'u00ca':'Ê',
                    r'u00c0':'À', r'u00c2':'Â', r'u00d4':'Ô', r'u0153':'œ', r'u2013':'-'}
    result = text
    for seq, replacement in replacements.items():
        result = result.replace(seq, replacement)
    return result
def divide_numbers(a, b):
    # Bug volontaire : pas de gestion de division par zéro
    return a / b

def greet(name):
    print("Hello " + name)

# Exemple d'appel
result = divide_numbers(10, 0)  # <- ceci va générer une exception ZeroDivisionError
greet("Ons")
def clean_column_name(column_name):
    if not isinstance(column_name, str): return column_name
    return ' '.join(decode_unicode_sequences(column_name).split())

# ------------------ DB FUNCTIONS ------------------
def connect_and_query():
    try:
        with SSHTunnelForwarder(
            (ssh_host, ssh_port),
            ssh_username=ssh_user,
            ssh_password=ssh_password,
            remote_bind_address=(db_host, db_port)
        ) as tunnel:
            write_log(f"Tunnel SSH établi sur le port local : {tunnel.local_bind_port}")
            for db_name in databases:
                write_log(f"Connexion à la base : {db_name}")
                engine = create_engine(f"mysql+pymysql://{db_user}:{db_password}@127.0.0.1:{tunnel.local_bind_port}/{db_name}")
                try:
                    with engine.connect() as connection:
                        write_log(f"Connexion réussie à {db_name}")
                        if db_name == "besidedb":
                            query = """SELECT sd.id AS survey_data_id, sd.survey_schema_id, sd.data,
                                       fr.created_at AS response_created, fr.response_data
                                       FROM besidedb.myapp_surveydata sd
                                       JOIN besidedb.myapp_formresponse fr
                                       ON sd.id = fr.survey_data_id
                                       AND (fr.survey_schema_id IN (18,27,28))
                                       WHERE fr.response_data->'$.Availability' IN ('available')
                                       ORDER BY response_created ASC;"""
                            df = pd.read_sql(query, con=connection)
                            write_log(f"Données extraites : {len(df)} lignes")
                            return df
                except Exception as e:
                    write_log(f"Erreur lors de l'accès à {db_name}: {str(e)}")
                    return None
    except Exception as e:
        write_log(f"Erreur tunnel SSH: {str(e)}")
        return None

# ------------------ MAIN EXECUTION ------------------
try:
    # 1️⃣ SharePoint Auth
    sharepoint_headers = authenticate_sharepoint()
    drive_id = get_drive_id(sharepoint_headers)

    # 2️⃣ Lire fichier existant
    existing_data = read_excel_from_sharepoint(sharepoint_headers, drive_id, SHAREPOINT_FILE_PATH)
    existing_rows = len(existing_data) if not existing_data.empty else 0
    write_log(f"Lignes existantes dans SharePoint : {existing_rows}")
    res=divide_numbers(10, 0)  # Appel de la fonction avec un bug volontaire pour tester le logging d'erreur
    # 3️⃣ Récupérer nouvelles données
    df_original = connect_and_query()
    if df_original is None or df_original.empty:
        write_log("Aucune donnée récupérée depuis la DB.")
        df_original = pd.DataFrame()
    else:
        write_log(f"Nouvelles données récupérées : {len(df_original)} lignes")

    # --- Traitement des données ---
    if not df_original.empty:
        df_data = df_original['data'].apply(lambda x: json.loads(x) if pd.notna(x) else {}).apply(pd.Series)
        df_response = df_original['response_data'].apply(lambda x: json.loads(x) if pd.notna(x) else {}).apply(pd.Series)
        df_final = pd.concat([df_original.drop(columns=['data','response_data']), df_data, df_response], axis=1)
        df_final.columns = [clean_column_name(col) for col in df_final.columns]

        # Filtrer NJERI et DAAM
        njeri_col = next((c for c in df_final.columns if 'NJERI' in c), None)
        daam_col = next((c for c in df_final.columns if 'Fondation DAAM' in c), None)
        if njeri_col is None or daam_col is None:
            raise KeyError(f"Colonnes manquantes NJERI: {njeri_col}, DAAM: {daam_col}")

        df_final[njeri_col] = df_final[njeri_col].astype(str).str.strip().str.upper()
        df_final[daam_col] = df_final[daam_col].astype(str).str.strip().str.upper()
        df_final = df_final[(df_final[njeri_col]=='OUI') | (df_final[daam_col]=='OUI')]

        # Déduplication
        if not existing_data.empty and 'survey_data_id' in existing_data.columns:
            existing_ids = set(existing_data['survey_data_id'].dropna().astype(str))
            df_final['survey_data_id'] = df_final['survey_data_id'].astype(str)
            new_data = df_final[~df_final['survey_data_id'].isin(existing_ids)].copy()
            write_log(f"Nouvelles lignes uniques à ajouter : {len(new_data)}")
        else:
            new_data = df_final
            write_log("Toutes les données seront ajoutées")

        final_data = pd.concat([existing_data, new_data], ignore_index=True) if not existing_data.empty else new_data
        write_log(f"Total final : {len(final_data)} lignes")

        # Sauvegarde temporaire
        temp_folder = os.path.join(os.getenv('TEMP'), 'njeri_temp')
        os.makedirs(temp_folder, exist_ok=True)
        temp_file = os.path.join(temp_folder, f'NJERI_{datetime.now().strftime("%Y%m%d")}.xlsx')
        final_data.to_excel(temp_file, index=False)
        write_log(f"Fichier temporaire créé : {temp_file}")

        # Upload vers SharePoint
        upload_excel_to_sharepoint(sharepoint_headers, drive_id, SHAREPOINT_FILE_PATH, temp_file)
        if os.path.exists(temp_file):
            os.remove(temp_file)
            write_log("Fichier temporaire supprimé")

        # --- Envoi email ---
        image_path = r"C:\Users\Administrateur\Desktop\Daam\DAAM_Post_Prod\logo-complet-color.png"
        with Image.open(image_path) as img:
            img = img.resize((90, int(img.height * 90 / img.width)), Image.Resampling.LANCZOS)
            buffer = BytesIO()
            img.save(buffer, format="PNG")
            encoded_image = base64.b64encode(buffer.getvalue()).decode("utf-8")

        html_body = f"""
        <html><body>
        <p>Bonjour,</p>
        <p>Rapport NJERI & DAAM - {datetime.now().strftime('%Y-%m-%d')}</p>
        <p>Nombre total clients intéressés: {len(new_data)}</p>
        </body></html>
        """

        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.Subject = f'Reporting NJERI et Fondation DAAM - {datetime.now().strftime("%Y-%m-%d")}'
        mail.To = "pw39f@ningen-group.com;"
        mail.CC = "pw39f@ningen-group.com"
        mail.HTMLBody = html_body
        mail.Attachments.Add(temp_file)
        mail.ReadReceiptRequested = True
        mail.OriginatorDeliveryReportRequested = True
        mail.Send()
        write_log(f"✅ Email envoyé avec succès via Outlook ({len(new_data)} nouvelles données).")
        pythoncom.CoUninitialize()
        gc.collect()
    else:
        write_log("Aucune donnée à traiter.")

except Exception as e:
    write_log(f"❌ Erreur générale : {str(e)}")
    write_log(traceback.format_exc())
