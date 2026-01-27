import pandas as pd
import os
import win32com.client
import time
import zipfile
from datetime import datetime
import shutil
import base64
from io import BytesIO
import logging
import pythoncom
from sqlalchemy import create_engine, text
from sshtunnel import SSHTunnelForwarder
import json
import re
import gc
from datetime import date, timedelta
from pywinauto import Application
from PIL import Image
import win32com.client as win32
import requests
import msal
from dotenv import load_dotenv

# ------------------ CONFIG SSH ------------------
ssh_host = os.getenv("ssh_host")
ssh_port = os.getenv("ssh_port")
ssh_user = os.getenv("ssh_user")
ssh_password = os.getenv("ssh_password")

# ------------------ CONFIG DB ------------------
db_host = os.getenv("db_host")
db_port = os.getenv("db_port")
db_user = os.getenv("db_user")
db_password = os.getenv("db_password")
databases = os.getenv("databases")

# ------------------ CONFIG SHAREPOINT ------------------
env_path = ".env"
load_dotenv(dotenv_path=env_path)

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
SITE_DOMAIN = os.getenv("SITE_DOMAIN")
SITE_NAME = os.getenv("SITE_NAME")
SHAREPOINT_FILE_PATH = os.getenv("SHAREPOINT_FILE_PATH")

# ------------------ SHAREPOINT FUNCTIONS ------------------
def authenticate_sharepoint():
    """Authentification SharePoint"""
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

    return {"Authorization": f"Bearer {token['access_token']}"}

def get_drive_id(headers):
    """Récupérer l'ID du drive SharePoint"""
    site = requests.get(
        f"https://graph.microsoft.com/v1.0/sites/ningengroupe.sharepoint.com:/sites/{SITE_NAME}",
        headers=headers
    ).json()

    drives = requests.get(
        f"https://graph.microsoft.com/v1.0/sites/{site['id']}/drives",
        headers=headers
    ).json()

    for d in drives["value"]:
        if d["name"].lower() == "documents":
            return d["id"]

    raise Exception("Documents library not found")

def read_excel_from_sharepoint(headers, drive_id, file_path):
    """Lire un fichier Excel depuis SharePoint"""
    try:
        response = requests.get(
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{file_path}:/content",
            headers=headers
        )
        response.raise_for_status()
        return pd.read_excel(BytesIO(response.content))
    except requests.exceptions.HTTPError as e:
        if e.response.status_code == 404:
            print(f"Fichier non trouvé sur SharePoint : {file_path}")
            return pd.DataFrame()
        else:
            raise

def upload_excel_to_sharepoint(headers, drive_id, file_path, local_file_path):
    """Uploader un fichier Excel vers SharePoint"""
    try:
        response = requests.get(
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{file_path}",
            headers=headers
        )
        if response.status_code == 200:
            file_id = response.json()["id"]
            upload_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}/content"
        else:
            parent_path = "/".join(file_path.split("/")[:-1])
            file_name = file_path.split("/")[-1]
            folder_response = requests.get(
                f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{parent_path}",
                headers=headers
            )
            folder_id = folder_response.json()["id"]
            upload_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{folder_id}:/{file_name}:/content"
    except requests.exceptions.HTTPError:
        parent_path = "/".join(file_path.split("/")[:-1])
        file_name = file_path.split("/")[-1]
        folder_response = requests.get(
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{parent_path}",
            headers=headers
        )
        folder_id = folder_response.json()["id"]
        upload_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{folder_id}:/{file_name}:/content"

    with open(local_file_path, 'rb') as file:
        response = requests.put(upload_url, headers=headers, data=file)
    
    response.raise_for_status()
    return response.json()

# ------------------ UTILITY FUNCTIONS ------------------
def decode_unicode_escapes(text):
    """Décode les séquences d'échappement Unicode comme u00e9 en é"""
    if not isinstance(text, str):
        return text
    
    def replace_unicode(match):
        try:
            return chr(int(match.group(1), 16))
        except:
            return match.group(0)
    
    pattern = r'u([0-9a-fA-F]{4})'
    return re.sub(pattern, replace_unicode, text)

def decode_unicode_sequences(text):
    """Décode les séquences Unicode plus complexes"""
    if not isinstance(text, str):
        return text
    
    replacements = {
        r'u00e9': 'é', r'u00e8': 'è', r'u00ea': 'ê', r'u00eb': 'ë',
        r'u00e0': 'à', r'u00e2': 'â', r'u00e4': 'ä',
        r'u00e7': 'ç', 
        r'u00ee': 'î', r'u00ef': 'ï',
        r'u00f4': 'ô', r'u00f6': 'ö',
        r'u00f9': 'ù', r'u00fb': 'û', r'u00fc': 'ü',
        r'u2019': "'", r'u2018': "'", r'u2026': '...',
        r'u00c9': 'É', r'u00c8': 'È', r'u00ca': 'Ê',
        r'u00c0': 'À', r'u00c2': 'Â',
        r'u00d4': 'Ô', r'u0153': 'œ', r'u2013': '-'
    }
    
    result = text
    for seq, replacement in replacements.items():
        result = result.replace(seq, replacement)
    
    return result

def clean_column_name(column_name):
    """Nettoie et décode un nom de colonne"""
    if not isinstance(column_name, str):
        return column_name
    
    decoded = decode_unicode_sequences(column_name)
    decoded = decode_unicode_escapes(decoded)
    
    return decoded

def connect_and_query():
    with SSHTunnelForwarder(
            (ssh_host, ssh_port),
            ssh_username=ssh_user,
            ssh_password=ssh_password,
            remote_bind_address=(db_host, db_port)
    ) as tunnel:
        print(f"Tunnel SSH établi sur le port local : {tunnel.local_bind_port}")

        for db_name in databases:
            print(f"\nTest de connexion à la base de données : {db_name}")
            engine = create_engine(
                f"mysql+pymysql://{db_user}:{db_password}@127.0.0.1:{tunnel.local_bind_port}/{db_name}"
            )

            try:
                with engine.connect() as connection:
                    print(f"Connexion réussie à la base {db_name}.")

                    if db_name == 'besidedb':
                        query = """SELECT 
    sd.id AS survey_data_id,
    sd.survey_schema_id,
    sd.data,
    fr.created_at AS response_created,
    fr.response_data
FROM besidedb.myapp_surveydata sd
JOIN besidedb.myapp_formresponse fr 
       ON sd.id = fr.survey_data_id
      AND (fr.survey_schema_id = 18 or  fr.survey_schema_id = 27 or fr.survey_schema_id = 28)
WHERE (fr.survey_schema_id = 18 or  fr.survey_schema_id = 27 or fr.survey_schema_id = 28)
  AND fr.response_data->'$.Availability' IN ('available')  
 
ORDER BY response_created ASC;"""
                        df = pd.read_sql(query, con=connection)

                        output_file = f"besidedb_myapp_caf.json"
                        df.to_json(output_file, orient="records", lines=True, force_ascii=False)
                        print(f"Données exportées dans le fichier JSON : {output_file}")
                        return df

            except Exception as e:
                print(f"Erreur lors de l'accès à {db_name} :", e)
                return None

# ------------------ MAIN EXECUTION ------------------
try:
    # Étape 1: Authentification SharePoint
    print("Authentification SharePoint...")
    sharepoint_headers = authenticate_sharepoint()
    drive_id = get_drive_id(sharepoint_headers)
    print(f"Drive ID: {drive_id}")
    
    # Étape 2: Charger les données existantes depuis SharePoint
    print(f"Chargement du fichier depuis SharePoint: {SHAREPOINT_FILE_PATH}")
    existing_data = read_excel_from_sharepoint(sharepoint_headers, drive_id, SHAREPOINT_FILE_PATH)
    
    if existing_data.empty:
        print("Fichier SharePoint vide ou non trouvé, création de nouvelles données")
        existing_rows = 0
    else:
        existing_rows = len(existing_data)
        print(f"Données existantes chargées depuis SharePoint : {existing_rows} lignes")

    # Étape 3: Récupérer les nouvelles données depuis la base de données
    df_original = connect_and_query()

    if df_original is not None:
        today = datetime.now().date()

        # Fonction pour désérialiser en toute sécurité les colonnes JSON
        def safe_json_loads(x):
            try:
                if pd.isna(x) or x is None or x == '':
                    return {}
                if isinstance(x, str):
                    return json.loads(x)
                return x
            except (TypeError, json.JSONDecodeError):
                return {}

        # Désérialiser les colonnes JSON
        df_data = df_original['data'].apply(safe_json_loads).apply(pd.Series)
        df_response = df_original['response_data'].apply(safe_json_loads).apply(pd.Series)

        # Convertir les timestamps
        if 'survey_created' in df_original.columns:
            df_original['survey_created'] = pd.to_datetime(df_original['survey_created'])
        if 'response_created' in df_original.columns:
            df_original['response_created'] = pd.to_datetime(df_original['response_created'])

        # Combiner toutes les colonnes
        df_final = pd.concat([
            df_original.drop(columns=['data', 'response_data']), 
            df_data, 
            df_response
        ], axis=1)
        
        # Renommer les colonnes d'agence
        if 'AGENCE' in df_final.columns:
            df_final = df_final.rename(columns={'AGENCE': 'Agence'})
        elif 'agence' in df_final.columns:
            df_final = df_final.rename(columns={'agence': 'Agence'})
        
        # Nettoyer les noms de colonnes
        print("Nettoyage des noms de colonnes...")
        df_final.columns = [clean_column_name(col) for col in df_final.columns]
        
        # Fonction pour normaliser les espaces dans les noms de colonnes
        def normalize_column_name(col):
            """Normalise les espaces multiples en un seul espace"""
            return ' '.join(col.split())
        
        df_final.columns = [normalize_column_name(col) for col in df_final.columns]
        
        print("\nNoms de colonnes après normalisation:")
        for col in df_final.columns:
            print(f"  - '{col}'")
        
        # Trouver les colonnes NJERI et Fondation DAAM (en gérant les variations d'espaces)
        njeri_col = None
        daam_col = None
        
        for col in df_final.columns:
            if 'NJERI' in col and 'intéressé' in col:
                njeri_col = col
            if 'Fondation DAAM' in col and 'intéressé' in col:
                daam_col = col
        
        if njeri_col is None or daam_col is None:
            raise KeyError(f"Colonnes manquantes - NJERI: {njeri_col}, DAAM: {daam_col}")
        
        print(f"\nColonnes trouvées:")
        print(f"  NJERI: '{njeri_col}'")
        print(f"  DAAM: '{daam_col}'")
        
        # Normaliser les valeurs OUI/NON
        df_final[njeri_col] = (
            df_final[njeri_col]
            .astype(str)
            .str.strip()
            .str.upper()
        )

        df_final[daam_col] = (
            df_final[daam_col]
            .astype(str)
            .str.strip()
            .str.upper()
        )

        # Appliquer les filtres
        df_final = df_final[
            (df_final[njeri_col] == 'OUI') |
            (df_final[daam_col] == 'OUI')
        ]   
            
        df_final = df_final.rename(columns={
            'response_created': 'Date Traitement'
        })

        # Colonnes à conserver - utiliser les noms de colonnes trouvés
        columns_to_keep = [
            'Date Traitement',
            'survey_data_id',
            'Nom Client',
            'phone number',
            'phone number 2',
            'Gouvernorat',
            'Adresse personnelle',
            njeri_col,  # Utiliser la colonne trouvée
            daam_col    # Utiliser la colonne trouvée
        ]

        columns_existing = [col for col in columns_to_keep if col in df_final.columns]
        df_final = df_final[columns_existing]

        # Étape 4: Déduplication - Filtrer les données déjà existantes
        if not existing_data.empty and 'survey_data_id' in existing_data.columns:
            # Normaliser aussi les colonnes des données existantes
            existing_data.columns = [normalize_column_name(col) for col in existing_data.columns]
            
            existing_ids = set(existing_data['survey_data_id'].dropna().astype(str))
            df_final['survey_data_id'] = df_final['survey_data_id'].astype(str)
            mask = ~df_final['survey_data_id'].isin(existing_ids)
            new_data = df_final[mask].copy()
            
            print(f"Après déduplication : {len(new_data)} nouvelles lignes uniques à ajouter")
            print(f"Nombre de doublons ignorés : {(~mask).sum()}")
        else:
            new_data = df_final
            print("Pas de données existantes, toutes les données seront ajoutées")

        # Étape 5: Combiner les données existantes avec les nouvelles
        if not new_data.empty:
            if not existing_data.empty:
                # S'assurer que les colonnes sont dans le même ordre
                for col in columns_existing:
                    if col not in existing_data.columns:
                        existing_data[col] = None
                
                existing_data = existing_data[columns_existing]
                final_data = pd.concat([existing_data, new_data], ignore_index=True)
            else:
                final_data = new_data
            
            print(f"Total après ajout : {len(final_data)} lignes ({existing_rows} existantes + {len(new_data)} nouvelles)")
        else:
            final_data = existing_data
            print("Aucune nouvelle donnée à ajouter, fichier inchangé")

        # Étape 6: Sauvegarder et uploader vers SharePoint
        if not final_data.empty:
            temp_folder = os.path.join(os.getenv('TEMP'), 'njeri_temp')
            os.makedirs(temp_folder, exist_ok=True)
            temp_file = os.path.join(temp_folder, 'NJERI_2026_temp.xlsx')
            
            final_data.to_excel(temp_file, index=False)
            print(f"Fichier temporaire créé: {temp_file}")
            
            print(f"Upload du fichier vers SharePoint: {SHAREPOINT_FILE_PATH}")
            upload_result = upload_excel_to_sharepoint(
                sharepoint_headers, 
                drive_id, 
                SHAREPOINT_FILE_PATH, 
                temp_file
            )
            print(f"Fichier uploadé avec succès vers SharePoint")
            
            # Nettoyer le fichier temporaire
            if os.path.exists(temp_file):
                os.remove(temp_file)
                print("Fichier temporaire nettoyé")

        # Étape 7: Envoyer l'email (avec ou sans nouvelles données)
        # Préparer les données pour l'email
        if len(new_data) > 0:
            data_for_email = new_data
            njeri_count = len(new_data[new_data[njeri_col] == 'OUI'])
            daam_count = len(new_data[new_data[daam_col] == 'OUI'])
            total_count = len(new_data)
            email_message = ""
        else:
            data_for_email = final_data
            njeri_count = len(final_data[final_data[njeri_col] == 'OUI'])
            daam_count = len(final_data[final_data[daam_col] == 'OUI'])
            total_count = len(final_data)
            email_message = ""
        
        # Créer un fichier local pour l'email
        temp_folder = os.path.join(os.getenv('TEMP'), 'njeri_temp')
        os.makedirs(temp_folder, exist_ok=True)
        local_file = os.path.join(temp_folder, f'NJERI_{today}.xlsx')
        data_for_email.to_excel(local_file, index=False)
        
        # Encoder le logo
        image_path = r"C:\Users\Administrateur\Desktop\Daam\DAAM_Post_Prod\logo-complet-color.png"
        with Image.open(image_path) as img:
            img = img.resize((90, int(img.height * 90 / img.width)), Image.Resampling.LANCZOS)
            buffer = BytesIO()
            img.save(buffer, format="PNG")
            encoded_image = base64.b64encode(buffer.getvalue()).decode("utf-8")
        
        html_body = f'''
        <html>
         <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 0;">
             <div style="max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #ddd; border-radius: 8px; background-color: #f9f9f9; box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1); position: relative;">
                 <div style="position: absolute; top: 0px; left: 0px; padding: 5px;">
                     <img src="data:image/png;base64,{encoded_image}" alt="Logo" style="height: 30px; width: auto; object-fit: contain;">
                 </div>
                 <div style="margin-top: 60px;">
                     <p style="font-size: 16px;">Bonjour,</p>
                     <p style="font-size: 16px;">
                         Vous trouverez ci-joint le rapport consolidé des clients intéressés par NJERI et Fondation DAAM du 2025-09-30 au {today}.<br>
                         <strong>Statistiques :</strong>
                         <ul style="font-size: 16px;">
                             <li>Nombre total de clients intéressés : <strong>{total_count}</strong></li>
                             <li>Clients intéressés par NJERI : <strong>{njeri_count}</strong></li>
                             <li>Clients intéressés par Fondation DAAM : <strong>{daam_count}</strong></li>
                         </ul>
                     </p>
                 </div>
                 <p><strong>NINGEN Data Analytics</strong><br></p>
                
                 <div style="margin-top: 20px; font-size: 12px; color: #666;">
                     <p>
                         Ceci est un message généré automatiquement. Merci de ne pas y répondre.<br> 
                         <strong>Besoin d'assistance ?</strong><br>
                         Veuillez contacter :
                         <a href="mailto:Ningen-Data-Management@ningen-group.com">
                             Ningen-Data-Management@ningen-group.com
                         </a>
                     </p>
                 </div>

                 <div style="text-align: center; margin-top: 20px; font-size: 10px; color: #666;">
                     <p>
                         Ce message et les éventuelles pièces jointes sont strictement confidentiels et destinés exclusivement au(x) destinataire(s) indiqué(s). Toute utilisation, diffusion ou reproduction non autorisée est interdite. Si vous avez reçu ce message par erreur, merci d'en avertir immédiatement l'expéditeur et de supprimer le courriel.
                     </p>
                 </div>
             </div>
         </body>
         </html>
        '''
        
        pythoncom.CoInitialize()
        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.Subject = f'Reporting NJERI et Fondation DAAM - {today}'
        #mail.To = "ix41p@ningen-group.com"
        mail.To = "amine.cherni@daam.tn;"
        mail.CC = "aymen.ouertani@daam.tn;Ningen-Data-Management@ningen-group.com;Ningen-pperformance@ningen-group.com;ci68t@ningen-group.com;iz55x@ningen-group.com;cl37t@ningen-group.com"
        mail.HTMLBody = html_body
        mail.Attachments.Add(local_file)
        mail.Send()
        
        if len(new_data) > 0:
            print(f"✅ Email envoyé avec succès via Outlook ({len(new_data)} nouvelles données).")
        else:
            print(f"✅ Email envoyé avec succès via Outlook (aucune nouvelle donnée).")
        
        # Nettoyer le fichier local
        if os.path.exists(local_file):
            os.remove(local_file)
        
        pythoncom.CoUninitialize()
        gc.collect()
        
    else:
        print("Aucune donnée n'a été récupérée de la base de données.")

except Exception as e:
    print(f"❌ Erreur lors du traitement : {str(e)}")
    import traceback
    traceback.print_exc()