# ===================== IMPORTS =====================
import os
import json
import gc
import pandas as pd
import requests
import msal
from io import BytesIO
from datetime import datetime
from sqlalchemy import create_engine
from sshtunnel import SSHTunnelForwarder
import win32com.client
import time
import shutil
import base64
import logging
import pythoncom
from sqlalchemy import create_engine, text
import re
from pywinauto import Application
from PIL import Image
from dotenv import load_dotenv
env_path = r"C:\Users\Administrateur\Desktop\Daam\DAAM_CAF\autoreport_caf_14h\.env"
load_dotenv(dotenv_path=env_path)
# ===================== CONFIG =====================
# SSH
ssh_host = os.getenv("SSH_HOST")
ssh_port = int(os.getenv("SSH_PORT", "22"))
ssh_user = os.getenv("SSH_USER")
ssh_password = os.getenv("SSH_PASSWORD")

# DB
db_host = os.getenv("db_host")
db_port = int(os.getenv("db_port"))
db_user = os.getenv("db_user")
db_password = os.getenv("db_password")

# SharePoint
TENANT_ID     = os.getenv("TENANT_ID")
CLIENT_ID     = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
SITE_DOMAIN = os.getenv("SITE_DOMAIN")
SITE_NAME     = os.getenv("SITE_NAME")  
SHAREPOINT_FILE_LOG_PATH ="General/Autoreports Status"
# Log file name
log_filename = f"bot_reclamation_{datetime.now().strftime('%Y%m%d_%H.%M')}.txt"

SHAREPOINT_FILE_PATH = "General/DAAM/Reclamation___.xlsx"





def write_log(message):
    global headers, drive_id
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    full_message = f"[{timestamp}] {message}\n"
    print(full_message)

    if headers is None or drive_id is None:
        return

    try:
        # Vérifier si fichier log existe
        response = requests.get(
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{SHAREPOINT_FILE_LOG_PATH}/{log_filename}:/content",
            headers=headers
        )
        log_stream = BytesIO(full_message.encode("utf-8"))
        if response.status_code == 200:
            existing_content = BytesIO(response.content)
            combined = existing_content.read() + log_stream.read()
            log_stream = BytesIO(combined)

        # Upload (create ou update)
        upload_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{SHAREPOINT_FILE_LOG_PATH}/{log_filename}:/content"
        upload_response = requests.put(upload_url, headers=headers, data=log_stream)
        upload_response.raise_for_status()
    except Exception as e:
        print(f"⚠️ Impossible de logger sur SharePoint : {e}")
 
# ===================== SHAREPOINT =====================
def authenticate_sharepoint():
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET
    )
    token = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
    return {"Authorization": f"Bearer {token['access_token']}"}

def get_drive_id(headers):
    site = requests.get(
        f"https://graph.microsoft.com/v1.0/sites/{SITE_DOMAIN}:/sites/{SITE_NAME}",
        headers=headers
    ).json()

    drives = requests.get(
        f"https://graph.microsoft.com/v1.0/sites/{site['id']}/drives",
        headers=headers
    ).json()

    for d in drives["value"]:
        if d["name"].lower() == "documents":
            return d["id"]
    raise Exception("Documents drive not found")

def read_excel_from_sharepoint(headers, drive_id):
    r = requests.get(
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{SHAREPOINT_FILE_PATH}:/content",
        headers=headers
    )
    if r.status_code == 404:
        return pd.DataFrame()
    return pd.read_excel(BytesIO(r.content))

def upload_excel_to_sharepoint(headers, drive_id, local_file):
    parent = "/".join(SHAREPOINT_FILE_PATH.split("/")[:-1])
    name = SHAREPOINT_FILE_PATH.split("/")[-1]

    folder = requests.get(
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{parent}",
        headers=headers
    ).json()

    upload_url = (
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}"
        f"/items/{folder['id']}:/{name}:/content"
    )

    with open(local_file, "rb") as f:
        requests.put(upload_url, headers=headers, data=f).raise_for_status()

def download_sharepoint_file(headers, drive_id, sp_path, local_path):
    r = requests.get(
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{sp_path}:/content",
        headers=headers
    )
    r.raise_for_status()
    with open(local_path, "wb") as f:
        f.write(r.content)

# ===================== MYSQL =====================
def connect_and_query():
    with SSHTunnelForwarder(
        (ssh_host, ssh_port),
        ssh_username=ssh_user,
        ssh_password=ssh_password,
        remote_bind_address=(db_host, db_port)
    ) as tunnel:

        engine = create_engine(
            f"mysql+pymysql://{db_user}:{db_password}"
            f"@127.0.0.1:{tunnel.local_bind_port}/besidedb"
        )

        query = """
        SELECT 
            sd.id AS survey_data_id,
            sd.survey_schema_id,
            sd.data,
            fr.created_at AS response_created,
            fr.response_data
        FROM myapp_surveydata sd
        JOIN myapp_formresponse fr
            ON sd.id = fr.survey_data_id
        WHERE fr.survey_schema_id IN (26,33)
        """

        return pd.read_sql(query, engine)


#  Ajouter le préfixe 00216 aux numéros de téléphone
def add_prefix_to_phone_numbers(df):
    phone_columns = ['phone number', 'phone number 2']
    
    for col in phone_columns:
        if col in df.columns:
            
            def format_phone_number(phone):
                if pd.isna(phone) or phone == '' or phone is None:
                    return phone
                
                # Convert to string and keep only digits
                phone_str = ''.join(filter(str.isdigit, str(phone).strip()))
                
                # 👉 ADD PREFIX ONLY IF LENGTH = 8
                if len(phone_str) == 8:
                    return '00216' + phone_str
                
                # 👉 OTHERWISE DO NOTHING
                return phone_str
            
            df[col] = df[col].apply(format_phone_number)
            
    
    return df

# ===================== MAIN =====================
try:
    print(" Réclamations → SharePoint")
    headers = authenticate_sharepoint()
    drive_id = get_drive_id(headers)
    # Step 1: Get all reclamations from database
    df_raw = connect_and_query()

    def safe_json(x):
        try:
            return json.loads(x) if isinstance(x, str) else {}
        except:
            return {}

    df_data = df_raw['data'].apply(safe_json).apply(pd.Series)
    df_resp = df_raw['response_data'].apply(safe_json).apply(pd.Series)

    df = pd.concat(
        [df_raw.drop(columns=['data','response_data']), df_data, df_resp],
        axis=1
    )
   
    # ===== AGENCE NORMALIZATION =====
    agence_variants = ['Agence', 'AGNECE', 'agancey']
    existing_agence_cols = [c for c in agence_variants if c in df.columns]

    if existing_agence_cols:
        df['Agence'] = df[existing_agence_cols].bfill(axis=1).iloc[:, 0]
        df.drop(columns=[c for c in existing_agence_cols if c != 'Agence'], inplace=True)

    # ===== FILTER RECLAMATIONS =====
    df = df[df['Réclamation'].astype(str).str.upper() == 'OUI']

    # ===== SHAREPOINT EXISTING =====
    
    existing_df = read_excel_from_sharepoint(headers, drive_id)

    # Store total count BEFORE filtering for new records
    total_count = len(existing_df) if not existing_df.empty else 0
    
    if not existing_df.empty:
        max_id = existing_df['survey_data_id'].max()
        df_new = df[df['survey_data_id'] > max_id]
        write_log(f" Max existing ID: {max_id}")
        write_log(f" New records found: {len(df_new)}")
        
        # Update total count with new records
        if not df_new.empty:
            total_count += len(df_new)
    else:
        df_new = df
        total_count = len(df_new)
        write_log(f" No existing file, all records are new: {total_count}")

    if df_new.empty:
        write_log("ℹ️ No new reclamations to add")
        # We still want to send email with existing data
        final_df = existing_df.copy() if not existing_df.empty else pd.DataFrame()
    else:
        # Process new records
        if 'AGENCE' in df_new.columns:
            df_new = df_new.rename(columns={'AGENCE': 'Agence'})
            write_log("Colonne 'AGENCE' renommée en 'Agence'")
        elif 'agence' in df_new.columns:
            df_new = df_new.rename(columns={'agence': 'Agence'})
            write_log("Colonne 'agence' renommée en 'Agence'")
        
        # ===== KEEP ONLY REQUIRED COLUMNS =====
        final_columns = [
            "survey_data_id","IDu","survey_schema_id","response_created","Agence","Garant",
            "Code Client","Nom complet","Numéro CIN","phone number","phone number 2",
            "Charge Credit","Type Garantie","Adresse Projet","Duree aprouvee",
            "Montant aprouve","Objet du crédit","Adresse Personnelle","Numero credit",
            "Reference IDC","Qualification","Sous qualification","Comment",
            "Adresse professionnelle","Correction de l'adresse professionnelle",
            "Client bénéficiaire du crédit","Réclamation","Si réclamation","Availability",
            "Nom du bénéficiaire","Numéro de téléphone du bénéficiaire",
            "Autre numéro de téléphone","Adresse personnelle","Adresse sur CIN"
        ]
        
        # Process Qualification
        if 'Qualification' in df_new.columns:
            qualification_mapping = {
                'Client ne veut pas parler': 'Client ne veut pas parler : Red Flag',
                'Client se désiste': 'Client se désiste : Red Flag'
            }
            df_new['Qualification'] = df_new['Qualification'].replace(qualification_mapping)
        
        # Process Availability
        if 'Availability' in df_new.columns:
            if 'Qualification' not in df_new.columns:
                df_new['Qualification'] = None
            
            conditions = {
                'wrongnumber': 'Faux Numéro : Red Flag',
                'offtarget': 'Faux Numéro : Red Flag',
                'unavailable': 'Injoignable : Red Flag'
            }
            
            df_new['Qualification'] = df_new.apply(
                lambda row: conditions.get(
                    str(row['Availability']).lower(),
                    row['Qualification']
                ),
                axis=1
            )
        
        # Split Qualification
        if 'Qualification' in df_new.columns:
            df_new['Sous qualification'] = None
            
            split_mapping = {
                'Ecart dossier détecté : Orange Flag': ('Orange Flag', 'Ecart dossier détecté'),
                'Aucun écart détecté : Green Flag': ('Green Flag', None),
                'Injoignable : Red Flag': ('Red Flag', 'Injoignable'),
                'Client ne veut pas parler : Red Flag': ('Red Flag', 'Client ne veut pas parler'),
                'Client se désiste : Red Flag': ('Red Flag', 'Client se désiste'),
                'Faux Numéro : Red Flag': ('Red Flag', 'Faux Numéro'),
                'Hors Cible : Red Flag': ('Red Flag', "N'a pas demandé de crédit"),
                'Black Flag': ('Black Flag', None)
            }
            
            def split_qualification(val):
                if val in split_mapping:
                    return split_mapping[val]
                return val, None
            
            df_new[['Qualification', 'Sous qualification']] = (
                df_new['Qualification']
                .apply(split_qualification)
                .apply(pd.Series)
            )
        
        # Keep only required columns
        df_new = df_new.reindex(columns=final_columns)
        
        # ===== MERGE & UPLOAD =====
        if not existing_df.empty:
            final_df = pd.concat([existing_df, df_new], ignore_index=True)
        else:
            final_df = df_new
        

        # Appliquer le préfixe 00216 aux numéros de téléphone
        if not final_df.empty:
            final_df = add_prefix_to_phone_numbers(final_df)
            write_log("✅ Préfixe 00216 ajouté aux numéros de téléphone")

        # Upload updated file
        temp_file = os.path.join(os.getenv("TEMP"), "Reclamation___.xlsx")
        final_df.to_excel(temp_file, index=False)
        upload_excel_to_sharepoint(headers, drive_id, temp_file)
        os.remove(temp_file)
        
        write_log(f" SharePoint updated ({len(df_new)} new rows, total: {total_count})")

    # ===== SEND EMAIL =====
    write_log(f" Preparing reclamations email (Total: {total_count})")
    
    today = datetime.now().date()
    
    # Download final SharePoint file for attachment
    temp_dir = os.path.join(os.getenv("TEMP"), "reclamation_mail")
    os.makedirs(temp_dir, exist_ok=True)
    attachment_path = os.path.join(temp_dir, f"Reclamations_{today}.xlsx")
    
    download_sharepoint_file(
        headers,
        drive_id,
        SHAREPOINT_FILE_PATH,
        attachment_path
    )
    
    # Encode logo
    image_path = r"C:\Users\Administrateur\Desktop\Daam\DAAM_Post_Prod\logo-complet-color.png"
    with Image.open(image_path) as img:
        img = img.resize((90, int(img.height * 90 / img.width)), Image.Resampling.LANCZOS)
        buffer = BytesIO()
        img.save(buffer, format="PNG")
        encoded_image = base64.b64encode(buffer.getvalue()).decode("utf-8")
    
    # ---- HTML BODY ----
    html_body = f"""
    <html>
    <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 0;">
        <div style="max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #ddd; border-radius: 8px; background-color: #f9f9f9; box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1); position: relative;">
            <div style="position: absolute; top: 0px; left: 0px; padding: 5px;">
                <img src="data:image/png;base64,{encoded_image}" alt="Logo" style="height: 30px; width: auto; object-fit: contain;">
            </div>
            <div style="margin-top: 60px;">
                <p style="font-size: 16px;">Bonjour,</p>
                <p style="font-size: 16px;">
                    Veuillez trouver ci-joint le rapport consolidé des réclamations CAF
                    à la date du <strong>{today}</strong>.
                </p>

                <ul>
                    <li>Total réclamations : <strong>{total_count}</strong></li>
                </ul>

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
        </div>
    </body>
    </html>
    """
    
    # ---- OUTLOOK ----
    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    
    mail.Subject = f"Réclamations CAF - {today}"
    #mail.To = "ix41p@ningen-group.com"
    # Uncomment for production:
    mail.To = "amine.cherni@daam.tn;Amani.Bouali@daam.tn;"
    mail.CC = "Ningen-Data-Management@ningen-group.com;ci68t@ningen-group.com;iz55x@ningen-group.com;cl37t@ningen-group.com;pw39f@ningen-group.com"
    mail.HTMLBody = html_body
    mail.Attachments.Add(attachment_path)
    
    mail.Send()
    time.sleep(3)
    
    # Cleanup
    shutil.rmtree(temp_dir, ignore_errors=True)
    
    write_log(" Email sent successfully!")
    

except Exception as e:
    
    write_log(f"❌ Erreur : {str(e)}")
    import traceback
    write_log(traceback.format_exc())
finally:
    gc.collect()