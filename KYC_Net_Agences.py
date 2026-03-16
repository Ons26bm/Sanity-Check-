import pandas as pd
import os
import tempfile
import re
import time
import win32com.client as win32
from openpyxl import load_workbook
import shutil
import zipfile

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
import traceback
from dotenv import load_dotenv
env_path = r"C:\Users\Administrateur\Desktop\Daam\DAAM_C_KYC\C_KYC_Nets_Agences\.env"
load_dotenv(dotenv_path=env_path)

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
SITE_DOMAIN = os.getenv("SITE_DOMAIN")
SITE_NAME = os.getenv("SITE_NAME")
SHAREPOINT_FILE_LOG_PATH ="General/Autoreports Status"
# Log file name
log_filename = f"KYC_Net_Agences_{datetime.now().strftime('%Y%m%d_%H.%M')}.txt"




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
        write_log(f" Impossible de logger sur SharePoint : {e}")
 
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

def add_prefix_to_phone_numbers(df):
    import pandas as pd
    
    phone_columns = ['phone number', 'phone number 2']
    
    for col in phone_columns:
        if col in df.columns:
            
            def format_phone_number(phone):
                
                # If value is really empty → keep it empty
                if pd.isna(phone) or phone == '':
                    return ''
                
                phone_str = str(phone).strip()
                
                # Remove .0 caused by float conversion
                if phone_str.endswith('.0'):
                    phone_str = phone_str[:-2]
                
                # Keep only digits
                phone_str = ''.join(filter(str.isdigit, phone_str))
                
                # If after cleaning nothing remains → keep empty
                if phone_str == '':
                    return ''
                
                # Add prefix only if missing
                if not phone_str.startswith('00216'):
                    phone_str = '00216' + phone_str
                
                return phone_str
            
            df[col] = df[col].apply(format_phone_number)
    
    return df



headers = authenticate_sharepoint()
drive_id = get_drive_id(headers)
# --- Fonction pour protéger un fichier Excel ---
def proteger_excel_avec_mot_de_passe(chemin_fichier, mot_de_passe):
    """Protège le fichier Excel avec mot de passe pour l'ouverture et les feuilles"""
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.DisplayAlerts = False
        excel.Visible = False

        workbook = excel.Workbooks.Open(chemin_fichier)
        # Protection ouverture + modification
        workbook.SaveAs(chemin_fichier, Password=mot_de_passe, WriteResPassword=mot_de_passe)

        # Protéger chaque feuille
        for sheet in workbook.Worksheets:
            sheet.Protect(Password=mot_de_passe)

        workbook.Save()
        workbook.Close(True)
        excel.Quit()

        del workbook
        del excel

        write_log(f"Fichier Excel protégé avec mot de passe: {chemin_fichier}")
        return True
    except Exception as e:
        write_log(f"Erreur lors de la protection du fichier Excel : {str(e)}")
        return False

# --- Fonction pour créer un ZIP protégé ---
def create_password_protected_zip(input_file, output_zip, password):
    """Crée une archive ZIP protégée par mot de passe"""
    try:
        with zipfile.ZipFile(output_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
            zipf.setpassword(password.encode('utf-8'))
            zipf.write(input_file, os.path.basename(input_file), compress_type=zipfile.ZIP_DEFLATED)

        if os.path.exists(output_zip) and os.path.getsize(output_zip) > 0:
            print(f"Archive ZIP protégée créée avec succès: {output_zip}")
            return True
        else:
            write_log(f"Échec de création du ZIP: {output_zip}")
            return False
    except Exception as e:
        write_log(f"Erreur lors de la création du ZIP protégé: {str(e)}")
        return False

# --- Fonction principale d'envoi ---
def envoyer_emails_agences_html(df_filtre, fichier_annuaire, mot_de_passe="DAAM@2025"):
    try:
        df_annuaire = pd.read_excel(fichier_annuaire, sheet_name='Annuaire')
        df_annuaire.columns = df_annuaire.columns.str.strip()

        # Identifier colonnes agence / email
        col_agence = next((c for c in df_annuaire.columns if re.search(r'destination', c, re.I)), None)
        col_email = next((c for c in df_annuaire.columns if re.search(r'courriel|email|e-mail', c, re.I)), None)
        if not col_agence or not col_email:
            raise ValueError("Colonnes agence ou email introuvables dans l'annuaire.")

        # Créer dossier temporaire
        temp_dir = tempfile.mkdtemp()

        outlook = win32.DispatchEx('Outlook.Application')

        # Dictionnaires destinataires
        to_par_agence = {}
        cc_par_agence = {}

        # Collecter emails par agence
        for _, row in df_annuaire.iterrows():
            agence = row[col_agence]
            email = row[col_email]

            if pd.isna(agence) or pd.isna(email) or '@' not in str(email):
                continue

            if agence not in to_par_agence:
                to_par_agence[agence] = []
                cc_par_agence[agence] = []

            to_par_agence[agence].append(email)

        # Boucle sur chaque agence
        for agence in to_par_agence.keys():
            df_agence = df_filtre[df_filtre['Agence'] == agence]
            if df_agence.empty:
                write_log(f"Aucune donnée KYC pour l'agence {agence}")
                continue


            df_agence_email = df_agence.copy()
            df_agence_email = add_prefix_to_phone_numbers(df_agence_email)

            fichier_temp = os.path.join(temp_dir, f"KYC_{agence.replace(' ', '_')}.xlsx")
            df_agence_email.to_excel(fichier_temp, index=False)

            # Protection Excel
            proteger_excel_avec_mot_de_passe(fichier_temp, mot_de_passe)

            # Créer ZIP protégé
            zip_file = os.path.join(temp_dir, f"KYC_{agence.replace(' ', '_')}.zip")
            create_password_protected_zip(fichier_temp, zip_file, mot_de_passe)

            # Créer email
            mail = outlook.CreateItem(0)
            mail.Subject = f"Fiches KYC Conformes - {agence}"
            mail.Sensitivity = 3
            mail.To = ";".join(to_par_agence[agence])

            # HTML conservé
            corps_email = f"""
            <html>
            <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 0;">
                <div style="max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #ddd; border-radius: 8px; background-color: #f9f9f9; box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);">
                    <div style="margin-top: 20px;">
                        <p style="font-size: 16px;">Bonjour,</p>
                        <p style="font-size: 16px;">Veuillez trouver ci-joint les fiches KYC Conformes pour votre agence.</p>
                        <ul style="list-style-type: none; padding: 0;">
                            <li style="margin-bottom: 10px; padding: 10px; background-color: #ffffff; border-radius: 4px; border: 1px solid #ddd;">
                                <span style="font-weight: bold; color: #004986; font-size: 14px;">Agence :</span>
                                <span style="font-size: 14px;">{agence}</span>
                            </li>
                            <li style="margin-bottom: 10px; padding: 10px; background-color: #ffffff; border-radius: 4px; border: 1px solid #ddd;">
                                <span style="font-weight: bold; color: #004986; font-size: 14px;">Nombre de fiches :</span>
                                <span style="font-size: 14px;">{len(df_agence)} Fiches</span>
                            </li>
                        </ul>
                        

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
                            Ce message et les éventuelles pièces jointes sont strictement confidentiels et destinés exclusivement au(x) destinataire(s) indiqué(s). Toute utilisation, diffusion ou reproduction non autorisée est interdite. Si vous avez reçu ce message par erreur, merci d’en avertir immédiatement l’expéditeur et de supprimer le courriel.
                        </p>
                    </div>
                </div>
            </body>
            </html>
            """
            mail.HTMLBody = corps_email
            mail.Attachments.Add(zip_file)
            mail.Send()

            write_log(f"Email envoyé pour l'agence {agence} avec {len(df_agence)} lignes")

            time.sleep(2)

        shutil.rmtree(temp_dir)
        write_log("Tous les emails ont été envoyés.")

    except Exception as e:
        write_log(f"Erreur lors de l'envoi des emails: {str(e)}")
        if 'temp_dir' in locals() and os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)

# --- Exemple d'utilisation ---
fichier_kyc = r"C:\Users\Administrateur\Desktop\Daam\DAAM_C_KYC\C_KYC_Nets_Agences\KYC_Conforme_Net.xlsx"
fichier_annuaire = r"C:\Users\Administrateur\Desktop\Daam\DAAM_C_KYC\C_KYC_Nets_Agences\Annuaire agence C-KYC  11122025.xlsx"

df_filtre = pd.read_excel(fichier_kyc)
envoyer_emails_agences_html(df_filtre, fichier_annuaire)
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             