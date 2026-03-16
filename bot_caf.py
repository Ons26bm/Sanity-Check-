# # # -*- coding: utf-8 -*-
# # """
# # Created on Fri Aug 15 18:31:17 2025

# # @author: RZ35N
# # """
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
env_path = r"C:\Users\Administrateur\Desktop\Daam\DAAM_CAF\autoreport_caf_14h\.env"
load_dotenv(dotenv_path=env_path)

# Détails de la connexion SSH
ssh_host = os.getenv("SSH_HOST")
ssh_port = int(os.getenv("SSH_PORT", "22"))
ssh_user = os.getenv("SSH_USER")
ssh_password = os.getenv("SSH_PASSWORD")

# Détails de la base de données
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
log_filename = f"caf_conso_{datetime.now().strftime('%Y%m%d_%H.%M')}.txt"

# Bases de données à tester
databases = ['besidedb']





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

headers = authenticate_sharepoint()
drive_id = get_drive_id(headers)

def decode_unicode_escapes(text):
    """Décode les séquences d'échappement Unicode comme u00e9 en é"""
    if not isinstance(text, str):
        return text
    
    def replace_unicode(match):
        try:
            return chr(int(match.group(1), 16))
        except:
            return match.group(0)
    
    # Pattern pour les séquences Unicode de type u00e9
    pattern = r'u([0-9a-fA-F]{4})'
    return re.sub(pattern, replace_unicode, text)

def decode_unicode_sequences(text):
    """Décode les séquences Unicode plus complexes"""
    if not isinstance(text, str):
        return text
    
    # Remplacer les séquences Unicode courantes
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
    
    # Décoder les séquences Unicode
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
        write_log(f"Tunnel SSH établi sur le port local : {tunnel.local_bind_port}")

        for db_name in databases:
            write_log(f"\nTest de connexion à la base de données : {db_name}")
            engine = create_engine(
                f"mysql+pymysql://{db_user}:{db_password}@127.0.0.1:{tunnel.local_bind_port}/{db_name}"
            )

            try:
                with engine.connect() as connection:
                    write_log(f"Connexion réussie à la base {db_name}.")

                    # Exemple de requête spécifique
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
      AND (fr.survey_schema_id = 26 or fr.survey_schema_id = 33 )
WHERE (sd.survey_schema_id = 26 or sd.survey_schema_id = 33 ) 
  AND fr.response_data->'$.Availability' IN ('available','unavailable','wrongnumber','offtarget')  
 
ORDER BY response_created ASC;"""
                        df = pd.read_sql(query, con=connection)

                        # Exporter les données au format JSON
                        output_file = f"besidedb_myapp_caf.json"
                        df.to_json(output_file, orient="records", lines=True, force_ascii=False)
                        write_log(f"Données exportées dans le fichier JSON : {output_file}")
                        return df

            except Exception as e:
                write_log(f"Erreur lors de l'accès à {db_name} :", e)
                return None

# Exécuter la connexion et récupérer les données
df_original = connect_and_query()

#  Ajouter le préfixe 00216 aux numéros de téléphone
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
# ===================== MAIN =====================



# Si des données ont été récupérées, continuer le traitement
if df_original is not None:
    today = datetime.now().date()
    output_file = rf"C:\Users\Administrateur\Desktop\Daam\DAAM_CAF\autoreport_caf_14h\post_prod_caf\caf_{today}__.xlsx"

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

    # Désérialiser les colonnes JSON en gérant les valeurs None
    df_data = df_original['data'].apply(safe_json_loads).apply(pd.Series)
    df_response = df_original['response_data'].apply(safe_json_loads).apply(pd.Series)

    # Convertir les timestamps en dates lisibles
    if 'survey_created' in df_original.columns:
        df_original['survey_created'] = pd.to_datetime(df_original['survey_created'])
    if 'response_created' in df_original.columns:
        df_original['response_created'] = pd.to_datetime(df_original['response_created'])

    # Combiner toutes les colonnes désérialisées avec le reste
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
    df_final.columns = [clean_column_name(col) for col in df_final.columns]
    
    


    ########################### new code ###############################

    # df_final = pd.concat([df_final.iloc[:, :63], df_final.iloc[:, -1:]],axis=1)
    
    # df_final = df_final.drop(columns=['Feedback','Est-ce que votre chargé de crédit vous a demander des paiements autre que les frais stipulés dans la liste officielle ?','Préciser la raison avancée par le CC','De combien était le montant ?'
    #                                   ,'Et vous les avez payés ?','Autre numéro de téléphone','Vous êtes d’accord sur tous les conditions relatives à votre  crédit, vous voulez qu’on poursuive les démarches pour le déboursement ?','Avez-vous un commentaire sur votre expérience avec DAAM ou sur la relation avec votre chargé de crédit ?','On vous a tout expliquer concernant les frais associés à votre octroi de crédit ?',
    #                                   'Correction garantie','Et vous avez ……. Comme garantie, c’est bien ça ?','Et votre crédit, vous le demander pour financer quoi ?',"Quelle est l'activité de votre projet ?","Correction de l'adresse du projet","Et l’adresse de votre projet ?", "Correction de l'adresse personnelle" ,'Quelle est votre adresse personnelle actuelle ?',
    #                                   'Correction du garant',"C’est qui le garant de votre crédit ?",'Numéro de téléphone du bénéficiaire','Et vous êtes vous-même le bénéficiaire de ce crédit ?','Correction nom','Et votre nom complet ?','Correction CIN',
    #                                   'Pouvez vous me communiquer le numéro de votre carte d’indentité ?','2éme numéro','Vous avez un 2éme numéro de téléphone ?','Autre numéro',"C’est bien celui là le numéro sur lequel vous êtes joignable ?",'Vous avez été contacté par mon collégue pour vous communiquer ses informations ?','Nom du bénéficiaire',
    #                                   'Comment CAF','Survey ID CAF','Qualification CAF','Date Traitement CAF','Sous qualification CAF'
    #                                   ], errors='ignore')
    
    df_final = df_final.drop(columns=['Comment CAF','Survey ID CAF','Qualification CAF','Date Traitement CAF','Sous qualification CAF','Feedback'], errors='ignore')
    
    
    if 'Qualification' in df_final.columns:

        qualification_mapping = {
            'Client ne veut pas parler': 'Client ne veut pas parler : Red Flag',
            'Client se désiste': 'Client se désiste : Red Flag'
        }

        df_final['Qualification'] = (
            df_final['Qualification']
            .replace(qualification_mapping)
        )

       
    
    
    
    
    # Remplir la colonne Qualification selon Availability
    if 'Availability' in df_final.columns:

        # Créer la colonne si elle n'existe pas
        if 'Qualification' not in df_final.columns:
            df_final['Qualification'] = None

        conditions = {
            'wrongnumber': 'Faux Numéro : Red Flag',
            'offtarget': 'Faux Numéro : Red Flag',
            'unavailable': 'Injoignable : Red Flag'
        }

        df_final['Qualification'] = df_final.apply(
            lambda row: conditions.get(
                str(row['Availability']).lower(),
                row['Qualification']
            ),
            axis=1
        )

        
    
    
    # Création et remplissage de la colonne Actions selon Qualification
    if 'Qualification' in df_final.columns:

        action_mapping = {
            'Ecart dossier détecté : Orange Flag': 'Poursuivre le processus, remise de contrat autorisée',
            'Aucun écart détecté : Green Flag': 'Poursuivre le processus, remise de contrat autorisée',
            'Injoignable : Red Flag': "Prise de décision chef d'agence",
            'Client ne veut pas parler : Red Flag': "Prise de décision chef d'agence",
            'Client se désiste : Red Flag': "Prise de décision chef d'agence",
            'Faux Numéro : Red Flag': "Prise de décision chef d'agence",
            'Hors Cible : Red Flag': "Prise de décision chef d'agence",
            'Black Flag': 'Remise de contrat INTERDITE en attente autorisation claire et explicite du responsable de conformité'
        }

        df_final['Actions'] = df_final['Qualification'].map(action_mapping)

        # Placer la colonne Actions juste après Qualification
        cols = list(df_final.columns)
        qual_index = cols.index('Qualification')
        cols.insert(qual_index + 1, cols.pop(cols.index('Actions')))
        df_final = df_final[cols]

        
        
        
        # Séparer Qualification en "Qualification" (Flag) et "Sous qualification"
        if 'Qualification' in df_final.columns:

            # Créer la colonne Sous qualification
            df_final['Sous qualification'] = None

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

            df_final[['Qualification', 'Sous qualification']] = (
                df_final['Qualification']
                .apply(split_qualification)
                .apply(pd.Series)
            )

            # Réordonner les colonnes :
            cols = list(df_final.columns)

            # Retirer si déjà présentes
            for c in ['Sous qualification', 'Actions']:
                if c in cols:
                    cols.remove(c)

            qual_idx = cols.index('Qualification')
            cols.insert(qual_idx + 1, 'Sous qualification')
            cols.insert(qual_idx + 2, 'Actions')

            df_final = df_final[cols]

           
        
    # ===================== FILTRE MOIS COURANT =====================

    # S'assurer que response_created est bien en datetime
    df_final['response_created'] = pd.to_datetime(
        df_final['response_created'], errors='coerce'
    )

    # Sauvegarder le dataframe complet AVANT filtre (pour Réclamations)
    df_final_all_months = df_final.copy()

    # Début et fin du mois courant
    start_of_month = pd.Timestamp.today().replace(day=1).normalize()
    end_of_month = start_of_month + pd.offsets.MonthEnd(1)

    # Filtrer uniquement les données du mois courant
    df_final = df_final[
        (df_final['response_created'] >= start_of_month) &
        (df_final['response_created'] <= end_of_month)
    ]


    df_final = df_final.loc[:, ~df_final.columns.duplicated()]

    # Appliquer le préfixe 00216 aux numéros de téléphone
    if not df_final.empty:
            final_data = add_prefix_to_phone_numbers(df_final)
            
    ########################### end  new code ###############################
    # Corriger l'erreur ExcelWriter
    try:
        # Méthode 1 : Sans options
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Data')
    except:
        # Méthode 2 : Avec xlsxwriter si openpyxl ne fonctionne pas
        try:
            with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Data')
        except:
            # Méthode 3 : Méthode simple
            df_final.to_excel(output_file, index=False)
    
    write_log(f"Fichier Excel créé : {output_file}")
    
    
else:
    write_log("Aucune donnée n'a été récupérée de la base de données.")











# # # #*****************************************PART 2********************************************************

import os
import base64
import time
import gc
import pythoncom
from datetime import date, timedelta
from pywinauto import Application
from PIL import Image
from io import BytesIO
import win32com.client as win32  # Ajout nécessaire si absent
from datetime import datetime

try:
    # --- Données ---
    today = datetime.now().date()



    temp_excel_path = rf"C:\Users\Administrateur\Desktop\Daam\DAAM_CAF\autoreport_caf_14h\post_prod_caf\caf_{today}__.xlsx"
   


    # --- Encode le logo ---
    image_path = r"C:\Users\Administrateur\Desktop\Daam\DAAM_Post_Prod\logo-complet-color.png"
    with Image.open(image_path) as img:
        img = img.resize((90, int(img.height * 90 / img.width)), Image.Resampling.LANCZOS)
        buffer = BytesIO()
        img.save(buffer, format="PNG")
        encoded_image = base64.b64encode(buffer.getvalue()).decode("utf-8")

    # --- Corps HTML du mail ---
    html_body = f'''
    <html>
    <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 0;">
        <div style="max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #ddd; border-radius: 8px; background-color: #f9f9f9; box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1); position: relative;">
            <div style="position: absolute; top: 0px; left: 0px; padding: 5px;">
                <img src="data:image/png;base64,{encoded_image}" alt="Logo" style="height: 30px; width: auto; object-fit: contain;">
            </div>
            <div style="margin-top: 60px;">
                <p style="font-size: 16px;">Bonjour,</p>
                <p style="font-size: 16px;">Vous trouverez ci-joint la consolidation du traitement CAF.</p>
                
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
    '''

    # --- Initialisation Outlook ---
    pythoncom.CoInitialize()
    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.Subject = f'Consolidation CAF au {today}'
    #mail.To = "ix41p@ningen-group.com "
    
    mail.To = "aymen.ouertani@daam.tn;amine.cherni@daam.tn;"
    mail.CC = "hamdi.ouerfelli@daam.tn;hichem.trigui@daam.tn;Yesmine.CHARFI@daam.tn;Khouloud.MAIZA@daam.tn;maha.elaroui@daam.tn;Bilel.machat@daam.tn;ali.bouasida@daam.tn;ali.bouasida@daam.tn;mehdi.chaabouni@daam.tn;ayoub.romdhane@daam.tn;belhsan.chebbi@daam.tn;mohamed.karray@daam.tn;mohamed.abida@daam.tn;Mohamed.essafi@daam.tn;Conformite@daam.tn;Conformite_Juridique@daam.tn;Equipe_formation_developpement@daam.tn;fatma.elfouzi@daam.tn;Zeineb.felli@daam.tn;ines.kahlaoui@daam.tn;Ningen-pperformance@ningen-group.com;Ningen-Data-Management@ningen-group.com;iz55x@ningen-group.com;ci68t@ningen-group.com;pw39f@ningen-group.com"
    mail.HTMLBody = html_body
    mail.Attachments.Add(temp_excel_path)
    mail.Send()

    # --- Attente d'ouverture d'Outlook ---
    time.sleep(5)

    # --- Connexion avec pywinauto et envoi automatique ---
    try:
        app = Application(backend="uia").connect(title_re=f".*Consolidation CAF au {today}.*")
        window = app.window(title_re=f".*Consolidation CAF au {today}.*")
        send_button = window.child_window(title="Envoyer", control_type="Button")
        send_button.wait("enabled", timeout=10)
        send_button.click_input()
        print("Email envoyé avec succès via Outlook.")
    except Exception as send_error:
        print(f" Impossible d'envoyer l'e-mail automatiquement : {send_error}")
        

    
    pythoncom.CoUninitialize()
    gc.collect()

except Exception as e:
    write_log(f"Erreur lors de la préparation ou de l'envoi de l'e-mail : {str(e)}")
    write_log(traceback.format_exc())










# #****************************************PART 3*********************************************************
today = datetime.now().date()
# Configuration du logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def nettoyer_colonne_agence(df):
    """Nettoie la colonne Agence en supprimant les codes numériques et espaces superflus"""
    if 'Agence' in df.columns:
        logging.info("Nettoyage de la colonne Agence")
        
        # Convertir en string pour éviter les erreurs
        df['Agence'] = df['Agence'].astype(str)
        
        # 1. Supprimer les codes numériques au début (format "00011-")
        df['Agence'] = df['Agence'].str.replace(r'^\d+-', '', regex=True)
        
        # 2. Supprimer les espaces superflus au début et à la fin
        df['Agence'] = df['Agence'].str.strip()
        
        # 3. Remplacer les espaces multiples par un seul espace
        df['Agence'] = df['Agence'].str.replace(r'\s+', ' ', regex=True)
        
        logging.info(f"Exemple de nettoyage: {df['Agence'].iloc[0] if len(df) > 0 else 'Aucune donnée'}")
    else:
        logging.warning("La colonne 'Agence' n'existe pas dans le DataFrame")
        write_log("La colonne 'Agence' n'existe pas dans le DataFrame")
    
    return df

def proteger_excel_avec_mot_de_passe(chemin_fichier, mot_de_passe):
    """Protège un fichier Excel avec un mot de passe en utilisant win32com."""
    excel = None
    workbook = None
    try:
        # Initialiser COM en mode single-threaded
        pythoncom.CoInitialize()
        
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.DisplayAlerts = False
        excel.Visible = False
        excel.EnableEvents = False
        
        # Ouvrir le workbook
        workbook = excel.Workbooks.Open(chemin_fichier)
        
        # Protéger le workbook avec mot de passe
        workbook.Password = mot_de_passe
        
        # Sauvegarder et fermer
        workbook.Save()
        workbook.Close(False)
        
        logging.info(f"Fichier Excel protégé avec succès: {chemin_fichier}")
        return True
        
    except Exception as e:
        logging.error(f"Erreur lors de la protection du fichier Excel {chemin_fichier}: {e}")
        write_log(f"Erreur lors de la protection du fichier Excel {chemin_fichier}: {str(e)}")
        
        return False
        
    finally:
        # Nettoyer les objets COM proprement
        try:
            if workbook:
                workbook.Close(False)
                workbook = None
            if excel:
                excel.Quit()
                excel = None
        except:
            pass
        finally:
            pythoncom.CoUninitialize()

def creer_zip_protege(chemin_fichier, mot_de_passe_zip):
    """Crée une archive ZIP protégée par mot de passe."""
    try:
        # Créer un buffer en mémoire pour le ZIP
        zip_buffer = BytesIO()
        
        # Créer le ZIP avec mot de passe
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
            zipf.setpassword(mot_de_passe_zip.encode('utf-8'))
            zipf.write(chemin_fichier, os.path.basename(chemin_fichier))
        
        # Sauvegarder le ZIP sur disque
        zip_path = f"{chemin_fichier}.zip"
        with open(zip_path, 'wb') as f:
            f.write(zip_buffer.getvalue())
        
        logging.info(f"Archive ZIP créée avec succès: {zip_path}")
        return zip_path
        
    except Exception as e:
        logging.error(f"Erreur lors de la création du ZIP: {e}")
        write_log(f"Erreur lors de la création du ZIP: {str(e)}")

        return None

def envoyer_email(destinataire, cc, sujet, corps_html, piece_jointe, logo_path=None):
    """Envoie un email avec pièce jointe et marque-le comme Confidentiel"""
    outlook = None
    try:
        # Initialiser COM pour Outlook
        pythoncom.CoInitialize()
        
        outlook = win32com.client.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        
        # Marquer comme confidentiel
        mail.Sensitivity = 3  # 0=Normal, 1=Personnel, 2=Privé, 3=Confidentiel
        mail.Subject = sujet
        mail.HTMLBody = corps_html
        mail.Attachments.Add(piece_jointe)
        
        # Ajouter le logo en pièce jointe intégrée si le chemin est fourni
        if logo_path and os.path.exists(logo_path):
            logo_attachment = mail.Attachments.Add(logo_path)
            logo_attachment.PropertyAccessor.SetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x3712001F", 
                "logo1.png"
            )
            logging.info("Logo ajouté en pièce jointe intégrée")
        
        # Ajouter les destinataires
        if destinataire:
            mail.To = destinataire
        if cc:
            mail.CC = cc
        
        # Essayer de définir l'étiquette avancée
        try:
            mail.PropertyAccessor.SetProperty(
                "http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/keywords",
                "Confidentiel"
            )
        except Exception as e:
            logging.warning(f"Impossible de définir l'étiquette avancée: {e}")
            write_log(f"Impossible de définir l'étiquette avancée: {str(e)}")
        
        # Envoyer l'email
        mail.Send()
        logging.info(f"Email envoyé avec succès à: {destinataire}")
        return True
        
    except Exception as e:
        logging.error(f"Erreur lors de l'envoi de l'email: {e}")
        write_log(f"Erreur lors de l'envoi de l'email: {str(e)}")
        return False
        
    finally:
        # Nettoyer les objets COM
        try:
            if outlook:
                outlook = None
        except:
            pass
        finally:
            pythoncom.CoUninitialize()

def creer_corps_email(nombre_fiches, agence_nom):
    """Crée le corps HTML de l'email avec le design"""
    LOGO_PATH = r"C:\Users\Administrateur\Desktop\Daam\DAAM_CAF\autoreport_caf_14h\logo-complet-color.png"
    
    # Vérifier si le logo existe
    logo_exists = os.path.exists(LOGO_PATH)
    
    corps_email = f"""
<html>
<head>
    <meta charset="UTF-8">
</head>
<body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 0;">
    <div style="max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #ddd; border-radius: 8px; background-color: #f9f9f9; box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);">
        <div style="position: relative;">
            {"<div style='position: absolute; top: 0; right: 0;'><img src='cid:logo1.png' alt='Logo' style='height: 50px;'></div>" if logo_exists else ""}
            <div style="margin-top: 20px;">
                <p style="font-size: 16px;">Bonjour,</p>
                <p style="font-size: 16px;">Veuillez trouver en pièce jointe les fiches CAF pour votre agence.</p>

                <ul style="list-style-type: none; padding: 0;">
                    <li style="margin-bottom: 10px; padding: 10px; background-color: #ffffff; border-radius: 4px; border: 1px solid #ddd;">
                        <span style="font-weight: bold; color: #004986; font-size: 14px;">Agence :</span>
                        <span style="font-size: 14px;"> {agence_nom}</span>
                    </li>
                    <li style="margin-bottom: 10px; padding: 10px; background-color: #ffffff; border-radius: 4px; border: 1px solid #ddd;">
                        <span style="font-weight: bold; color: #004986; font-size: 14px;">Nombre de fiches :</span>
                        <span style="font-size: 14px;"> {nombre_fiches} Fiche{'s' if nombre_fiches > 1 else ''}</span>
                    </li>
                    <li style="margin-bottom: 10px; padding: 10px; background-color: #ffffff; border-radius: 4px; border: 1px solid #ddd;">
                        <span style="font-weight: bold; color: #004986; font-size: 14px;">Mot de passe :</span>
                        <span style="font-size: 14px;"> DAAM@2025</span>
                    </li>
                </ul>
            </div>
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
                Ce message et les éventuelles pièces jointes sont strictement confidentiels et destinés exclusivement
                au(x) destinataire(s) indiqué(s). Toute utilisation, diffusion ou reproduction non autorisée est
                interdite. Si vous avez reçu ce message par erreur, merci d'en avertir immédiatement l'expéditeur et de
                supprimer le courriel.
            </p>
        </div>
    </div>
</body>
</html>
    """
    return corps_email

def traiter_et_envoyer_agence(agency_name, agency_df, df_annuaire, temp_dir, excel_password, zip_password):
    """Traite et envoie les emails pour une agence spécifique"""
    try:
        logging.info(f"Traitement de l'agence normale: {agency_name}")
        
        # Créer le fichier Excel pour cette agence
        agency_file = os.path.join(temp_dir, f"CAF_{agency_name.replace(' ', '')}.xlsx")
        agency_df.to_excel(agency_file, index=False)
        logging.info(f"Fichier Excel créé: {agency_file}")
        
        # Protéger le fichier Excel avec mot de passe
        if proteger_excel_avec_mot_de_passe(agency_file, excel_password):
            # Créer un ZIP protégé
            zip_file = creer_zip_protege(agency_file, zip_password)
            
            if not zip_file:
                logging.error(f"Échec de la création du ZIP pour {agency_name}")
                write_log(f"Échec de la création du ZIP pour {agency_name}")
                try:
                    os.remove(agency_file)
                except Exception as e:
                    logging.warning(f"Impossible de supprimer le fichier Excel {agency_file}: {e}")
                    write_log(f"Impossible de supprimer le fichier Excel {agency_file}: {str(e)}")
                return False
            
            # Trouver les destinataires dans l'annuaire pour l'agence normale
            agency_contacts = df_annuaire[
                df_annuaire['Destination'].str.strip().str.upper() == agency_name.strip().upper()
            ]
            
            # Préparer les listes d'emails
            to_emails = agency_contacts[
                (agency_contacts['Mail'] == 'destinataire') & 
                (agency_contacts['courriel Externe'].notnull())
            ]['courriel Externe'].tolist()
            
            cc_emails = agency_contacts[
                (agency_contacts['Mail'] == 'En copie') & 
                (agency_contacts['courriel Externe'].notnull())
            ]['courriel Externe'].tolist()
            
            # Vérifier qu'on a au moins un destinataire ou copie
            if not to_emails and not cc_emails:
                logging.warning(f"Aucun email trouvé pour {agency_name} - Fichier non envoyé")
                write_log(f"Aucun email trouvé pour {agency_name} - Fichier non envoyé")
                try:
                    os.remove(zip_file)
                    os.remove(agency_file)
                except Exception as e:
                    logging.warning(f"Impossible de supprimer les fichiers: {e}")
                    write_log(f"Impossible de supprimer les fichiers: {str(e)}")
                return False
            
            # Préparer le corps de l'email
            nombre_fiches = len(agency_df)
            LOGO_PATH = r"C:\Users\Administrateur\Desktop\Daam\DAAM_CAF\autoreport_caf_14h\logo-complet-color.png"
            
            corps_email = creer_corps_email(nombre_fiches, agency_name)
            
            # Envoyer l'email avec le logo
            if envoyer_email(
                destinataire=";".join(to_emails) if to_emails else "",
                cc=";".join(cc_emails) if cc_emails else "",
                sujet=f"[CONFIDENTIEL] Fiches CAF - {agency_name} - {datetime.now().strftime('%d/%m/%Y')}",
                corps_html=corps_email,
                piece_jointe=zip_file,
                logo_path=LOGO_PATH
            ):
                logging.info(f"Email envoyé avec succès pour {agency_name}")
                success = True
            else:
                logging.error(f"Échec de l'envoi de l'email pour {agency_name}")
                write_log(f"Échec de l'envoi de l'email pour {agency_name}")
                success = False
            
            # Supprimer le fichier ZIP après envoi
            try:
                os.remove(zip_file)
            except Exception as e:
                logging.warning(f"Impossible de supprimer le fichier ZIP {zip_file}: {e}")
                write_log(f"Impossible de supprimer le fichier ZIP {zip_file}: {str(e)}")
        else:
            success = False
        
        # Supprimer le fichier Excel original
        try:
            os.remove(agency_file)
        except Exception as e:
            logging.warning(f"Impossible de supprimer le fichier Excel {agency_file}: {e}")
            write_log(f"Impossible de supprimer le fichier Excel {agency_file}: {str(e)}")
        
        return success
        
    except Exception as e:
        logging.error(f"Erreur lors du traitement de l'agence {agency_name}: {e}")
        write_log(f"Erreur lors du traitement de l'agence {agency_name}: {str(e)}")
        return False

def main():
    # Initialiser COM dans le thread principal
    pythoncom.CoInitialize()
    
    try:
        # Chemin vers le fichier Excel à charger
        excel_file_path = rf"C:\Users\Administrateur\Desktop\Daam\DAAM_CAF\autoreport_caf_14h\post_prod_caf\caf_{today}__.xlsx"
        
        # Mots de passe
        excel_password = "DAAM@2025"
        zip_password = "DAAM@2025"
        
        # Charger le fichier Excel
        logging.info(f"Chargement du fichier Excel: {excel_file_path}")
        df = pd.read_excel(excel_file_path)
        df = df.drop(columns=['Détails Black Flag'], errors='ignore')
        # Afficher les colonnes disponibles pour debug
        logging.info(f"Colonnes disponibles: {list(df.columns)}")
        
        # Nettoyer la colonne Agence si elle existe
        df = nettoyer_colonne_agence(df)
        
        # Vérifier si la colonne Agence existe avant de continuer
        if 'Agence' not in df.columns:
            logging.error("La colonne 'Agence' est introuvable dans le fichier. Arrêt du traitement.")
            write_log("La colonne 'Agence' est introuvable dans le fichier. Arrêt du traitement.")

            return
        
        # Lire le fichier annuaire des agences
        try:
            annuaire_path = r"C:\Users\Administrateur\Desktop\Daam\DAAM_CAF\autoreport_caf_14h\Annuaire conso agence CAF  FINAL  11122025.xlsx"
            #annuaire_path = r"C:\Users\Administrateur\Desktop\Daam\DAAM_CAF\autoreport_caf_14h\Annuaire CAF TEST.xlsx"
            logging.info(f"Chargement de l'annuaire: {annuaire_path}")
            df_annuaire = pd.read_excel(annuaire_path, sheet_name='Annuaire')
            
            # Nettoyer les noms de colonnes
            df_annuaire.columns = df_annuaire.columns.str.strip()
            
            # Créer un dossier temporaire pour les fichiers Excel par agence
            temp_dir = r"C:\Users\Administrateur\Desktop\Daam\DAAM_CAF\autoreport_caf_14h\temp_agences"
            os.makedirs(temp_dir, exist_ok=True)
            logging.info(f"Dossier temporaire créé: {temp_dir}")
            
            # Vérifier si la colonne Source existe pour séparer les données
            if 'Source' in df.columns:
                # Séparer les données : agences normales vs agence mobile
                df_mobile = df[df['Source'].str.lower() == 'agence mobile'] if 'Source' in df.columns else pd.DataFrame()
                df_normal = df[df['Source'].str.lower() != 'agence mobile'] if 'Source' in df.columns else df
            else:
                # Si pas de colonne Source, traiter tout comme agences normales
                logging.warning("Colonne 'Source' non trouvée, traitement de toutes les données comme agences normales")
                df_mobile = pd.DataFrame()
                df_normal = df
            
            # Traitement des agences normales (envoi par agence)
            if not df_normal.empty:
                grouped_normal = df_normal.groupby('Agence')
                
                for agency_name, agency_df in grouped_normal:
                    success = traiter_et_envoyer_agence(
                        agency_name, agency_df, df_annuaire, temp_dir, excel_password, zip_password
                    )
                    
                    # Petite pause entre les envois
                    if success:
                        time.sleep(3)
            
            # Traitement des agences mobile (regroupées dans un seul fichier)
            if not df_mobile.empty:
                logging.info("Traitement des fiches Agence Mobile")
                
                # Créer un seul fichier pour toutes les agences mobile
                mobile_file = os.path.join(temp_dir, "caf__AGENCE_MOBILE.xlsx")
                df_mobile.to_excel(mobile_file, index=False)
                logging.info(f"Fichier Excel Agence Mobile créé: {mobile_file}")
                
                # Protéger le fichier Excel avec mot de passe
                if proteger_excel_avec_mot_de_passe(mobile_file, excel_password):
                    # Créer un ZIP protégé
                    zip_file = creer_zip_protege(mobile_file, zip_password)
                    
                    if not zip_file:
                        logging.error("Échec de la création du ZIP pour Agence Mobile")
                        write_log("Échec de la création du ZIP pour Agence Mobile")
                        try:
                            os.remove(mobile_file)
                        except Exception as e:
                            logging.warning(f"Impossible de supprimer le fichier Excel {mobile_file}: {e}")
                            write_log(f"Impossible de supprimer le fichier Excel {mobile_file}: {str(e)}")
                    else:
                        # Trouver les destinataires dans l'annuaire pour AGENCE MOBILE
                        agency_contacts = df_annuaire[
                            df_annuaire['Destination'].str.strip().str.upper() == 'AGENCE MOBILE'
                        ]
                        
                        # Préparer les listes d'emails
                        to_emails = agency_contacts[
                            (agency_contacts['Mail'] == 'destinataire') & 
                            (agency_contacts['courriel Externe'].notnull())
                        ]['courriel Externe'].tolist()
                        
                        cc_emails = agency_contacts[
                            (agency_contacts['Mail'] == 'En copie') & 
                            (agency_contacts['courriel Externe'].notnull())
                        ]['courriel Externe'].tolist()
                        
                        # Vérifier qu'on a au moins un destinataire ou copie
                        if not to_emails and not cc_emails:
                            logging.warning("Aucun email trouvé pour AGENCE MOBILE - Fichier non envoyé")
                            write_log("Aucun email trouvé pour AGENCE MOBILE - Fichier non envoyé")
                            try:
                                os.remove(zip_file)
                                os.remove(mobile_file)
                            except Exception as e:
                                logging.warning(f"Impossible de supprimer les fichiers: {e}")
                                write_log(f"Impossible de supprimer les fichiers: {str(e)}")
                        else:
                            # Préparer le corps de l'email
                            nombre_fiches_mobile = len(df_mobile)
                            LOGO_PATH = r"C:\Users\Administrateur\Desktop\Daam\DAAM_CAF\autoreport_caf_14h\logo-complet-color.png"
                            
                            corps_email = creer_corps_email(nombre_fiches_mobile, "Agence Mobile")
                            
                            # Envoyer l'email
                            if envoyer_email(
                                destinataire=";".join(to_emails) if to_emails else "",
                                cc=";".join(cc_emails) if cc_emails else "",
                                sujet=f"[CONFIDENTIEL] Fiches CAF - Agence Mobile - {datetime.now().strftime('%d/%m/%Y')}",
                                corps_html=corps_email,
                                piece_jointe=zip_file,
                                logo_path=LOGO_PATH
                            ):
                                logging.info("Email envoyé avec succès pour Agence Mobile")
                            else:
                                logging.error("Échec de l'envoi de l'email pour Agence Mobile")
                                write_log("Échec de l'envoi de l'email pour Agence Mobile")
                            
                            # Supprimer le fichier ZIP après envoi
                            try:
                                os.remove(zip_file)
                            except Exception as e:
                                logging.warning(f"Impossible de supprimer le fichier ZIP {zip_file}: {e}")
                                write_log(f"Impossible de supprimer le fichier ZIP {zip_file}: {str(e)}")
                
                # Supprimer le fichier Excel original
                try:
                    os.remove(mobile_file)
                except Exception as e:
                    logging.warning(f"Impossible de supprimer le fichier Excel {mobile_file}: {e}")
                    write_log(f"Impossible de supprimer le fichier Excel {mobile_file}: {str(e)}")
            
            # Supprimer le dossier temporaire après envoi
            try:
                shutil.rmtree(temp_dir)
                logging.info(f"Dossier temporaire supprimé: {temp_dir}")
            except Exception as e:
                logging.warning(f"Impossible de supprimer le dossier temporaire {temp_dir}: {e}")
                write_log(f"Impossible de supprimer le dossier temporaire {temp_dir}: {str(e)}")
            
        except Exception as e:
            logging.error(f"Erreur lors de l'envoi des emails: {e}")
            write_log(f"Erreur lors de l'envoi des emails: {str(e)}")
            
    except Exception as e:
        logging.error(f"Erreur lors du traitement du fichier Excel: {e}")
        write_log(f"Erreur lors du traitement du fichier Excel: {str(e)}")
        
    finally:
        # Désinitialiser COM dans le thread principal
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    main()