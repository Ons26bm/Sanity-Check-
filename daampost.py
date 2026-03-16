# import logging
# from sshtunnel import SSHTunnelForwarder
# import pymysql
# import mysql.connector
# import pandas as pd
# import json
# from datetime import datetime, timedelta
# import re
# import unicodedata
# import win32com.client as win32
# import io
# import os
# import warnings
# from datetime import datetime, timedelta
# from datetime import date, timedelta
# import base64
# warnings.filterwarnings('ignore')

# def normalize_json_columns(df, json_columns):
#     """
#     Normalise les colonnes JSON dans un DataFrame.

#     Arguments:
#         df (pd.DataFrame): Le DataFrame contenant les colonnes JSON.
#         json_columns (list): Une liste des noms de colonnes à normaliser.

#     Retourne:
#         pd.DataFrame: Un DataFrame avec les colonnes JSON aplaties.
#     """
#     for column in json_columns:
#         try:
#             # Convertir la colonne en JSON et normaliser
#             json_data = df[column].apply(lambda x: json.loads(x) if isinstance(x, str) else x)
#             normalized = pd.json_normalize(json_data)

#             # Renommer les colonnes pour éviter les collisions
#             normalized.columns = [f"{column}_{subcol}" for subcol in normalized.columns]

#             # Fusionner avec le DataFrame d'origine
#             df = pd.concat([df, normalized], axis=1).drop(columns=[column])
#         except Exception as e:
#             write_log(f"Erreur lors de la normalisation de la colonne {column}: {e}")

#     return df

# try:
#     conn = mysql.connector.connect(
#         host='dataserver.mysql.database.azure.com',
#         port=3306,
#         user='Admach',
#         password='NiGan10gEn',
#         database='datawarehouse',
#         ssl_ca='path_to_cert.pem'
#     )

#     if conn.is_connected():
#         write_log("Connexion réussie à la base de données Azure MySQL")
#         cursor = conn.cursor()

#         query1 = """
#         SELECT *
#         FROM datawarehouse.outbound_surveyresults AS rslt
#         INNER JOIN datawarehouse.outbound_surveydata AS inj ON rslt.surveydata_beside_id = inj.surveydata_beside_id where rslt.surveyschema_beside_id in (13,16,17,18,19,22,23,25,26,27,28,32,33,36)
#         and   rslt.created_at>'2025-12-01';
        
#         """
#             #(13,16,17,18,19,22,23,25,26,27,28,32,33)



#         """SELECT *
#         FROM datawarehouse.outbound_surveyresults AS rslt
#         INNER JOIN datawarehouse.outbound_surveydata AS inj ON rslt.surveydata_beside_id = inj.surveydata_beside_id where rslt.surveyschema_beside_id in (13,16,17,18,19,22,23,25,26,27,28,32,33)
#         and   rslt.created_at>'2025-12-01';;"""



#         cursor.execute(query1)
#         Data_ResponseForms = cursor.fetchall()
#         columns1 = [column[0] for column in cursor.description]
#         Data = pd.DataFrame(Data_ResponseForms, columns=columns1)

#         # Renommer les colonnes 'created_at' selon leur occurrence
#         created_at_indices = [i for i, col in enumerate(Data.columns) if col == 'created_at']
#         if len(created_at_indices) > 0:
#             Data.columns.values[created_at_indices[0]] = 'Date Traitement'
#         if len(created_at_indices) > 1:
#             Data.columns.values[created_at_indices[1]] = 'Date Injection'

#         # Normaliser les colonnes JSON
#         json_columns = ['response_data', 'data']  # Ajuste les noms si nécessaire
#         Data = normalize_json_columns(Data, json_columns)

#         # Exporter les données vers Excel
#         excel_file = r"C:\Users\Administrateur\Desktop\Daam\DAAM_Post_Prod\Post_Prod.xlsx"
#         import pandas as pd
#         from datetime import datetime

#         # Convertir les dates en format datetime
#         date_debut = datetime(2025, 6, 1)
#         date_aujourdhui = pd.to_datetime(datetime.today()) 

#         # Supposons que la colonne 'Date Traitement' est déjà au format datetime
#         # Si ce n'est pas le cas, il faut d'abord la convertir:
#         # Data['Date Traitement'] = pd.to_datetime(Data['Date Traitement'], dayfirst=True)

#         # Filtrer le dataframe
#         Data = Data[(Data['Date Traitement'] >= date_debut) & 
#                         (Data['Date Traitement'] <= date_aujourdhui)]
#         Data.to_excel(excel_file, index=False)
#         write_log(f"Les données ont été exportées avec succès vers {excel_file}")

# except mysql.connector.Error as e:
#     write_log(f"Erreur lors de la connexion à la base de données : {e}")

# finally:
#     if conn.is_connected():
#         cursor.close()
#         conn.close()
#         write_log("Connexion à la base de données fermée")




# import os
# import base64
# import time
# import gc
# import pythoncom
# from datetime import date, timedelta
# from pywinauto import Application
# from PIL import Image
# from io import BytesIO
# import win32com.client as win32  # Ajout nécessaire si absent

# try:
#     # --- Données ---
#     start_date = "01-06-2025"
#     end_date = (date.today() - timedelta(days=0)).strftime("%d-%m-%Y")

#     # --- Création fichier Excel temporaire ---
#     temp_folder = os.path.join(os.getenv('TEMP'), 'survey_attachments')
#     os.makedirs(temp_folder, exist_ok=True)
#     temp_excel_path = os.path.join(temp_folder, f"Extract_Post_Prod_{end_date}.xlsx")
#     Data.to_excel(temp_excel_path, sheet_name='Data_N', index=False)

#     if not os.path.exists(temp_excel_path):
#         raise Exception("Fichier Excel temporaire non créé.")
#     else:
#         write_log(f"Fichier Excel temporaire créé avec succès : {temp_excel_path}")

#     # --- Encode le logo ---
#     image_path = r"C:\Users\Administrateur\Desktop\Daam\DAAM_Post_Prod\logo-complet-color.png"
#     with Image.open(image_path) as img:
#         img = img.resize((90, int(img.height * 90 / img.width)), Image.Resampling.LANCZOS)
#         buffer = BytesIO()
#         img.save(buffer, format="PNG")
#         encoded_image = base64.b64encode(buffer.getvalue()).decode("utf-8")

#     # --- Corps HTML du mail ---
#     html_body = f'''
#     <html>
#     <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 0;">
#         <div style="max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #ddd; border-radius: 8px; background-color: #f9f9f9; box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1); position: relative;">
#             <div style="position: absolute; top: 0px; left: 0px; padding: 5px;">
#                 <img src="data:image/png;base64,{encoded_image}" alt="Logo" style="height: 30px; width: auto; object-fit: contain;">
#             </div>
#             <div style="margin-top: 60px;">
#                 <p style="font-size: 16px;">Bonjour,</p>
#                 <p style="font-size: 16px;">Veuillez trouver ci-joint l'extract post Production de l'activité DAAM pour la période du 
#                 <strong style="color: #004986;">{start_date}</strong> au <strong style="color: #004986;">{end_date}</strong>.</p>
                
#             </div>
#             <p><strong>NINGEN Data Analytics</strong><br></p>
                    
#                     <div style="margin-top: 20px; font-size: 12px; color: #666;">
#                         <p>
#                             Ceci est un message généré automatiquement. Merci de ne pas y répondre.<br> 
#                             <strong>Besoin d'assistance ?</strong><br>
#                             Veuillez contacter :
#                             <a href="mailto:Ningen-Data-Management@ningen-group.com">
#                                 Ningen-Data-Management@ningen-group.com
#                             </a>
#                         </p>
#                     </div>

#             <div style="text-align: center; margin-top: 20px; font-size: 10px; color: #666;">
#                         <p>
#                             Ce message et les éventuelles pièces jointes sont strictement confidentiels et destinés exclusivement au(x) destinataire(s) indiqué(s). Toute utilisation, diffusion ou reproduction non autorisée est interdite. Si vous avez reçu ce message par erreur, merci d’en avertir immédiatement l’expéditeur et de supprimer le courriel.
#                         </p>
#             </div>
#         </div>
#     </body>
#     </html>
#     '''

#     # --- Initialisation Outlook ---
#     pythoncom.CoInitialize()
#     outlook = win32.Dispatch('Outlook.Application')
#     mail = outlook.CreateItem(0)
#     mail.Subject = f'DAAM - Extract post Production du {start_date} au {end_date}'
#     #mail.To = "ix41p@ningen-group.com"
#     #mail.CC = "yq68h@ningen-group.com"
#     mail.To = "ci68t@ningen-group.com;ue27p@ningen-group.com;yq68h@ningen-group.com;qq13f@ningen-group.com;cl37t@ningen-group.com"
#     mail.CC = "sl06h@ningen-group.com;td75s@ningen-group.com;iz55x@ningen-group.com;Ningen-Data-Management@ningen-group.com;ex59j@ningen-group.com;pw39f@ningen-group.com;Ningen-pperformance@ningen-group.com"
#     mail.HTMLBody = html_body
#     mail.Attachments.Add(temp_excel_path)
#     mail.Send()

#     # --- Attente d'ouverture d'Outlook ---
#     time.sleep(5)

#     # --- Connexion avec pywinauto et envoi automatique ---
#     try:
#         app = Application(backend="uia").connect(title_re=f".*DAAM - Extract post Production du {start_date}.*")
#         window = app.window(title_re=f".*DAAM - Extract post Production du {start_date}.*")
#         send_button = window.child_window(title="Envoyer", control_type="Button")
#         send_button.wait("enabled", timeout=10)
#         send_button.click_input()
#         write_log("✅ Email envoyé avec succès via Outlook.")
#     except Exception as send_error:
#         write_log(f"⚠️ Impossible d'envoyer l'e-mail automatiquement : {send_error}")
#         write_log("➡️ Veuillez vérifier manuellement dans Outlook.")

#     # --- Nettoyage fichier temporaire ---
#     if os.path.exists(temp_excel_path):
#         os.remove(temp_excel_path)
#         write_log(f"🗑️ Fichier temporaire supprimé : {temp_excel_path}")

#     pythoncom.CoUninitialize()
#     gc.collect()

# except Exception as e:
#     write_log(f"❌ Erreur lors de la préparation ou de l'envoi de l'e-mail : {str(e)}")








##########################################################################################
#########################################################################################

#====================================================================================

#=================== new code ====================================================

#===================================================================================

#############################################################################################
############################################################################################




import os
import json
import re
import html
import unicodedata
from io import BytesIO

import pandas as pd
import requests
import msal
from dotenv import load_dotenv
from sshtunnel import SSHTunnelForwarder
import mysql.connector
from mysql.connector import Error
import pandas as pd
import requests
import msal
from io import BytesIO
from datetime import datetime, date, timedelta
import os
from dotenv import load_dotenv
import win32com.client as win32
import pythoncom
from dotenv import load_dotenv

# ============================================================
# LOAD ENV
# ============================================================
env_path = r"C:\Users\Administrateur\Desktop\Daam\DAAM_Post_Prod\.env"
load_dotenv(dotenv_path=env_path)

# =========================
# SSH (GATEWAY)
# =========================

SSH_HOST = os.getenv("SSH_HOST")
SSH_PORT = int(os.getenv("SSH_PORT", "22"))
SSH_USER = os.getenv("SSH_USER")
SSH_PASSWORD = os.getenv("SSH_PASSWORD")

DB_USER = os.getenv("DB_USER")
DB_PASSWORD = os.getenv("DB_PASSWORD")
DB_NAME = os.getenv("DB_NAME")
REMOTE_MYSQL_HOST  = os.getenv("REMOTE_MYSQL_HOST")
REMOTE_MYSQL_PORT  = int(os.getenv("REMOTE_MYSQL_PORT"))


# =========================
# SHAREPOINT
# =========================
TENANT_ID     = os.getenv("TENANT_ID")
CLIENT_ID     = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
SITE_DOMAIN = os.getenv("SITE_DOMAIN")
SITE_NAME     = os.getenv("SITE_NAME")    
DRIVE_NAME = "Documents"

SHAREPOINT_FILE_LOG_PATH ="General/Autoreports Status"
# Log file name
log_filename = f"Post_Prod_DAAM_{datetime.now().strftime('%Y%m%d_%H.%M')}.txt"

headers = None
drive_id = None
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
        print(f" Impossible de logger sur SharePoint : {e}")

def main_surveydata():
    TABLE_NAME = "myapp_surveydata"
    TARGETS = [
        {
            "path": "General/DAAM/Prospection Lead Chaud/raw_data/outbound_surveydata.xlsx",
            "survey_schema_id": "36",
        },
        {
            "path": "General/DAAM/C-KYC/raw_data/outbound_surveydata.xlsx",
            "survey_schema_id": "25",
        },
        {
            "path": "General/DAAM/CAF/raw_data/outbound_surveydata.xlsx",
            "survey_schema_id": "26",
        },
        {
            "path": "General/DAAM/Welcome Call/raw_data/outbound_surveydata.xlsx",
            "survey_schema_id": "27",
        },
        {
            "path": "General/DAAM/Prospection/raw_data/outbound_surveydata.xlsx",
            "survey_schema_id": "28",
        },
        {
            "path": "General/DAAM/suivi C-KYC/raw_data/outbound_surveydata.xlsx",
            "survey_schema_id": "32",
        },
        {
            "path": "General/DAAM/Relance CAF/raw_data/outbound_surveydata.xlsx",
            "survey_schema_id": "33",
        },
    ]

    # ============================================================
    # UTILS
    # ============================================================
    def require(v, name):
        if v is None or str(v).strip() == "":
            raise ValueError(f"❌ Missing env var: {name}")
        return str(v).strip()

    def normalize_key(s: str) -> str:
        s = str(s).strip().lower()
        s = unicodedata.normalize("NFKD", s)
        s = "".join(ch for ch in s if not unicodedata.combining(ch))
        s = re.sub(r"[^a-z0-9]+", "", s)
        return s

    def pick_existing_key_column(cols):
        """
        Priorité matching:
        1) inboundformresponse beside id (si existe)
        2) formresponse beside id / beside id / id (fallback)
        """
        if not cols:
            return None
        norm_map = {normalize_key(c): c for c in cols}

        candidates = [
            "inboundformresponse beside id",
            "inboundformresponse_beside_id",
            "formresponse beside id",
            "formresponse_beside_id",
            "beside id",
            "beside_id",
            "id",
        ]
        for cand in candidates:
            k = normalize_key(cand)
            if k in norm_map:
                return norm_map[k]
        return None

    def to_key_series(df: pd.DataFrame, key_col: str) -> pd.Series:
        s = df[key_col].astype("string")
        s = s.fillna(pd.NA)
        return s


    # ============================================================
    # SHAREPOINT AUTH / DRIVE / DOWNLOAD / UPLOAD
    # ============================================================
    def authenticate_sharepoint():
        TENANT_ID_ = require(TENANT_ID, "TENANT_ID")
        CLIENT_ID_ = require(CLIENT_ID, "CLIENT_ID")
        CLIENT_SECRET_ = require(CLIENT_SECRET, "CLIENT_SECRET")
        require(SITE_DOMAIN, "SITE_DOMAIN")
        require(SITE_NAME, "SITE_NAME")

        authority = f"https://login.microsoftonline.com/{TENANT_ID_}"
        scope = ["https://graph.microsoft.com/.default"]

        app = msal.ConfidentialClientApplication(
            CLIENT_ID_,
            authority=authority,
            client_credential=CLIENT_SECRET_,
        )
        token = app.acquire_token_for_client(scopes=scope)
        if "access_token" not in token:

            raise Exception(f" Token error: {token}")

        return {"Authorization": f"Bearer {token['access_token']}"}

    def get_site_id(headers):
        site_url = f"https://graph.microsoft.com/v1.0/sites/{SITE_DOMAIN}:/sites/{SITE_NAME}"
        site = requests.get(site_url, headers=headers).json()
        if "id" not in site:
            raise Exception(f" Site not found/unauthorized: {site}")
        return site["id"]

    def get_drive_id(headers, site_id):
        drives = requests.get(
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives",
            headers=headers,
        ).json()
        if "value" not in drives:
            raise Exception(f" Could not list drives: {drives}")

        wanted = (DRIVE_NAME or "Documents").strip().lower()
        for d in drives["value"]:
            if d.get("name", "").strip().lower() == wanted:
                return d["id"]

        for d in drives["value"]:
            if d.get("name", "").strip().lower() == "documents":
                return d["id"]

        raise Exception(" Drive not found (Documents)")

    def download_sharepoint_file_bytes(headers, drive_id, sp_path: str) -> bytes:
        if not sp_path:
            raise ValueError(" SHAREPOINT_FILE_PATH is empty")
        path = "/" + sp_path.strip().strip("/")
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:{path}:/content"
        r = requests.get(url, headers=headers)
        if r.status_code != 200:
            raise Exception(f" Download failed: {r.status_code} {r.text}")
        return r.content

    def upload_sharepoint_file_bytes(headers, drive_id, sp_path: str, content: bytes):
        path = "/" + sp_path.strip().strip("/")
        write_log(f" Upload overwrite: {path}")
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:{path}:/content"
        r = requests.put(url, headers=headers, data=content)
        if r.status_code not in (200, 201):
            raise Exception(f" Upload failed: {r.status_code} {r.text}")
        write_log(" Uploaded (updated file)")


    # ============================================================
    # JSON FLATTEN + CLEAN COLUMNS (safe)
    # ============================================================
    _u00 = re.compile(r"u00([0-9a-fA-F]{2})")
    _CuXXXX = re.compile(r"([A-Za-z])u([0-9a-fA-F]{4})")

    def fix_mojibake(s: str) -> str:
        if not isinstance(s, str):
            s = str(s)

        suspects = ("Ã", "Â", "â€™", "â€œ", "â€�", "â€“", "â€”", "â€¦")
        if any(x in s for x in suspects):
            try:
                fixed = s.encode("latin-1", errors="ignore").decode("utf-8", errors="ignore")
                if fixed.count("Ã") + fixed.count("Â") < s.count("Ã") + s.count("Â"):
                    s = fixed
            except Exception:
                pass

        s = (s.replace("â€™", "’")
            .replace("â€œ", "“")
            .replace("â€�", "”")
            .replace("â€“", "–")
            .replace("â€”", "—")
            .replace("â€¦", "…")
            .replace("Â", ""))

        return s

    def decode_unicode_escapes(s: str) -> str:
        s = html.unescape(str(s))
        s = _CuXXXX.sub(lambda m: f"{m.group(1)}\\u{m.group(2)}", s)
        s = _u00.sub(lambda m: r"\u00" + m.group(1), s)
        s = re.sub(r"(?<!\\)u([0-9a-fA-F]{4})", lambda m: r"\u" + m.group(1), s)
        s = s.replace("\\\\u", "\\u")

        if "\\u" in s or "\\U" in s:
            try:
                s = bytes(s, "utf-8").decode("unicode_escape")
            except Exception:
                pass

        return unicodedata.normalize("NFC", s)

    def clean_column_name(col: str) -> str:
        s = str(col)
        s = fix_mojibake(s)
        s = decode_unicode_escapes(s)
        s = re.sub(r"\s*\(\d+\)\s*", " ", s)
        s = s.replace("_", " ")
        s = re.sub(r"\s+", " ", s).strip()
        s = "".join(c for c in s if unicodedata.category(c)[0] != "C")
        return s if s else "Colonne"

    def make_unique(cols):
        seen = {}
        out = []
        for c in cols:
            if c not in seen:
                seen[c] = 0
                out.append(c)
            else:
                seen[c] += 1
                out.append(f"{c} ({seen[c]})")
        return out

    def clean_dataframe_columns(df: pd.DataFrame) -> pd.DataFrame:
        df.columns = make_unique([clean_column_name(c) for c in df.columns])
        return df

    def try_parse_json(v):
        if isinstance(v, (dict, list)):
            return v
        if isinstance(v, bytes):
            v = v.decode("utf-8", "ignore")
        if isinstance(v, str):
            s = v.strip()
            if not s:
                return None
            if s.startswith("{") or s.startswith("["):
                try:
                    return json.loads(s)
                except Exception:
                    return None
        return None

    def flatten_all_json_columns(df: pd.DataFrame) -> pd.DataFrame:
        if df is None or df.empty:
            return df

        for col in df.columns:
            parsed_list = []
            parsed_any = False
            for v in df[col].tolist():
                p = try_parse_json(v)
                if p is not None:
                    parsed_any = True
                    parsed_list.append(p)
                else:
                    parsed_list.append(v)
            if parsed_any:
                df[col] = parsed_list

        while True:
            dict_cols = [
                c for c in df.columns
                if df[c].dropna().apply(lambda x: isinstance(x, dict)).any()
            ]
            if not dict_cols:
                break

            for c in dict_cols:
                expanded = pd.json_normalize(df[c]).add_prefix(f"{c}_")
                df = pd.concat([df.drop(columns=[c]), expanded], axis=1)

        for col in df.columns:
            if df[col].dropna().apply(lambda x: isinstance(x, list)).any():
                df[col] = df[col].apply(lambda x: json.dumps(x, ensure_ascii=False) if isinstance(x, list) else x)

        df = clean_dataframe_columns(df)
        return df


    # ============================================================
    # DB: CONNECT via SSH Tunnel
    # ============================================================
    def open_ssh_tunnel():
        require(SSH_HOST, "SSH_HOST")
        require(SSH_USER, "SSH_USER")
        require(SSH_PASSWORD, "SSH_PASSWORD")
        require(DB_USER, "DB_USER")
        require(DB_PASSWORD, "DB_PASSWORD")
        require(DB_NAME, "DB_NAME")

        tunnel = SSHTunnelForwarder(
            (SSH_HOST, SSH_PORT),
            ssh_username=SSH_USER,
            ssh_password=SSH_PASSWORD,
            remote_bind_address=(REMOTE_MYSQL_HOST, REMOTE_MYSQL_PORT),
            local_bind_address=("127.0.0.1", 0),
        )
        tunnel.start()
        write_log(f" SSH tunnel started: 127.0.0.1:{tunnel.local_bind_port} -> {REMOTE_MYSQL_HOST}:{REMOTE_MYSQL_PORT}")
        return tunnel

    def connect_mysql_via_tunnel(tunnel: SSHTunnelForwarder):
        try:
            conn = mysql.connector.connect(
                host="127.0.0.1",
                port=int(tunnel.local_bind_port),
                user=DB_USER,
                password=DB_PASSWORD,
                database=DB_NAME,
                use_pure=True,
                buffered=True,
            )
            if conn.is_connected():
                write_log(" Connected to MySQL (besidedb) via SSH tunnel")
            return conn
        except Error as e:
            write_log(f" MySQL connection error via SSH tunnel: {str(e)}")
            raise Exception(f" MySQL connection error via SSH tunnel: {e}")

    def list_table_columns(conn, table: str) -> list[str]:
        cur = conn.cursor()
        cur.execute(f"SHOW COLUMNS FROM {table}")
        cols = [row[0] for row in cur.fetchall()]
        cur.close()
        return cols

    def find_column_name(columns, variants):
        wanted = {normalize_key(v) for v in variants}
        for c in columns:
            if normalize_key(c) in wanted:
                return c
        return None

    def export_table_filtered(conn, survey_schema_id) -> pd.DataFrame:
        write_log(f"\n Export: myapp_surveydata (survey_schema_id='{survey_schema_id}')")
        cols = list_table_columns(conn, TABLE_NAME)

        form_id_col = find_column_name(
            cols,
            variants=["survey_schema_id", "surveyschemaid", "surveyschema beside id", "survey_schema_beside_id", "surveyschema id", "surveyschema_beside_id"],
        )
        if not form_id_col:
            write_log(f" Impossible de trouver la colonne surveyschema beside id dans myapp_surveydata. Colonnes: {cols}")
            raise Exception(f" Impossible de trouver la colonne surveyschema beside id dans myapp_surveydata. Colonnes: {cols}")

        q = f"SELECT id,created_at,id as surveydata_beside_id,survey_schema_id as surveyschema_beside_id,data FROM myapp_surveydata WHERE {form_id_col} = %s"
        cur = conn.cursor(dictionary=True)
        cur.execute(q, (survey_schema_id,))
        rows = cur.fetchall()
        cur.close()

        write_log(f" Rows fetched from DB: {len(rows)}")
        if not rows:
            return pd.DataFrame()

        df = pd.DataFrame(rows)
        df = flatten_all_json_columns(df)
        return df


    # ============================================================
    # ALIGN DB DF -> EXISTING EXCEL COLUMNS (keep same order)
    # ============================================================
    def build_rename_map_to_match_existing(db_cols, existing_cols):
        existing_norm = {}
        for c in existing_cols:
            existing_norm.setdefault(normalize_key(c), c)

        rename_map = {}
        for c in db_cols:
            nk = normalize_key(c)
            if nk in existing_norm:
                rename_map[c] = existing_norm[nk]
        return rename_map

    def align_df_to_existing_columns(df_db: pd.DataFrame, existing_cols: list[str]) -> pd.DataFrame:
        if df_db is None or df_db.empty:
            return df_db

        rename_map = build_rename_map_to_match_existing(df_db.columns.tolist(), existing_cols)
        df_db = df_db.rename(columns=rename_map)

        for c in existing_cols:
            if c not in df_db.columns:
                df_db[c] = pd.NA

        df_db = df_db[[c for c in existing_cols]].copy()
        return df_db


    # ============================================================
    # MAIN
    # ============================================================
    def main():
        global headers,drive_id
        try:
            # SharePoint init (une seule fois)
            headers = authenticate_sharepoint()
            site_id = get_site_id(headers)
            drive_id = get_drive_id(headers, site_id)

            tunnel = None
            conn = None

            try:
                tunnel = open_ssh_tunnel()
                conn = connect_mysql_via_tunnel(tunnel)

                for target in TARGETS:

                    sp_path = target["path"]
                    survey_schema_id = target["survey_schema_id"]

                    write_log("\n" + "="*70)
                    write_log(f" Processing: {sp_path}")
                    write_log(f" survey_schema_id = {survey_schema_id}")

                    # =========================
                    # DOWNLOAD EXISTING FILE
                    # =========================
                    existing_bytes = download_sharepoint_file_bytes(
                        headers, drive_id, sp_path
                    )

                    df_existing = pd.read_excel(BytesIO(existing_bytes), dtype="string")
                    write_log(f" Existing rows: {len(df_existing)}")

                    existing_cols = df_existing.columns.tolist()
                    key_col_existing = pick_existing_key_column(existing_cols)

                    if not key_col_existing:
                        write_log(
                            f" Matching column not found in {sp_path}"
                        )
                        raise Exception(
                            f" Matching column not found in {sp_path}"
                        )

                    existing_keys = set(
                        to_key_series(df_existing, key_col_existing)
                        .dropna()
                        .astype(str)
                        .tolist()
                    )

                    # =========================
                    # EXPORT DB FILTERED
                    # =========================
                    df_db = export_table_filtered(conn, survey_schema_id)

                    if df_db.empty:
                        write_log(" No DB data for this survey.")
                        continue

                    df_db_aligned = align_df_to_existing_columns(
                        df_db, existing_cols
                    )

                    if key_col_existing not in df_db_aligned.columns:
                        write_log(
                            f" Key column '{key_col_existing}' missing after alignment."
                        )
                        raise Exception(
                            f" Key column '{key_col_existing}' missing after alignment."
                        )

                    db_keys_series = to_key_series(
                        df_db_aligned, key_col_existing
                    )

                    mask_new = ~db_keys_series.fillna(pd.NA)\
                        .astype("string")\
                        .isin(list(existing_keys))

                    df_new = df_db_aligned[mask_new].copy()

                    write_log(f" New rows: {len(df_new)}")

                    if df_new.empty:
                        write_log("Already up-to-date.")
                        continue

                    df_out = pd.concat(
                        [df_existing, df_new], ignore_index=True
                    )

                    out_buf = BytesIO()
                    with pd.ExcelWriter(out_buf, engine="openpyxl") as writer:
                        df_out.to_excel(writer, index=False)

                    upload_sharepoint_file_bytes(
                        headers, drive_id, sp_path, out_buf.getvalue()
                    )

                    write_log(" File updated successfully.")

                write_log("\n ALL FILES PROCESSED SUCCESSFULLY.")

            finally:
                try:
                    if conn:
                        conn.close()
                except:
                    pass
                try:
                    if tunnel:
                        tunnel.stop()
                except:
                    pass
        except Exception as e:
            write_log(f"Erreur dans main_surveydata(): {e}")

    main()

def main_surveyresults():
    TABLE_NAME = "myapp_formresponse"
    TARGETS = [
        {
            "path": "General/DAAM/Prospection Lead Chaud/raw_data/outbound_surveyresults.xlsx",
            "survey_schema_id": "36",
        },
        {
            "path": "General/DAAM/C-KYC/raw_data/outbound_surveyresults.xlsx",
            "survey_schema_id": "25",
        },
        {
            "path": "General/DAAM/CAF/raw_data/outbound_surveyresults.xlsx",
            "survey_schema_id": "26",
        },
        {
            "path": "General/DAAM/Welcome Call/raw_data/outbound_surveyresults.xlsx",
            "survey_schema_id": "27",
        },
        {
            "path": "General/DAAM/Prospection/raw_data/outbound_surveyresults.xlsx",
            "survey_schema_id": "28",
        },
        {
            "path": "General/DAAM/suivi C-KYC/raw_data/outbound_surveyresults.xlsx",
            "survey_schema_id": "32",
        },
        {
            "path": "General/DAAM/Relance CAF/raw_data/outbound_surveyresults.xlsx",
            "survey_schema_id": "33",
        },
    ]


    # ============================================================
    # UTILS
    # ============================================================
    def require(v, name):
        if v is None or str(v).strip() == "":
            raise ValueError(f"❌ Missing env var: {name}")
        return str(v).strip()

    def normalize_key(s: str) -> str:
        s = str(s).strip().lower()
        s = unicodedata.normalize("NFKD", s)
        s = "".join(ch for ch in s if not unicodedata.combining(ch))
        s = re.sub(r"[^a-z0-9]+", "", s)
        return s

    def pick_existing_key_column(cols):
        
        if not cols:
            return None
        norm_map = {normalize_key(c): c for c in cols}

        candidates = [
            
            "formresponse beside id",
            "formresponse_beside_id",
            "beside id",
            "beside_id",
            "id",
        ]
        for cand in candidates:
            k = normalize_key(cand)
            if k in norm_map:
                return norm_map[k]
        return None

    def to_key_series(df: pd.DataFrame, key_col: str) -> pd.Series:
        s = df[key_col].astype("string")
        s = s.fillna(pd.NA)
        return s


    # ============================================================
    # SHAREPOINT AUTH / DRIVE / DOWNLOAD / UPLOAD
    # ============================================================
    def authenticate_sharepoint():
        TENANT_ID_ = require(TENANT_ID, "TENANT_ID")
        CLIENT_ID_ = require(CLIENT_ID, "CLIENT_ID")
        CLIENT_SECRET_ = require(CLIENT_SECRET, "CLIENT_SECRET")
        require(SITE_DOMAIN, "SITE_DOMAIN")
        require(SITE_NAME, "SITE_NAME")

        authority = f"https://login.microsoftonline.com/{TENANT_ID_}"
        scope = ["https://graph.microsoft.com/.default"]

        app = msal.ConfidentialClientApplication(
            CLIENT_ID_,
            authority=authority,
            client_credential=CLIENT_SECRET_,
        )
        token = app.acquire_token_for_client(scopes=scope)
        if "access_token" not in token:
            raise Exception(f" Token error: {token}")

        return {"Authorization": f"Bearer {token['access_token']}"}

    def get_site_id(headers):
        site_url = f"https://graph.microsoft.com/v1.0/sites/{SITE_DOMAIN}:/sites/{SITE_NAME}"
        site = requests.get(site_url, headers=headers).json()
        if "id" not in site:
            raise Exception(f" Site not found/unauthorized: {site}")
        return site["id"]

    def get_drive_id(headers, site_id):
        drives = requests.get(
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives",
            headers=headers,
        ).json()
        if "value" not in drives:
            raise Exception(f" Could not list drives: {drives}")

        wanted = (DRIVE_NAME or "Documents").strip().lower()
        for d in drives["value"]:
            if d.get("name", "").strip().lower() == wanted:
                return d["id"]

        for d in drives["value"]:
            if d.get("name", "").strip().lower() == "documents":
                return d["id"]

        raise Exception(" Drive not found (Documents)")

    def download_sharepoint_file_bytes(headers, drive_id, sp_path: str) -> bytes:
        if not sp_path:
            raise ValueError(" SHAREPOINT_FILE_PATH is empty")
        path = "/" + sp_path.strip().strip("/")
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:{path}:/content"
        r = requests.get(url, headers=headers)
        if r.status_code != 200:
            raise Exception(f" Download failed: {r.status_code} {r.text}")
        return r.content

    def upload_sharepoint_file_bytes(headers, drive_id, sp_path: str, content: bytes):
        path = "/" + sp_path.strip().strip("/")
        write_log(f" Upload overwrite: {path}")
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:{path}:/content"
        r = requests.put(url, headers=headers, data=content)
        if r.status_code not in (200, 201):
            raise Exception(f" Upload failed: {r.status_code} {r.text}")
        write_log(" Uploaded (updated file)")


    # ============================================================
    # JSON FLATTEN + CLEAN COLUMNS (safe)
    # ============================================================
    _u00 = re.compile(r"u00([0-9a-fA-F]{2})")
    _CuXXXX = re.compile(r"([A-Za-z])u([0-9a-fA-F]{4})")


    def remove_parenthesis_suffix_from_columns(df: pd.DataFrame) -> pd.DataFrame:
        """
        Supprime ' (1)', ' (2)', ... à la fin des noms de colonnes
        """
        df = df.copy()
        df.columns = [re.sub(r"\s*\(\d+\)\s*$", "", str(c)).strip() for c in df.columns]
        return df


    def coalesce_duplicate_columns(df: pd.DataFrame) -> pd.DataFrame:
        """
        Si des colonnes ont le même nom (après suppression de (1)),
        on les fusionne: on prend la première valeur non vide sur la ligne.
        Résultat: colonnes uniques => pd.concat ne plante plus.
        """
        df = df.copy()
        cols = list(df.columns)
        # garder l'ordre d'apparition
        seen = []
        for c in cols:
            if c not in seen:
                seen.append(c)

        out = pd.DataFrame(index=df.index)
        for c in seen:
            same = df.loc[:, df.columns == c]
            if same.shape[1] == 1:
                out[c] = same.iloc[:, 0]
            else:
                # first non-null across duplicated columns
                out[c] = same.bfill(axis=1).iloc[:, 0]
        return out

    def fix_mojibake(s: str) -> str:
        if not isinstance(s, str):
            s = str(s)

        suspects = ("Ã", "Â", "â€™", "â€œ", "â€�", "â€“", "â€”", "â€¦")
        if any(x in s for x in suspects):
            try:
                fixed = s.encode("latin-1", errors="ignore").decode("utf-8", errors="ignore")
                if fixed.count("Ã") + fixed.count("Â") < s.count("Ã") + s.count("Â"):
                    s = fixed
            except Exception:
                pass

        s = (s.replace("â€™", "’")
            .replace("â€œ", "“")
            .replace("â€�", "”")
            .replace("â€“", "–")
            .replace("â€”", "—")
            .replace("â€¦", "…")
            .replace("Â", ""))

        return s

    def decode_unicode_escapes(s: str) -> str:
        s = html.unescape(str(s))
        s = _CuXXXX.sub(lambda m: f"{m.group(1)}\\u{m.group(2)}", s)
        s = _u00.sub(lambda m: r"\u00" + m.group(1), s)
        s = re.sub(r"(?<!\\)u([0-9a-fA-F]{4})", lambda m: r"\u" + m.group(1), s)
        s = s.replace("\\\\u", "\\u")

        if "\\u" in s or "\\U" in s:
            try:
                s = bytes(s, "utf-8").decode("unicode_escape")
            except Exception:
                pass

        return unicodedata.normalize("NFC", s)

    def clean_column_name(col: str) -> str:
        s = str(col)
        s = fix_mojibake(s)
        s = decode_unicode_escapes(s)
        s = re.sub(r"\s*\(\d+\)\s*", " ", s)
        s = s.replace("_", " ")
        s = re.sub(r"\s+", " ", s).strip()
        s = "".join(c for c in s if unicodedata.category(c)[0] != "C")
        return s if s else "Colonne"

    def make_unique(cols):
        seen = {}
        out = []
        for c in cols:
            if c not in seen:
                seen[c] = 0
                out.append(c)
            else:
                seen[c] += 1
                out.append(f"{c} ({seen[c]})")
        return out

    def clean_dataframe_columns(df: pd.DataFrame) -> pd.DataFrame:
        df.columns = make_unique([clean_column_name(c) for c in df.columns])
        return df

    def try_parse_json(v):
        if isinstance(v, (dict, list)):
            return v
        if isinstance(v, bytes):
            v = v.decode("utf-8", "ignore")
        if isinstance(v, str):
            s = v.strip()
            if not s:
                return None
            if s.startswith("{") or s.startswith("["):
                try:
                    return json.loads(s)
                except Exception:
                    return None
        return None

    def flatten_all_json_columns(df: pd.DataFrame) -> pd.DataFrame:
        if df is None or df.empty:
            return df

        for col in df.columns:
            parsed_list = []
            parsed_any = False
            for v in df[col].tolist():
                p = try_parse_json(v)
                if p is not None:
                    parsed_any = True
                    parsed_list.append(p)
                else:
                    parsed_list.append(v)
            if parsed_any:
                df[col] = parsed_list

        while True:
            dict_cols = [
                c for c in df.columns
                if df[c].dropna().apply(lambda x: isinstance(x, dict)).any()
            ]
            if not dict_cols:
                break

            for c in dict_cols:
                expanded = pd.json_normalize(df[c]).add_prefix(f"{c}_")
                df = pd.concat([df.drop(columns=[c]), expanded], axis=1)

        for col in df.columns:
            if df[col].dropna().apply(lambda x: isinstance(x, list)).any():
                df[col] = df[col].apply(lambda x: json.dumps(x, ensure_ascii=False) if isinstance(x, list) else x)

        df = clean_dataframe_columns(df)
        return df

    def strip_response_data_prefix(df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()
        df.columns = [
            c.replace("response data ", "", 1) if c.lower().startswith("response data ") else c
            for c in df.columns
        ]
        return df

    # ============================================================
    # DB: CONNECT via SSH Tunnel
    # ============================================================
    def open_ssh_tunnel():
        require(SSH_HOST, "SSH_HOST")
        require(SSH_USER, "SSH_USER")
        require(SSH_PASSWORD, "SSH_PASSWORD")
        require(DB_USER, "DB_USER")
        require(DB_PASSWORD, "DB_PASSWORD")
        require(DB_NAME, "DB_NAME")

        tunnel = SSHTunnelForwarder(
            (SSH_HOST, SSH_PORT),
            ssh_username=SSH_USER,
            ssh_password=SSH_PASSWORD,
            remote_bind_address=(REMOTE_MYSQL_HOST, REMOTE_MYSQL_PORT),
            local_bind_address=("127.0.0.1", 0),
        )
        tunnel.start()
        write_log(f" SSH tunnel started: 127.0.0.1:{tunnel.local_bind_port} -> {REMOTE_MYSQL_HOST}:{REMOTE_MYSQL_PORT}")
        return tunnel

    def connect_mysql_via_tunnel(tunnel: SSHTunnelForwarder):
        try:
            conn = mysql.connector.connect(
                host="127.0.0.1",
                port=int(tunnel.local_bind_port),
                user=DB_USER,
                password=DB_PASSWORD,
                database=DB_NAME,
                use_pure=True,
                buffered=True,
            )
            if conn.is_connected():
                write_log(" Connected to MySQL (besidedb) via SSH tunnel")
            return conn
        except Error as e:
            write_log(f" MySQL connection error via SSH tunnel: {str(e)}")
            raise Exception(f" MySQL connection error via SSH tunnel: {e}")

    def list_table_columns(conn, table: str) -> list[str]:
        cur = conn.cursor()
        cur.execute(f"SHOW COLUMNS FROM {table}")
        cols = [row[0] for row in cur.fetchall()]
        cur.close()
        return cols

    def find_column_name(columns, variants):
        wanted = {normalize_key(v) for v in variants}
        for c in columns:
            if normalize_key(c) in wanted:
                return c
        return None

    def export_table_filtered1(conn, survey_schema_id) -> pd.DataFrame:
        write_log(f"\n Export: {TABLE_NAME} (survey_schema_id='{survey_schema_id}')")
        cols = list_table_columns(conn, TABLE_NAME)

        form_id_col = find_column_name(
            cols,
            variants=["survey_schema_id", "surveyschemaid", "surveyschema beside id", "survey_schema_beside_id", "surveyschema id", "surveyschema_beside_id" ,"survey schema id"],
        )
        if not form_id_col:
            write_log(f" Impossible de trouver la colonne surveyschema beside id dans {TABLE_NAME}. Colonnes: {cols}")
            raise Exception(f" Impossible de trouver la colonne surveyschema beside id dans {TABLE_NAME}. Colonnes: {cols}")

        q = f""" select id,
                        response_data,
                        
                        created_at,
                        form_id as form_beside_id,
                        user_id as beside_id,
                        survey_data_id as surveydata_beside_id,
                        survey_schema_id as surveyschema_beside_id,
                        
                        id as formresponse_beside_id FROM myapp_formresponse  WHERE {form_id_col} = %s"""
        cur = conn.cursor(dictionary=True)
        cur.execute(q, (survey_schema_id,))
        rows = cur.fetchall()
        cur.close()

        write_log(f" Rows fetched from DB: {len(rows)}")
        if not rows:
            return pd.DataFrame()

        df = pd.DataFrame(rows)
        df = flatten_all_json_columns(df)
        return df


    # ============================================================
    # ALIGN DB DF -> EXISTING EXCEL COLUMNS (keep same order)
    # ============================================================
    def build_rename_map_to_match_existing(db_cols, existing_cols):
        existing_norm = {}
        for c in existing_cols:
            existing_norm.setdefault(normalize_key(c), c)

        rename_map = {}
        for c in db_cols:
            nk = normalize_key(c)
            if nk in existing_norm:
                rename_map[c] = existing_norm[nk]
        return rename_map

    def align_df_to_existing_columns1(df_db: pd.DataFrame, existing_cols: list[str]) -> pd.DataFrame:
        if df_db is None or df_db.empty:
            return df_db

        # Normalize DB column names first
        df_db = strip_response_data_prefix(df_db)

        # Build rename map (DB -> Excel)
        rename_map = build_rename_map_to_match_existing(
            df_db.columns.tolist(),
            existing_cols
        )
        df_db = df_db.rename(columns=rename_map)

        # Ensure all Excel columns exist
        for col in existing_cols:
            if col not in df_db.columns:
                df_db[col] = pd.NA

        # Keep Excel order + append any new DB columns at the end
        final_cols = existing_cols + [c for c in df_db.columns if c not in existing_cols]

        return df_db[final_cols]



    # ============================================================
    # MAIN
    # ============================================================
    def main():
        global headers,drive_id
        try:
            # SharePoint init (une seule fois)
            headers = authenticate_sharepoint()
            site_id = get_site_id(headers)
            drive_id = get_drive_id(headers, site_id)

            tunnel = None
            conn = None

            try:
                tunnel = open_ssh_tunnel()
                conn = connect_mysql_via_tunnel(tunnel)

                for target in TARGETS:

                    sp_path = target["path"]
                    survey_schema_id = target["survey_schema_id"]

                    write_log("\n" + "="*70)
                    write_log(f" Processing: {sp_path}")
                    write_log(f" survey_schema_id = {survey_schema_id}")

                    # =========================
                    # DOWNLOAD EXISTING FILE
                    # =========================
                    existing_bytes = download_sharepoint_file_bytes(
                        headers, drive_id, sp_path
                    )

                    df_existing = pd.read_excel(BytesIO(existing_bytes), dtype="string")
                    write_log(f" Existing rows: {len(df_existing)}")

                    existing_cols = df_existing.columns.tolist()
                    key_col_existing = pick_existing_key_column(existing_cols)

                    if not key_col_existing:
                        write_log(
                            f" Matching column not found in {sp_path}"
                        )
                        raise Exception(
                            f" Matching column not found in {sp_path}"
                        )

                    existing_keys = set(
                        to_key_series(df_existing, key_col_existing)
                        .dropna()
                        .astype(str)
                        .tolist()
                    )

                    # =========================
                    # EXPORT DB FILTERED
                    # =========================
                    df_db = export_table_filtered1(conn, survey_schema_id)

                    if df_db.empty:
                        write_log(" No DB data for this survey.")
                        continue

                    # Nettoyage spécifique à ton script
                    df_db = strip_response_data_prefix(df_db)
                    df_db = remove_parenthesis_suffix_from_columns(df_db)
                    df_db = coalesce_duplicate_columns(df_db)

                    # Align
                    df_db_aligned = align_df_to_existing_columns1(
                        df_db, existing_cols
                    )

                    if key_col_existing not in df_db_aligned.columns:
                        write_log(
                            f" Key column '{key_col_existing}' missing after alignment."
                        )
                        raise Exception(
                            f" Key column '{key_col_existing}' missing after alignment."
                        )

                    db_keys_series = to_key_series(
                        df_db_aligned, key_col_existing
                    )

                    mask_new = ~db_keys_series.fillna(pd.NA)\
                        .astype("string")\
                        .isin(list(existing_keys))

                    df_new = df_db_aligned[mask_new].copy()

                    write_log(f" New rows: {len(df_new)}")

                    if df_new.empty:
                        write_log(" Already up-to-date.")
                        continue

                    df_out = pd.concat(
                        [df_existing, df_new], ignore_index=True
                    )

                    out_buf = BytesIO()
                    with pd.ExcelWriter(out_buf, engine="openpyxl") as writer:
                        df_out.to_excel(writer, index=False)

                    upload_sharepoint_file_bytes(
                        headers, drive_id, sp_path, out_buf.getvalue()
                    )

                    write_log(" File updated successfully.")

                write_log("\n ALL FILES PROCESSED SUCCESSFULLY.")

            finally:
                try:
                    if conn:
                        conn.close()
                except:
                    pass
                try:
                    if tunnel:
                        tunnel.stop()
                except:
                    pass
        except Exception as e:
            write_log(f"Erreur dans main_surveyresults(): {e}")
    
            
    main()       
            
def main_mail():
    TARGETS = [
        # {
        #     "name": "Survey Conformité Prospection",
        #     "path": "General/DAAM/Survey Conformité Prospection/raw_data",
        # },
        
        # {
        #     "name": "Enquête sur les Pratiques de Recouvrement",
        #     "path": "General/DAAM/Enquête sur les Pratiques de Recouvrement/raw_data",
        # },
        # {
        #     "name": "Survey JUAKALI",
        #     "path": "General/DAAM/Survey JUAKALI/raw_data",
        # },
        #  {
        #     "name": "Renouvellement",
        #     "path": "General/DAAM/Renouvellement/raw_data",
        # },
        #  {
        #     "name": " Deal lock",
        #     "path": "General/DAAM/Deal lock/raw_data",
        # },
        
        # {
        #     "name": "Opening the door",
        #     "path": "General/DAAM/Opening the door/raw_data",
        # },
        # {
        #     "name": "First knock",
        #     "path": "General/DAAM/First knock/raw_data",
        # },
        
        
        
        {
            "name": "Prospection Lead Chaud",
            "path": "General/DAAM/Prospection Lead Chaud/raw_data",
        },
        {
            "name": "C-KYC",
            "path": "General/DAAM/C-KYC/raw_data",
        },
        {
            "name": "CAF",
            "path": "General/DAAM/CAF/raw_data",
        },
        {
            "name": "Welcome Call",
            "path": "General/DAAM/Welcome Call/raw_data",
        },
        {
            "name": "Prospection",
            "path": "General/DAAM/Prospection/raw_data",
        },
        {
            "name": "Suivi C-KYC",
            "path": "General/DAAM/suivi C-KYC/raw_data",
        },
        {
            "name": "Relance CAF",
            "path": "General/DAAM/Relance CAF/raw_data",
        },
    ]

    # =========================
    # AUTHENTIFICATION
    # =========================
    def authenticate():
        authority = f"https://login.microsoftonline.com/{TENANT_ID}"
        scope = ["https://graph.microsoft.com/.default"]

        app = msal.ConfidentialClientApplication(
            CLIENT_ID,
            authority=authority,
            client_credential=CLIENT_SECRET
        )

        token = app.acquire_token_for_client(scopes=scope)

        if "access_token" not in token:
            raise Exception("Erreur token")

        return {"Authorization": f"Bearer {token['access_token']}"}


    def get_drive_id(headers):
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

        raise Exception("Drive Documents non trouvé")


    def read_excel(headers, drive_id, path):
        r = requests.get(
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{path}:/content",
            headers=headers
        )
        return pd.read_excel(BytesIO(r.content))


    # =========================
    # TRAITEMENT POST PROD
    # =========================
    def create_post_prod(df_inj, df_rslt, source_name):

        # Convertir dates
        df_inj["created at"] = pd.to_datetime(df_inj["created at"], format='mixed', errors='coerce')
        df_inj["Date Injection"] = df_inj["created at"]
        
        df_rslt["created at"] = pd.to_datetime(df_rslt["created at"], format='mixed', errors='coerce')
        df_rslt["Date Traitement"] = df_rslt["created at"]
        
        
        # Merge
        df = df_rslt.merge(
            df_inj,
            on="surveydata beside id",
            how="left",
            suffixes=("_rslt", "_inj")
        )

        #df["SOURCE"] = source_name

        
        
        
    

        

        return df


    # =========================
    # MAIN
    # =========================
    def main():
        global headers,drive_id
        try:

            write_log(" Authentification...")
            headers = authenticate()
            drive_id = get_drive_id(headers)

            all_data = []

            write_log(" Lecture des dossiers...")

            for target in TARGETS:
                try:
                    write_log(f"   ➜ {target['name']}")

                    path_results = f"{target['path']}/outbound_surveyresults.xlsx"
                    path_data = f"{target['path']}/outbound_surveydata.xlsx"

                    df_rslt = read_excel(headers, drive_id, path_results)
                    df_inj = read_excel(headers, drive_id, path_data)

                    df_post = create_post_prod(df_inj, df_rslt, target["name"])

                    all_data.append(df_post)

                except Exception as e:
                    write_log(f" Erreur {target['name']} : {e}")

            # Consolidation globale
            final_df = pd.concat(all_data, ignore_index=True)

            # Filtre date (ex: depuis 01-06-2025)
            date_debut = datetime(2025, 12, 1)
            date_fin = datetime.today()

            final_df = final_df[
                (final_df["created at_rslt"] >= date_debut) &
                (final_df["created at_rslt"] <= date_fin)
            ]

            write_log(f" Total lignes finales: {len(final_df)}")

            # =========================
            # EXPORT LOCAL TEMP
            # =========================
            temp_path = os.path.join(os.getenv("TEMP"), "POST_PROD_DAAM.xlsx")
            final_df.to_excel(temp_path, index=False)

            write_log(" Envoi Email...")

            # =========================
            # EMAIL OUTLOOK
            # =========================
            pythoncom.CoInitialize()
            outlook = win32.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)

            start_date = "01-12-2025"
            end_date = date.today().strftime("%d-%m-%Y")

            mail.Subject = f"DAAM - Extract Post Production du {start_date} au {end_date}"

            #mail.To = "ix41p@ningen-group.com"
            #mail.CC = "pw39f@ningen-group.com"

            mail.To = "Ningen-pperformance@ningen-group.com;"
            mail.CC = "sl06h@ningen-group.com;td75s@ningen-group.com;iz55x@ningen-group.com;Ningen-Data-Management@ningen-group.com;ex59j@ningen-group.com;pw39f@ningen-group.com;qq13f@ningen-group.com;cl37t@ningen-group.com;ci68t@ningen-group.com"

            mail.HTMLBody = f"""
            <html>
            <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 0;">
                <div style="max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #ddd; border-radius: 8px; background-color: #f9f9f9; box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1); position: relative;">
                    <div style="position: absolute; top: 0px; left: 0px; padding: 5px;">
                        
                    </div>
                    <div style="margin-top: 60px;">
                        <p style="font-size: 16px;">Bonjour,</p>
                        <p style="font-size: 16px;">Veuillez trouver ci-joint l'extract post Production de l'activité DAAM pour la période du 
                        <strong style="color: #004986;">{start_date}</strong> au <strong style="color: #004986;">{end_date}</strong>.</p>
                        
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

            mail.Attachments.Add(temp_path)
            mail.Send()

            write_log(" Email envoyé avec succès !")
        except Exception as e:
            write_log(f"Erreur dans main_mail(): {str(e)}")
        
    main()

 
if __name__ == "__main__":
        
    try:
        main_surveydata()
        main_surveyresults()
        main_mail()
    except Exception as e:
        write_log(f"Erreur dans main(): {e}")