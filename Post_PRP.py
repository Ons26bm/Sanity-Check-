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
#             print(f"Erreur lors de la normalisation de la colonne {column}: {e}")

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
#         print("Connexion réussie à la base de données Azure MySQL")
#         cursor = conn.cursor()

#         query1 = """
#         SELECT *
#         FROM datawarehouse.outbound_surveyresults AS rslt
#         INNER JOIN datawarehouse.outbound_surveydata AS inj ON rslt.surveydata_beside_id = inj.surveydata_beside_id where rslt.surveyschema_beside_id in (30);
#         """

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
#         excel_file = r"C:\Users\Administrateur\Desktop\STAR\Post Prod\Post_Prod.xlsx"
#         import pandas as pd
#         from datetime import datetime

#         # Convertir les dates en format datetime
#         date_debut = datetime(2025, 10, 13)
#         date_aujourdhui = pd.to_datetime(datetime.today()) 

#         # Supposons que la colonne 'Date Traitement' est déjà au format datetime
#         # Si ce n'est pas le cas, il faut d'abord la convertir:
#         # Data['Date Traitement'] = pd.to_datetime(Data['Date Traitement'], dayfirst=True)

#         # Filtrer le dataframe
#         Data = Data[(Data['Date Traitement'] >= date_debut) & 
#                         (Data['Date Traitement'] <= date_aujourdhui)]
#         Data.to_excel(excel_file, index=False)
#         print(f"Les données ont été exportées avec succès vers {excel_file}")

# except mysql.connector.Error as e:
#     print(f"Erreur lors de la connexion à la base de données : {e}")

# finally:
#     if conn.is_connected():
#         cursor.close()
#         conn.close()
#         print("Connexion à la base de données fermée")




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
#     start_date = "27-10-2025"
#     end_date = (date.today() - timedelta(days=0)).strftime("%d-%m-%Y")

#     # --- Création fichier Excel temporaire ---
#     temp_folder = os.path.join(os.getenv('TEMP'), 'survey_attachments')
#     os.makedirs(temp_folder, exist_ok=True)
#     temp_excel_path = os.path.join(temp_folder, f"PostProd_STAR_PRP_{end_date}.xlsx")
#     Data.to_excel(temp_excel_path, sheet_name='Data_N', index=False)

#     if not os.path.exists(temp_excel_path):
#         raise Exception("Fichier Excel temporaire non créé.")
#     else:
#         print(f"Fichier Excel temporaire créé avec succès : {temp_excel_path}")

#     # --- Encode le logo ---
#     image_path = r"C:\Users\Administrateur\Desktop\STAR\Post Prod\logo-complet-color.png"
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
#                 <p style="font-size: 16px;">Veuillez trouver ci-joint l'extract post Production de l'activité STAR de la campagne PRP pour la période du 
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

#                     <div style="text-align: center; margin-top: 20px; font-size: 10px; color: #666;">
#                         <p>
#                             Ce message et les éventuelles pièces jointes sont strictement confidentiels et destinés exclusivement au(x) destinataire(s) indiqué(s). Toute utilisation, diffusion ou reproduction non autorisée est interdite. Si vous avez reçu ce message par erreur, merci d’en avertir immédiatement l’expéditeur et de supprimer le courriel.
#                         </p>
#                     </div>
#         </div>
#     </body>
#     </html>
#     '''

#     # --- Initialisation Outlook ---
#     pythoncom.CoInitialize()
#     outlook = win32.Dispatch('Outlook.Application')
#     mail = outlook.CreateItem(0)
#     mail.Subject = f' Post Prod STAR-PRP du {start_date} au {end_date}'
#     #mail.To = "rz35n@ningen-group.com"
#     #mail.CC = "yq68h@ningen-group.com"
#     mail.To = "Ningen-pperformance@ningen-group.com"
#     mail.CC = "sl06h@ningen-group.com;td75s@ningen-group.com;iz55x@ningen-group.com;Ningen-Data-Management@ningen-group.com;o.benmaaouia@ningen-group.com"
#     mail.HTMLBody = html_body
#     mail.Attachments.Add(temp_excel_path)
#     mail.Send()

#     # --- Attente d'ouverture d'Outlook ---
#     time.sleep(5)

#     # --- Connexion avec pywinauto et envoi automatique ---
#     try:
#         app = Application(backend="uia").connect(title_re=f".*STAR - Extract post Production du {start_date}.*")
#         window = app.window(title_re=f".*STAR - Extract post Production du {start_date}.*")
#         send_button = window.child_window(title="Envoyer", control_type="Button")
#         send_button.wait("enabled", timeout=10)
#         send_button.click_input()
#         print("✅ Email envoyé avec succès via Outlook.")
#     except Exception as send_error:
#         print(f"⚠️ Impossible d'envoyer l'e-mail automatiquement : {send_error}")
#         print("➡️ Veuillez vérifier manuellement dans Outlook.")

#     # --- Nettoyage fichier temporaire ---
#     if os.path.exists(temp_excel_path):
#         os.remove(temp_excel_path)
#         print(f"🗑️ Fichier temporaire supprimé : {temp_excel_path}")

#     pythoncom.CoUninitialize()
#     gc.collect()

# except Exception as e:
#     print(f"❌ Erreur lors de la préparation ou de l'envoi de l'e-mail : {str(e)}")



###########################################################
#=======================================================
#=================== new code =========================
#=======================================================
############################################################

from datetime import datetime, date, timedelta
import unicodedata
import win32com.client as win32
from dotenv import load_dotenv
import warnings
import mysql.connector
from mysql.connector import Error
import html
import os
import json
import gc
import pandas as pd
import requests
import msal
from io import BytesIO
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
warnings.filterwarnings('ignore')

# =====================
# Load environment
# =====================
#load_dotenv(dotenv_path=r"C:\Users\Administrateur\Desktop\STAR\Post Prod\STAR\Post Prod\env 9")
env_path = r"C:\Users\Administrateur\Desktop\STAR\Post Prod\.env"
load_dotenv(dotenv_path=env_path)
b = 20

# pylint: disable=duplicate-code
TENANT_ID     = os.getenv("TENANT_ID")
CLIENT_ID     = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
SITE_DOMAIN = os.getenv("SITE_DOMAIN")
SITE_NAME     = os.getenv("SITE_NAME")

RAW_DATA_PATH = "General/STAR/PRP/raw_data"

FILE_RSLT = f"{RAW_DATA_PATH}/outbound_surveyresults.xlsx"
FILE_INJ  = f"{RAW_DATA_PATH}/outbound_surveydata.xlsx"

# pylint: disable=duplicate-code
# SSH / MySQL
SSH_HOST = os.getenv("SSH_HOST")
SSH_PORT = int(os.getenv("SSH_PORT", "22"))
SSH_USER = os.getenv("SSH_USER")
SSH_PASSWORD = os.getenv("SSH_PASSWORD")
# pylint: disable=duplicate-code
DB_USER = os.getenv("DB_USER")
DB_PASSWORD = os.getenv("DB_PASSWORD")
DB_NAME = os.getenv("DB_NAME")
REMOTE_MYSQL_HOST  = os.getenv("REMOTE_MYSQL_HOST")
REMOTE_MYSQL_PORT  = int(os.getenv("REMOTE_MYSQL_PORT"))
# pylint: disable=duplicate-code
SURVEY_SCHEMA_ID  = 30  # ID du survey schema à filtrer dans les résultats
TABLE_NAME_surveydata = os.getenv("TABLE_NAME_surveydata")
TABLE_NAME = os.getenv("TABLE_NAME")
survey_schema_id=30

SHAREPOINT_FILE_LOG_PATH ="General/Autoreports Status"
# Log file name
log_filename = f"Post_Prod_PRP_{datetime.now().strftime('%Y%m%d_%H.%M')}.txt"
# =====================
# SharePoint Helpers
# =====================

headers = None
drive_id = None
# pylint: disable=duplicate-code
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
def require(v, name):
    if v is None or str(v).strip() == "":
        raise ValueError(f" Missing env var: {name}")
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




def authenticate():
    """Authentification Microsoft Graph API"""
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=authority,
        client_credential=CLIENT_SECRET
    )
    token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in token:
        raise Exception(f"Erreur lors de l'acquisition du token : {token}")
    return {"Authorization": f"Bearer {token['access_token']}"}


def get_drive_id(headers):
    """Récupère l'ID du drive SharePoint"""
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

    raise Exception("Bibliothèque Documents non trouvée")


def read_excel(headers, drive_id, path):
    """Lit un fichier Excel depuis SharePoint"""
    r = requests.get(
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{path}:/content",
        headers=headers
    )
    r.raise_for_status()
    return pd.read_excel(BytesIO(r.content))

def download_file_bytes(headers, drive_id, path: str) -> bytes:
    r = requests.get(
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{path}:/content",
        headers=headers
    )
    r.raise_for_status()
    return r.content

def download_sharepoint_file_bytes(headers, drive_id, sp_path: str) -> bytes:
    if not sp_path:
        raise ValueError(" SHAREPOINT_FILE_PATH is empty")
    path = "/" + sp_path.strip().strip("/")
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:{path}:/content"
    r = requests.get(url, headers=headers)
    if r.status_code != 200:
        write_log(f" Download failed: {r.status_code} {r.text}")
        raise Exception(f" Download failed: {r.status_code} {r.text}")
    return r.content



def upload_sharepoint_file_bytes(headers, drive_id, sp_path: str, content: bytes):
    path = "/" + sp_path.strip().strip("/")
    write_log(f" Upload overwrite: {path}")
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:{path}:/content"
    r = requests.put(url, headers=headers, data=content)
    if r.status_code not in (200, 201):
        write_log(f" Upload failed: {r.status_code} {r.text}")
        raise Exception(f" Upload failed: {r.status_code} {r.text}")
    write_log("Uploaded (updated file)")
    
    
def upload_file_bytes(headers, drive_id, path: str, content: bytes):
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{path}:/content"
    r = requests.put(url, headers=headers, data=content)
    if r.status_code not in (200, 201):
        write_log(f" Upload failed: {r.status_code} {r.text}")
        raise Exception(f" Upload failed: {r.status_code} {r.text}")
    write_log(f"✅ Fichier mis à jour sur SharePoint : {path}")


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

def export_table_filtered(conn) -> pd.DataFrame:
    
    cols = list_table_columns(conn, TABLE_NAME_surveydata)

    form_id_col = find_column_name(
        cols,
        variants=["survey_schema_id", "surveyschemaid", "surveyschema beside id", "survey_schema_beside_id", "surveyschema id", "surveyschema_beside_id"],
    )
    if not form_id_col:
        write_log(f" Impossible de trouver la colonne surveyschema beside id dans myapp_surveydata. Colonnes: {cols}")
        raise Exception(f" Impossible de trouver la colonne surveyschema beside id dans myapp_surveydata. Colonnes: {cols}")

    q = f"SELECT id,created_at,id as surveydata_beside_id,data FROM myapp_surveydata WHERE {form_id_col} = %s"
    cur = conn.cursor(dictionary=True)
    cur.execute(q, (30,))
    rows = cur.fetchall()
    cur.close()

    write_log(f"Rows fetched from DB: {len(rows)}")
    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame(rows)
    df = flatten_all_json_columns(df)
    return df

def export_table_filtered1(conn) -> pd.DataFrame:
    
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
                    user_id ,
                    survey_data_id as surveydata_beside_id,
                    survey_schema_id as surveyschema_beside_id,
                    survey_schema_id,
                    id as formresponse_beside_id FROM myapp_formresponse  WHERE {form_id_col} = %s"""
    cur = conn.cursor(dictionary=True)
    cur.execute(q, (30,))
    rows = cur.fetchall()
    cur.close()

    write_log(f"Rows fetched from DB: {len(rows)}")
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


###################################"
# Data Processing
# ################################
def build_data(df_rslt: pd.DataFrame, df_inj: pd.DataFrame) -> pd.DataFrame:
    

    # ── Normalise column names (strip spaces, lower for matching) ──────────
    df_rslt.columns = df_rslt.columns.str.strip()
    df_inj.columns  = df_inj.columns.str.strip()

    # The Excel files use "beside id" (space) while the SQL used "beside_id" (underscore).
    # Rename to a common key for the join.
    df_rslt = df_rslt.rename(columns={"surveydata beside id": "surveydata_beside_id",
                                       "surveyschema beside id": "surveyschema_beside_id",
                                       "user id": "user_id"})
    df_inj  = df_inj.rename(columns={"surveydata beside id": "surveydata_beside_id",
                                      "surveyschema beside id": "surveyschema_beside_id"})

    # ── Parse dates ─────────────────────────────────────────────────────────
    
    
    #df_rslt["Date Traitement"] = pd.to_datetime(df_rslt["created at"], format='mixed', errors='coerce')
    df_rslt["created at"] = pd.to_datetime(df_rslt["created at"], format='mixed', errors='coerce')
    
    #df_inj["Date Injection"]  = pd.to_datetime(df_inj["created at"],  format='mixed', errors='coerce')
    df_inj["created at"]  = pd.to_datetime(df_inj["created at"],  format='mixed', errors='coerce')

    # ── Rename inj "data *" columns to remove the "data " prefix ────────────
    

    # ── Apply SQL WHERE filters ─────────────────────────────────────────────
    cutoff = pd.Timestamp(date.today() - timedelta(days=1))

    df_rslt_filtered = df_rslt.copy()

    df_inj_filtered = df_inj.copy()

    # ── Join on surveydata_beside_id ────────────────────────────────────────
    Data = df_rslt_filtered.merge(
        df_inj_filtered,
        on="surveydata_beside_id",
        how="left",
        suffixes=("_rslt", "_inj")
    )
    date_debut = datetime(2025, 10, 13)
    date_aujourdhui = pd.to_datetime(datetime.today()) 
    
    Data = Data.rename(columns={"created at_rslt": "Date Traitement",
                                      "created at_inj": "Date Injection"})
    
    Data = Data[(Data['Date Traitement'] >= date_debut) & 
                        (Data['Date Traitement'] <= date_aujourdhui)]
    
    
    
    return Data


# =====================
# Main
# =====================

def main():
        
    global headers,drive_id
    try:
        print("Authentification SharePoint...")
        headers  = authenticate()
        drive_id = get_drive_id(headers)
        
        
        write_log("add to suvey SharePoint...")
        existing_bytes = download_sharepoint_file_bytes(headers, drive_id, FILE_INJ)
        df_existing = pd.read_excel(BytesIO(existing_bytes), dtype="string")
        existing_cols = df_existing.columns.tolist()
        key_col_existing = pick_existing_key_column(existing_cols)
        if not key_col_existing:
            write_log(" Impossible de déterminer la colonne de matching dans le fichier existant. ")
            raise Exception(
                " Impossible de déterminer la colonne de matching dans le fichier existant. "
                "Ajoute une colonne unique (ex: 'inboundformresponse beside id' ou 'id')."
            )
        
        existing_keys = set(
            to_key_series(df_existing, key_col_existing).dropna().astype(str).tolist()
        )

        # DB via SSH
        tunnel = None
        conn = None
        try:
            tunnel = open_ssh_tunnel()
            conn = connect_mysql_via_tunnel(tunnel)

            df_db = export_table_filtered(conn)
            if df_db.empty:
                write_log("No data in DB for this filter. Nothing to add.")
                

            # Align DB columns to existing columns (keep same order)
            df_db_aligned = align_df_to_existing_columns(df_db, existing_cols)

            if key_col_existing not in df_db_aligned.columns:
                write_log(f"Key column '{key_col_existing}' not found after alignment. DB columns: {df_db_aligned.columns.tolist()}")
                raise Exception(f"Key column '{key_col_existing}' not found after alignment.")

            db_keys_series = to_key_series(df_db_aligned, key_col_existing)
            mask_new = ~db_keys_series.fillna(pd.NA).astype("string").isin(list(existing_keys))
            df_new = df_db_aligned[mask_new].copy()

            write_log(f" New rows to append: {len(df_new)}")
            if df_new.empty:
                write_log("Nothing to append (file already up-to-date).")
                

            df_out = pd.concat([df_existing, df_new], ignore_index=True)

            out_buf = BytesIO()
            with pd.ExcelWriter(out_buf, engine="openpyxl") as writer:
                df_out.to_excel(writer, index=False)
            out_bytes = out_buf.getvalue()

            upload_sharepoint_file_bytes(headers, drive_id, FILE_INJ, out_bytes)

            write_log("\n DONE — appended missing rows only  into SharePoint Excel.")
        finally:
            try:
                if conn is not None:
                    conn.close()
            except Exception:
                pass
            try:
                if tunnel is not None:
                    tunnel.stop()
            except Exception:
                pass
        
        # ============================================================
        # add to result SharePoint...
        # ============================================================
        write_log("\n\nadd to result SharePoint...")
        existing_bytes1 = download_sharepoint_file_bytes(headers, drive_id, FILE_RSLT)
        df_existing1 = pd.read_excel(BytesIO(existing_bytes1), dtype="string")
        
        existing_cols1 = df_existing1.columns.tolist()
        key_col_existing1 = pick_existing_key_column(existing_cols1)
        if not key_col_existing1:
            write_log(" Impossible de déterminer la colonne de matching dans le fichier existant. ")
            raise Exception(
                " Impossible de déterminer la colonne de matching dans le fichier existant. "
                "Ajoute une colonne unique (ex: 'inboundformresponse beside id' ou 'id')."
            )
        
        existing_keys = set(
            to_key_series(df_existing1, key_col_existing1).dropna().astype(str).tolist()
        )

        # DB via SSH
        tunnel = None
        conn = None
        try:
            tunnel = open_ssh_tunnel()
            conn = connect_mysql_via_tunnel(tunnel)

            df_db1 = export_table_filtered1(conn)
            df_db1 = strip_response_data_prefix(df_db1)
            if df_db1.empty:
                write_log(" No data in DB for this filter. Nothing to add.")
                
            
            # Align DB columns to existing columns (keep same order)
            df_db_aligned1 = align_df_to_existing_columns1(df_db1, existing_cols1)
            

            if key_col_existing1 not in df_db_aligned1.columns:
                write_log(f" Key column '{key_col_existing1}' not found after alignment. DB columns: {df_db_aligned1.columns.tolist()}")
                raise Exception(f" Key column '{key_col_existing1}' not found after alignment.")

            db_keys_series1 = to_key_series(df_db_aligned1, key_col_existing1)
            mask_new1 = ~db_keys_series1.fillna(pd.NA).astype("string").isin(list(existing_keys))
            df_new1 = df_db_aligned1[mask_new1].copy()

            write_log(f" New rows to append: {len(df_new1)}")
            if df_new1.empty:
                write_log(" Nothing to append (file already up-to-date).")
                

            df_out1 = pd.concat([df_existing1, df_new1], ignore_index=True)

            out_buf = BytesIO()
            with pd.ExcelWriter(out_buf, engine="openpyxl") as writer:
                df_out1.to_excel(writer, index=False)
            out_bytes1 = out_buf.getvalue()

            upload_sharepoint_file_bytes(headers, drive_id, FILE_RSLT, out_bytes1)

            write_log("\n DONE — appended missing rows only  into SharePoint Excel.")
        finally:
            try:
                if conn is not None:
                    conn.close()
            except Exception:
                pass
            try:
                if tunnel is not None:
                    tunnel.stop()
            except Exception:
                pass
        
        
        
        
        
        write_log("Lecture des fichiers depuis SharePoint...")
        df_rslt = read_excel(headers, drive_id, FILE_RSLT)
        df_inj  = read_excel(headers, drive_id, FILE_INJ)
        write_log(f"   - Résultats  : {len(df_rslt)} lignes")
        write_log(f"   - Injections : {len(df_inj)} lignes")

        write_log("Construction du DataFrame post-production...")
        Data = build_data(df_rslt, df_inj)
        

        if Data.empty:
            write_log("Aucune donnée à exporter. Arrêt du script.")
            return

        # ── Dates pour nommage et corps du mail ─────────────────────────────────
        start_date = "27-10-2025"
        end_date   = (date.today() - timedelta(days=0)).strftime("%d-%m-%Y")

        # ── Fichier Excel temporaire ─────────────────────────────────────────────
        temp_folder    = os.path.join(os.getenv('TEMP'), 'survey_attachments')
        os.makedirs(temp_folder, exist_ok=True)
        temp_excel_path = os.path.join(
            temp_folder,
            f"PostProd_STAR_PRP_{end_date}.xlsx"
        )
        Data.to_excel(temp_excel_path, sheet_name='Data_N', index=False)

        if not os.path.exists(temp_excel_path):
            write_log(" Fichier Excel temporaire non créé.")
            raise Exception("Fichier Excel temporaire non créé.")
        write_log(f" Fichier Excel temporaire créé : {temp_excel_path}")

        # ── Encode le logo ────────────────────────────────────────────────────────
        image_path = r"C:\Users\Administrateur\Desktop\STAR\Post Prod\logo-complet-color.png"
        with Image.open(image_path) as img:
            img = img.resize((90, int(img.height * 90 / img.width)), Image.Resampling.LANCZOS)
            buffer = BytesIO()
            img.save(buffer, format="PNG")
            encoded_image = base64.b64encode(buffer.getvalue()).decode("utf-8")

        # ── Corps HTML ────────────────────────────────────────────────────────────
        html_body = f'''
        <html>
        <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 0;">
            <div style="max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #ddd; border-radius: 8px; background-color: #f9f9f9; box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1); position: relative;">
                <div style="position: absolute; top: 0px; left: 0px; padding: 5px;">
                    <img src="data:image/png;base64,{encoded_image}" alt="Logo" style="height: 30px; width: auto; object-fit: contain;">
                </div>
                <div style="margin-top: 60px;">
                    <p style="font-size: 16px;">Bonjour,</p>
                    <p style="font-size: 16px;">Veuillez trouver ci-joint l'extract post Production de l'activité STAR de la campagne PRP pour la période du 
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
        '''

        # ── Envoi Outlook ─────────────────────────────────────────────────────────
        write_log("Envoi de l'e-mail via Outlook...")
        try:
            pythoncom.CoInitialize()
            outlook = win32.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)
            mail.Subject = (
                f'Post Prod STAR-PRP du '
                f' {start_date} au {end_date}'
            )
            
            #mail.To = "ix41p@ningen-group.com"
            mail.To = "Ningen-pperformance@ningen-group.com"
            mail.CC = (
                "sl06h@ningen-group.com;"
                "td75s@ningen-group.com;"
                "iz55x@ningen-group.com;"
                "Ningen-Data-Management@ningen-group.com;"
                "o.benmaaouia@ningen-group.com"
            )
            mail.HTMLBody = html_body
            mail.Attachments.Add(temp_excel_path)
            mail.Send()
            write_log("E-mail envoyé avec succès via Outlook.")
        except Exception as send_error:
            write_log(f" Erreur lors de l'envoi de l'e-mail : {send_error}")
            write_log("Veuillez vérifier manuellement dans Outlook.")
        finally:
            pythoncom.CoUninitialize()
            gc.collect()

        # ── Nettoyage ─────────────────────────────────────────────────────────────
        if os.path.exists(temp_excel_path):
            os.remove(temp_excel_path)
            write_log(f" Fichier temporaire supprimé : {temp_excel_path}")

    except Exception as e:
        write_log(f"Erreur fatale dans main(): {e}")

if __name__ == "__main__":
    main()