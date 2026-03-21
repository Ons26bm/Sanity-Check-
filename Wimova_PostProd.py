
# import pandas as pd
# import mysql.connector
# from mysql.connector import Error
# import warnings
# import sys
# import json
# from datetime import datetime, date, timedelta
# import os
# import base64
# import win32com.client as win32

# # Configuration
# warnings.simplefilter("ignore", category=UserWarning)
# sys.stdout.reconfigure(encoding='utf-8')

# # Configuration de la base de données
# DB_CONFIG = {
#     'host': 'dataserver.mysql.database.azure.com',
#     'port': 3306,
#     'user': 'Admach',
#     'password': 'NiGan10gEn',
#     'database': 'datawarehouse',
#     'buffered': True
# }

# TARGET_TABLE = "wimova_batonnage"
# LOGO_PATH = r"C:\Users\Administrateur\Desktop\WIMOVA\Post Prod_Wimova\logo-complet-color.png"
# EXPORT_PATH = r"C:\Users\Administrateur\Desktop\WIMOVA\Post Prod_Wimova\downloads\batonnage_wimova_1.xlsx"

# # ------------------ Fonctions de base de données ------------------

# def connect_db():
#     """Établit une connexion à la base de données"""
#     try:
#         conn = mysql.connector.connect(**DB_CONFIG)
#         print("Connexion réussie à la base de données Azure")
#         return conn
#     except Exception as e:
#         print(f"Erreur de connexion à la base de données : {e}")
#         return None

# def fetch_survey_data():
#     """Récupère les données du survey depuis la base de données"""
#     conn = connect_db()
#     if conn is None:
#         return None

#     query = """
#         SELECT s.*, e.beside_name 
#         FROM datawarehouse.surveydatabatonnage s
#         LEFT JOIN datawarehouse.employees e ON s.beside_id = e.beside_id
#         WHERE s.survey = 'wimova survey' and s.created_at>'2025-12-16 22:00:00'; 
        

#     """
#     #s.created_at>'2025-09-1'
#     try:
#         df = pd.read_sql(query, con=conn)
#         print(f"Données récupérées : {len(df)} lignes")
#         return df
#     except Exception as e:
#         print(f"Erreur lors de l'exécution de la requête : {e}")
#         return None
#     finally:
#         if conn.is_connected():
#             conn.close()

# # ------------------ Fonctions de transformation des données ------------------

# def combine_date_time(df, date_col='Date', time_col='Heure', new_col='created_at'):
#     """Combine les colonnes Date et Heure en une colonne datetime"""
#     try:
#         if date_col in df.columns and time_col in df.columns:
#             df[date_col] = df[date_col].astype(str)
#             df[time_col] = df[time_col].astype(str)
#             df[new_col] = pd.to_datetime(df[date_col] + ' ' + df[time_col], errors='coerce')
#             df.drop(columns=[date_col, time_col], inplace=True)
#             print(f"Colonnes {date_col} et {time_col} fusionnées en {new_col}")
#         return df
#     except Exception as e:
#         print(f"Erreur lors de la fusion date/heure : {e}")
#         return df

# def normalize_json_column(df, col_name="data"):
#     """Normalise une colonne JSON en colonnes pandas"""
#     try:
#         if col_name in df.columns:
#             # Vérifier si les données sont déjà des dictionnaires ou des strings JSON
#             df[col_name] = df[col_name].apply(
#                 lambda x: json.loads(x) if isinstance(x, str) and x.strip().startswith('{') else x
#             )
            
#             # Normaliser seulement si c'est un dictionnaire
#             if all(isinstance(x, dict) for x in df[col_name] if pd.notnull(x)):
#                 df_json = pd.json_normalize(df[col_name])
#                 df = pd.concat([df.drop(columns=[col_name]), df_json], axis=1)
#                 print(f"Colonne '{col_name}' normalisée avec succès")
#                 # Afficher les colonnes disponibles pour débogage
#                 print(f"Colonnes disponibles après normalisation : {list(df.columns)}")
#             else:
#                 print(f"Colonne '{col_name}' ne contient pas de JSON valide")
#         return df
#     except Exception as e:
#         print(f"Erreur lors de la normalisation de '{col_name}' : {e}")
#         return df

# def merge_columns_with_conflict(row, columns):
#     """Fusionne plusieurs colonnes en gérant les conflits"""
#     values = []
#     for col in columns:
#         if col in row.index and pd.notnull(row[col]) and str(row[col]).strip() != '':
#             values.append(str(row[col]))
    
#     if not values:
#         return None
#     if len(set(values)) > 1:
#         return f"⚠️ Conflit: {', '.join(values)}"
#     return values[0]

# def merge_multiple_columns(df, prefix, new_col_name):
#     """Fusionne toutes les colonnes qui commencent par un certain préfixe"""
#     try:
#         cols_to_merge = [col for col in df.columns if col.startswith(prefix)]
#         if cols_to_merge:
#             df[new_col_name] = df.apply(
#                 lambda row: ', '.join(
#                     str(row[col]) for col in cols_to_merge 
#                     if pd.notnull(row[col]) and str(row[col]).strip() != ''
#                 ), 
#                 axis=1
#             )
#             df.drop(columns=cols_to_merge, inplace=True)
#             print(f"Colonnes avec préfixe '{prefix}' fusionnées en '{new_col_name}'")
#         return df
#     except Exception as e:
#         print(f"Erreur lors de la fusion des colonnes {prefix} : {e}")
#         return df

# # ------------------ Fonctions d'export ------------------

# def export_to_excel(df, filename):
#     """Exporte le DataFrame en fichier Excel"""
#     try:
#         # Créer le répertoire s'il n'existe pas
#         os.makedirs(os.path.dirname(filename), exist_ok=True)
        
#         df.to_excel(filename, index=False, engine="openpyxl")
#         print(f"Export réussi : {filename}")
#         return True
#     except Exception as e:
#         print(f"Erreur lors de l'export en Excel : {e}")
#         return False

# def save_excel_temp(df, sheet_name, temp_filename):
#     """Sauvegarde le DataFrame en Excel temporaire pour l'email"""
#     try:
#         if df.empty:
#             print("DataFrame vide - aucun fichier créé")
#             return False
        
#         # Créer le répertoire temporaire
#         os.makedirs(os.path.dirname(temp_filename), exist_ok=True)
        
#         df.to_excel(temp_filename, sheet_name=sheet_name, index=False, engine='openpyxl')
#         print(f"Fichier temporaire créé : {temp_filename}")
#         return True
#     except Exception as e:
#         print(f"Erreur lors de la création du fichier temporaire : {e}")
#         return False

# # ------------------ Fonctions d'email ------------------

# def send_email(df_final, attachment_filename):
#     """Envoie le DataFrame par email via Outlook"""
#     try:
#         start_date = "16-12-2025"
#         end_date = (date.today() - timedelta(days=1)).strftime("%d-%m-%Y")

#         outlook = win32.Dispatch('Outlook.Application')
#         mail = outlook.CreateItem(0)

#         mail.Subject = f'WIMOVA-Extract post Production du {start_date} au {end_date}'
#         LOGO_PATH = r"C:\Users\Administrateur\Desktop\WIMOVA\Post Prod_Wimova\logo-complet-color.png"

#         # Ajouter le logo en pièce jointe intégrée
#         if os.path.exists(LOGO_PATH):
#             logo_attachment = mail.Attachments.Add(LOGO_PATH)
#             logo_attachment.PropertyAccessor.SetProperty(
#                 "http://schemas.microsoft.com/mapi/proptag/0x3712001F", 
#                 "logo1.png"
#             )

#             with open(LOGO_PATH, "rb") as image_file:
#                 encoded_image = base64.b64encode(image_file.read()).decode('utf-8')
#         else:
#             encoded_image = ""
#             print("Avertissement : Logo introuvable")

#         mail.HTMLBody = f'''
#             <html>
#             <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
#                 <div style="max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #ddd; border-radius: 8px; background-color: #f9f9f9;">
#                     <div style="position: absolute; top: 20px; right: 20px;">
#                         <img src="cid:logo1.png" alt="Logo" style="height: 50px;">
#                     </div>
#                     <p>Bonjour,</p>
#                     <p>Veuillez trouver ci-joint l'extract post Production de l'activité WIMOVA pour la période du <strong>{start_date}</strong> au <strong>{end_date}</strong>.</p>
                    
#                     <p><strong>NINGEN Data Analytics</strong><br></p>
                    
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
#                 </div>
#             </body>
#             </html>
#             '''


#         mail.Sender = "rd22z@ningen-group.com"
#         mail.To = "Ningen-pperformance@ningen-group.com"
#         mail.CC = "sl06h@ningen-group.com;td75s@ningen-group.com;iz55x@ningen-group.com;Ningen-Data-Management@ningen-group.com;dn87g@ningen-group.com;qq13f@ningen-group.com"
#         #mail.To = "ix41p@ningen-group.com"
#         # Ajouter la pièce jointe
#         if os.path.exists(attachment_filename):
#             mail.Attachments.Add(attachment_filename)
#             mail.Send()
#             print("Email envoyé avec succès")
#             return True
#         else:
#             print("Erreur : Fichier joint introuvable")
#             return False
            
#     except Exception as e:
#         print(f"Erreur lors de l'envoi de l'email : {e}")
#         return False

# # ------------------ Workflow principal ------------------

# def main():
#     """Workflow principal"""
#     print("Début du traitement des données WIMOVA...")
    
#     # Récupération des données
#     df = fetch_survey_data()
#     if df is None or df.empty:
#         print("Aucune donnée récupérée ou DataFrame vide")
#         return
    
#     # Transformation des données
#     df = combine_date_time(df)
#     df = normalize_json_column(df, "data")
    
#     # Vérifier quelles colonnes sont disponibles pour le debug
#     print("Colonnes disponibles dans le DataFrame:")
#     for col in df.columns:
#         print(f"  - {col}")
    
#     # Fusion des colonnes selon la nouvelle structure
#     # Pour la colonne Solution: Basée sur les CD L2 pour modification de course
#     df['Solution'] = df.apply(
#         lambda row: merge_columns_with_conflict(
#             row, [
#                 'CD L2 modification de course pour prestataire appel sortant',
#                 'CD L2 pour modification de course Email',
#                 'CD L2 modification de course pour passager appel entrant',
#                 'CD L2 modification course pour client appel entrant'
#             ]
#         ), axis=1
#     )
    
#     # Pour la colonne Interlocuteur: Basée sur les interlocuteurs
#     df['Interlocuteur'] = df.apply(
#         lambda row: merge_columns_with_conflict(
#             row, ['Interlocuteur appel entrant', 'Interlocuteur appel sortant']
#         ), axis=1
#     )
    
#     # Suppression des anciennes colonnes qui n'existent plus
#     # On garde seulement celles qui existent vraiment
#     columns_to_drop = []
#     old_columns = ['Solution AE', 'Solution AS', 'Email Solution', 
#                    'AE Interlocuteur', 'AS Interlocuteur']
    
#     for col in old_columns:
#         if col in df.columns:
#             columns_to_drop.append(col)
    
#     if columns_to_drop:
#         df.drop(columns=columns_to_drop, inplace=True)
#         print(f"Anciennes colonnes supprimées : {columns_to_drop}")
    
#     # Fusion des colonnes par préfixe - ajuster selon la nouvelle structure
#     # Contact Driver L1 pour la nouvelle structure
#     df = merge_multiple_columns(df, "CD L1", "Contact Driver L1")
    
#     # CD L2 pour la nouvelle structure
#     df = merge_multiple_columns(df, "CD L2", "CD L2")
    
#     # Export Excel
#     if export_to_excel(df, EXPORT_PATH):
#         print("Export Excel terminé avec succès")
    
#     # Envoi par email
#     temp_folder = os.path.join(os.getenv('TEMP'), 'survey_attachments')
#     os.makedirs(temp_folder, exist_ok=True)
#     end_date = (date.today() - timedelta(days=1)).strftime("%Y%m%d")
#     temp_file = os.path.join(temp_folder, f"WIMOVA_Extract_Post_Prod_{end_date}.xlsx")
    
#     if save_excel_temp(df, 'Batonnage', temp_file):
#         if send_email(df, temp_file):
#             # Nettoyage du fichier temporaire
#             if os.path.exists(temp_file):
#                 os.remove(temp_file)
#                 print("Fichier temporaire nettoyé")
#     else:
#         print("Échec de la préparation de l'email")
    
#     print("Traitement terminé")

# if __name__ == "__main__":
#     main()









#####################################################################################"
# ################################################################################"
# ##############################################################################
# #####################################################################"




# Standard library
import os
import sys
import json
import warnings
import base64
from datetime import datetime, date, timedelta
from io import BytesIO

# Third-party
import pandas as pd
import requests
import msal
from sshtunnel import SSHTunnelForwarder
import mysql.connector
import win32com.client as win32
from dotenv import load_dotenv
env_path = r"C:\Users\Administrateur\Desktop\WIMOVA\Post Prod_Wimova/.env"
load_dotenv(dotenv_path=env_path)

# Configuration
warnings.simplefilter("ignore", category=UserWarning)
sys.stdout.reconfigure(encoding='utf-8')


# ============================================================
# ========================= CONFIG ===========================
# ============================================================

# ---------- MODE ----------
INITIAL_LOAD = False  # True => overwrite total même si fichier existe

# pylint: disable=duplicate-code
# ---------- SSH ----------
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

TABLE_NAME = os.getenv("TABLE_NAME")
FORM_ID_VALUE = os.getenv("FORM_ID_VALUE")


# Colonne double-encodée
RESPONSE_DATA_COL = "response_data"

# pylint: disable=duplicate-code
# ---------- SHAREPOINT ----------
TENANT_ID     = os.getenv("TENANT_ID")
CLIENT_ID     = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
SITE_DOMAIN = os.getenv("SITE_DOMAIN")
SITE_NAME     = os.getenv("SITE_NAME")     # ex: SIBTEL
DRIVE_NAME = "Documents"

SHAREPOINT_FILE_PATH = "General/WIMOVA/raw_data/surveydatabatonnage.xlsx"
TARGET_TABLE = "wimova_batonnage"

SHAREPOINT_FILE_PATH = "General/WIMOVA/raw_data/surveydatabatonnage.xlsx"

LOGO_PATH = r"C:\Users\Administrateur\Desktop\WIMOVA\Post Prod_Wimova\logo-complet-color.png"
EXPORT_PATH = r"C:\Users\Administrateur\Desktop\WIMOVA\Post Prod_Wimova\downloads\batonnage_wimova_1.xlsx"

SHAREPOINT_FILE_LOG_PATH ="General/Autoreports Status"
# Log file name
log_filename = f"Post_Prod_wimova_{datetime.now().strftime('%Y%m%d_%H.%M')}.txt"

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
# ============================================================
# ========================= UTILS =============================
# ============================================================

def add_data():
    def require(v, name):
        if v is None or str(v).strip() == "":
            raise ValueError(f"❌ Missing config: {name}")
        return str(v).strip()


    def normalize_key(s: str) -> str:
        s = str(s).strip().lower()
        s = unicodedata.normalize("NFKD", s)
        s = "".join(ch for ch in s if not unicodedata.combining(ch))
        s = re.sub(r"[^a-z0-9]+", "", s)
        return s


    def pick_key_column(cols):
        """
        Colonne unique pour append-only.
        Priorité selon tes conventions.
        """
        candidates = [
            "inboundformresponse beside id",
            "inboundformresponse_beside_id",
            "formresponse beside id",
            "formresponse_beside_id",
            "beside id",
            "beside_id",
            "id",
        ]
        norm = {normalize_key(c): c for c in cols}
        for cand in candidates:
            k = normalize_key(cand)
            if k in norm:
                return norm[k]
        return None


    def to_key_series(df: pd.DataFrame, key_col: str) -> pd.Series:
        return df[key_col].astype("string").fillna(pd.NA)


    def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            df.to_excel(w, index=False)
        return buf.getvalue()



    # ============================================================
    # ========= DOUBLE-ENCODED JSON (response_data) ==============
    # ============================================================
    def _safe_json_loads(s: str):
        try:
            return json.loads(s)
        except Exception:
            return None


    def decode_double_encoded_json(value):
        """
        Cas typique:
        value = "\"{\\\"a\\\":1,\\\"b\\\":\\\"x\\\"}\""
        -> dict {"a":1,"b":"x"}
        """
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return None
        if isinstance(value, (dict, list)):
            return value
        if isinstance(value, bytes):
            value = value.decode("utf-8", "ignore")

        s = str(value).strip()
        if not s:
            return None

        obj1 = _safe_json_loads(s)
        if obj1 is None:
            return None

        if isinstance(obj1, str):
            obj2 = _safe_json_loads(obj1)
            return obj2 if obj2 is not None else obj1

        return obj1


    def flatten_response_data_column(df: pd.DataFrame, colname: str) -> pd.DataFrame:
        
        if df is None or df.empty:
            return df

        wanted = normalize_key(colname)
        real_col = None
        for c in df.columns:
            if normalize_key(c) == wanted:
                real_col = c
                break

        if not real_col:
            write_log(f"Colonne '{colname}' non trouvée => skip.")
            return df

        decoded = df[real_col].apply(decode_double_encoded_json)

        if not decoded.dropna().apply(lambda x: isinstance(x, (dict, list))).any():
            write_log(f"'{real_col}' non JSON décodable => skip flatten.")
            return df

        # dict -> expand (sans prefixe)
        if decoded.dropna().apply(lambda x: isinstance(x, dict)).any():
            expanded = pd.json_normalize(
                decoded.where(decoded.apply(lambda x: isinstance(x, dict)))
            )

            # collisions : si le champ existe déjà dans df, on renomme en *_json
            expanded.columns = [
                c if c not in df.columns else f"{c}_json"
                for c in expanded.columns
            ]

            df = pd.concat([df.drop(columns=[real_col]), expanded], axis=1)

        else:
            # list only => stringify
            df[real_col] = decoded.apply(
                lambda x: json.dumps(x, ensure_ascii=False) if isinstance(x, list) else x
            )

        return df


    # ============================================================
    # ===================== SHAREPOINT API =======================
    # ============================================================
    def authenticate_sharepoint():
        require(TENANT_ID, "TENANT_ID")
        require(CLIENT_ID, "CLIENT_ID")
        require(CLIENT_SECRET, "CLIENT_SECRET")
        require(SITE_DOMAIN, "SITE_DOMAIN")
        require(SITE_NAME, "SITE_NAME")

        app = msal.ConfidentialClientApplication(
            CLIENT_ID,
            authority=f"https://login.microsoftonline.com/{TENANT_ID}",
            client_credential=CLIENT_SECRET,
        )
        token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        if "access_token" not in token:
            raise Exception(f" Token error: {token}")

        return {"Authorization": f"Bearer {token['access_token']}"}


    def get_site_id(headers):
        url = f"https://graph.microsoft.com/v1.0/sites/{SITE_DOMAIN}:/sites/{SITE_NAME}"
        r = requests.get(url, headers=headers).json()
        if "id" not in r:
            raise Exception(f" Site error: {r}")
        return r["id"]


    def get_drive_id(headers, site_id):
        r = requests.get(f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives", headers=headers).json()
        drives = r.get("value", [])
        wanted = DRIVE_NAME.strip().lower()

        for d in drives:
            if (d.get("name") or "").strip().lower() == wanted:
                return d["id"]
        for d in drives:
            if (d.get("name") or "").strip().lower() == "documents":
                return d["id"]

        raise Exception("Drive not found")


    def download_sharepoint_file_bytes(headers, drive_id, sp_path: str):
        require(sp_path, "SHAREPOINT_FILE_PATH")
        path = "/" + sp_path.strip().strip("/")
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:{path}:/content"
        r = requests.get(url, headers=headers)
        if r.status_code == 404:
            return None
        if r.status_code != 200:
            raise Exception(f" Download failed: {r.status_code} {r.text}")
        return r.content


    def upload_sharepoint_file_bytes(headers, drive_id, sp_path: str, content: bytes):
        path = "/" + sp_path.strip().strip("/")
        write_log(f"Upload SharePoint overwrite: {path}")
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:{path}:/content"
        r = requests.put(url, headers=headers, data=content)
        if r.status_code not in (200, 201):
            raise Exception(f" Upload failed: {r.status_code} {r.text}")
        write_log("Uploaded to SharePoint")


    # ============================================================
    # ========================== DB ===============================
    # ============================================================
    def open_ssh_tunnel():
        require(SSH_HOST, "SSH_HOST")
        require(SSH_USER, "SSH_USER")
        require(SSH_PASSWORD, "SSH_PASSWORD")
        require(DB_USER, "DB_USER")
        require(DB_PASSWORD, "DB_PASSWORD")
        require(DB_NAME, "DB_NAME")

        t = SSHTunnelForwarder(
            (SSH_HOST, SSH_PORT),
            ssh_username=SSH_USER,
            ssh_password=SSH_PASSWORD,
            remote_bind_address=(REMOTE_MYSQL_HOST, REMOTE_MYSQL_PORT),
            local_bind_address=("127.0.0.1", 0),
        )
        t.start()
        write_log(f"SSH tunnel: 127.0.0.1:{t.local_bind_port} -> {REMOTE_MYSQL_HOST}:{REMOTE_MYSQL_PORT}")
        return t


    def export_table_filtered() -> pd.DataFrame:
        write_log(f"Export DB: {TABLE_NAME} WHERE form_id={FORM_ID_VALUE}")
        tunnel = None
        conn = None
        try:
            tunnel = open_ssh_tunnel()
            conn = mysql.connector.connect(
                host="127.0.0.1",
                port=int(tunnel.local_bind_port),
                user=DB_USER,
                password=DB_PASSWORD,
                database=DB_NAME,
                use_pure=True,
                buffered=True,
            )
            cur = conn.cursor(dictionary=True)

            # IMPORTANT: on suppose la colonne s'appelle form_id
            cur.execute(f"SELECT s.*, e.name as user_name FROM {TABLE_NAME} s LEFT JOIN besidedb.myapp_customuser e ON s.user_id = e.id  WHERE s.form_id = %s ", (FORM_ID_VALUE,))
            rows = cur.fetchall()
            cur.close()

            write_log(f"Rows fetched: {len(rows)}")
            df = pd.DataFrame(rows)

            # ✅ response_data double encodé -> colonnes directes (sans prefixe)
            df = flatten_response_data_column(df, RESPONSE_DATA_COL)

            return df

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
    # ========================== MAIN =============================
    # ============================================================
    def main():
        global headers,drive_id
        
        try:
            # SharePoint init
            headers = authenticate_sharepoint()
            site_id = get_site_id(headers)
            drive_id = get_drive_id(headers, site_id)

            # 1) Export DB d'abord
            df_db = export_table_filtered()
            if df_db.empty:
                write_log("DB vide => rien à faire.")
                return

            # 2) Lire fichier existant SharePoint (si existe)
            existing_bytes = download_sharepoint_file_bytes(headers, drive_id, SHAREPOINT_FILE_PATH)

            # ---- IMPORT TOTAL (overwrite) ----
            if INITIAL_LOAD or existing_bytes is None:
                if INITIAL_LOAD:
                    write_log("INITIAL_LOAD=True => import total (overwrite).")
                else:
                    write_log("Fichier SharePoint introuvable => création + import total.")

                upload_sharepoint_file_bytes(headers, drive_id, SHAREPOINT_FILE_PATH, df_to_excel_bytes(df_db))
                
                write_log("DONE — import total + stockage SharePoint ")
                return

            # ---- APPEND ONLY ----
            try:
                df_existing = pd.read_excel(BytesIO(existing_bytes), dtype="string")
            except Exception as e:
                write_log(f"Fichier SharePoint illisible ({e}) => import total.")
                upload_sharepoint_file_bytes(headers, drive_id, SHAREPOINT_FILE_PATH, df_to_excel_bytes(df_db))
                
                return

            if df_existing is None or df_existing.shape[1] == 0:
                write_log("Fichier SharePoint vide/sans colonnes => import total.")
                upload_sharepoint_file_bytes(headers, drive_id, SHAREPOINT_FILE_PATH, df_to_excel_bytes(df_db))
                
                return

            key_col = pick_key_column(df_existing.columns.tolist())
            if (not key_col) or (key_col not in df_db.columns):
                write_log("Colonne clé absente (fichier ou DB) => import total.")
                upload_sharepoint_file_bytes(headers, drive_id, SHAREPOINT_FILE_PATH, df_to_excel_bytes(df_db))
                
                return

            existing_keys = set(to_key_series(df_existing, key_col).dropna().astype(str).tolist())
            db_keys = to_key_series(df_db, key_col).astype("string")

            df_new = df_db[~db_keys.isin(list(existing_keys))].copy()
            write_log(f"New rows to append: {len(df_new)}")

            if df_new.empty:
                write_log("Déjà à jour.")
                
                return

            df_out = pd.concat([df_existing, df_new], ignore_index=True)

            # Stockage 2 emplacements
            upload_sharepoint_file_bytes(headers, drive_id, SHAREPOINT_FILE_PATH, df_to_excel_bytes(df_out))
            

            write_log("DONE — append only + response_data décodé (sans prefixe) + SharePoint ")
        except Exception as e:
            write_log(f"Erreur dans main_add_data(): {e}")
   
   
   
    main()

def main_send_mail():
    def authenticate_sharepoint():
        app = msal.ConfidentialClientApplication(
            CLIENT_ID,
            authority=f"https://login.microsoftonline.com/{TENANT_ID}",
            client_credential=CLIENT_SECRET,
        )
        token = app.acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"]
        )

        return {"Authorization": f"Bearer {token['access_token']}"}


    def get_site_id(headers):
        url = f"https://graph.microsoft.com/v1.0/sites/{SITE_DOMAIN}:/sites/{SITE_NAME}"
        r = requests.get(url, headers=headers).json()
        return r["id"]


    def get_drive_id(headers, site_id):
        r = requests.get(
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives",
            headers=headers,
        ).json()

        for d in r["value"]:
            if d["name"].lower() == DRIVE_NAME.lower():
                return d["id"]

        raise Exception("Drive not found")


    def download_sharepoint_excel():
        headers = authenticate_sharepoint()
        site_id = get_site_id(headers)
        drive_id = get_drive_id(headers, site_id)

        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{SHAREPOINT_FILE_PATH}:/content"
        r = requests.get(url, headers=headers)

        if r.status_code != 200:
            raise Exception("Erreur téléchargement SharePoint")

        return pd.read_excel(BytesIO(r.content))

    

    def merge_columns_with_conflict(row, columns):
        """Fusionne plusieurs colonnes en gérant les conflits"""
        values = []
        for col in columns:
            if col in row.index and pd.notnull(row[col]) and str(row[col]).strip() != '':
                values.append(str(row[col]))
        
        if not values:
            return None
        if len(set(values)) > 1:
            return f" Conflit: {', '.join(values)}"
        return values[0]

    def merge_multiple_columns(df, prefix, new_col_name):
        """Fusionne toutes les colonnes qui commencent par un certain préfixe"""
        try:
            cols_to_merge = [col for col in df.columns if col.startswith(prefix)]
            if cols_to_merge:
                df[new_col_name] = df.apply(
                    lambda row: ', '.join(
                        str(row[col]) for col in cols_to_merge 
                        if pd.notnull(row[col]) and str(row[col]).strip() != ''
                    ), 
                    axis=1
                )
                df.drop(columns=cols_to_merge, inplace=True)
                write_log(f"Colonnes avec préfixe '{prefix}' fusionnées en '{new_col_name}'")
            return df
        except Exception as e:
            write_log(f"Erreur lors de la fusion des colonnes {prefix} : {e}")
            return df

    # ------------------ Fonctions d'export ------------------

    def export_to_excel(df, filename):
        """Exporte le DataFrame en fichier Excel"""
        try:
            # Créer le répertoire s'il n'existe pas
            os.makedirs(os.path.dirname(filename), exist_ok=True)
            
            df.to_excel(filename, index=False, engine="openpyxl")
            write_log(f"Export réussi : {filename}")
            return True
        except Exception as e:
            write_log(f"Erreur lors de l'export en Excel : {e}")
            return False

    def save_excel_temp(df, sheet_name, temp_filename):
        """Sauvegarde le DataFrame en Excel temporaire pour l'email"""
        try:
            if df.empty:
                write_log("DataFrame vide - aucun fichier créé")
                return False
            
            # Créer le répertoire temporaire
            os.makedirs(os.path.dirname(temp_filename), exist_ok=True)
            
            df.to_excel(temp_filename, sheet_name=sheet_name, index=False, engine='openpyxl')
            write_log(f"Fichier temporaire créé : {temp_filename}")
            return True
        except Exception as e:
            write_log(f"Erreur lors de la création du fichier temporaire : {e}")
            return False

    # ------------------ Fonctions d'email ------------------

    def send_email(df_final, attachment_filename):
        """Envoie le DataFrame par email via Outlook"""
        try:
            start_date = "16-12-2025"
            end_date = (date.today() - timedelta(days=1)).strftime("%d-%m-%Y")

            outlook = win32.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)

            mail.Subject = f'WIMOVA-Extract post Production du {start_date} au {end_date}'
            LOGO_PATH = r"C:\Users\Administrateur\Desktop\WIMOVA\Post Prod_Wimova\logo-complet-color.png"

            # Ajouter le logo en pièce jointe intégrée
            if os.path.exists(LOGO_PATH):
                logo_attachment = mail.Attachments.Add(LOGO_PATH)
                logo_attachment.PropertyAccessor.SetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x3712001F", 
                    "logo1.png"
                )

                with open(LOGO_PATH, "rb") as image_file:
                    encoded_image = base64.b64encode(image_file.read()).decode('utf-8')
            else:
                encoded_image = ""
                write_log("Avertissement : Logo introuvable")

            mail.HTMLBody = f'''
                <html>
                <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
                    <div style="max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #ddd; border-radius: 8px; background-color: #f9f9f9;">
                        <div style="position: absolute; top: 20px; right: 20px;">
                            <img src="cid:logo1.png" alt="Logo" style="height: 50px;">
                        </div>
                        <p>Bonjour,</p>
                        <p>Veuillez trouver ci-joint l'extract post Production de l'activité WIMOVA pour la période du <strong>{start_date}</strong> au <strong>{end_date}</strong>.</p>
                        
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


            mail.Sender = "rd22z@ningen-group.com"
            mail.To = "Ningen-pperformance@ningen-group.com"
            mail.CC = "sl06h@ningen-group.com;td75s@ningen-group.com;iz55x@ningen-group.com;Ningen-Data-Management@ningen-group.com;dn87g@ningen-group.com;qq13f@ningen-group.com"
            #mail.To = "ix41p@ningen-group.com"
            # Ajouter la pièce jointe
            if os.path.exists(attachment_filename):
                mail.Attachments.Add(attachment_filename)
                mail.Send()
                write_log("Email envoyé avec succès")
                return True
            else:
                write_log("Erreur : Fichier joint introuvable")
                return False
                
        except Exception as e:
            write_log(f"Erreur lors de l'envoi de l'email : {e}")
            return False

    # ------------------ Workflow principal ------------------

    def main():
        global headers,drive_id
        try:
            """Workflow principal"""
            write_log("Début du traitement des données WIMOVA...")
            

            
            df = download_sharepoint_excel()
            
            # Vérifier quelles colonnes sont disponibles pour le debug
            write_log(f"Données chargées depuis SharePoint : {len(df)} lignes")
            
            
            # Fusion des colonnes selon la nouvelle structure
            # Pour la colonne Solution: Basée sur les CD L2 pour modification de course
            df['Solution'] = df.apply(
                lambda row: merge_columns_with_conflict(
                    row, [
                        'CD L2 modification de course pour prestataire appel sortant',
                        'CD L2 pour modification de course Email',
                        'CD L2 modification de course pour passager appel entrant',
                        'CD L2 modification course pour client appel entrant'
                    ]
                ), axis=1
            )
            
            # Pour la colonne Interlocuteur: Basée sur les interlocuteurs
            df['Interlocuteur'] = df.apply(
                lambda row: merge_columns_with_conflict(
                    row, ['Interlocuteur appel entrant', 'Interlocuteur appel sortant']
                ), axis=1
            )
            
            # Suppression des anciennes colonnes qui n'existent plus
            # On garde seulement celles qui existent vraiment
            columns_to_drop = []
            old_columns = ['Solution AE', 'Solution AS', 'Email Solution', 
                        'AE Interlocuteur', 'AS Interlocuteur']
            
            for col in old_columns:
                if col in df.columns:
                    columns_to_drop.append(col)
            
            if columns_to_drop:
                df.drop(columns=columns_to_drop, inplace=True)
                write_log(f"Anciennes colonnes supprimées : {columns_to_drop}")
            
            # Fusion des colonnes par préfixe - ajuster selon la nouvelle structure
            # Contact Driver L1 pour la nouvelle structure
            df = merge_multiple_columns(df, "CD L1", "Contact Driver L1")
            
            # CD L2 pour la nouvelle structure
            df = merge_multiple_columns(df, "CD L2", "CD L2")
            
            final_columns = [
                "id",
                "survey",
                "created_at",
                "user_id",
                
                
                
                "user_name",
                "KM Approche",
                "# Tentatives",
                "N° de Mission",
                "Typologie de Clients",
                "Typologie de contact",
                "Interlocuteur appel entrant",
                "Interlocuteur appel sortant",
                "Traitement",
                "Commentaire",
                "Solution",
                "Interlocuteur",
                "Contact Driver L1",
                "CD L2"
            ]

            df = df[[col for col in final_columns if col in df.columns]]
            
            # Export Excel
            if export_to_excel(df, EXPORT_PATH):
                write_log("Export Excel terminé avec succès")
            
            # Envoi par email
            temp_folder = os.path.join(os.getenv('TEMP'), 'survey_attachments')
            os.makedirs(temp_folder, exist_ok=True)
            end_date = (date.today() - timedelta(days=1)).strftime("%Y%m%d")
            temp_file = os.path.join(temp_folder, f"WIMOVA_Extract_Post_Prod_{end_date}.xlsx")
            
            if save_excel_temp(df, 'Batonnage', temp_file):
                if send_email(df, temp_file):
                    # Nettoyage du fichier temporaire
                    if os.path.exists(temp_file):
                        os.remove(temp_file)
                        write_log("Fichier temporaire nettoyé")
            else:
                write_log("Échec de la préparation de l'email")
            
            write_log("Traitement terminé")
        
        except Exception as e:
            write_log(f"Erreur in sending mail: {e}")
        
    main()





if __name__ == "__main__":
    
    try:
        add_data()
        main_send_mail()
        
    except Exception as e:
        write_log(f"Erreur dans main(): {e}")










