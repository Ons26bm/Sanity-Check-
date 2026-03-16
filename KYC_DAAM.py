# import pandas as pd
# from datetime import datetime, timedelta
# from sqlalchemy import create_engine, text
# from sshtunnel import SSHTunnelForwarder
# import json
# import os
# import traceback
# import base64
# import win32com.client as win32
# import zipfile

# # ------------------ CONFIG SSH ------------------
# ssh_host = '169.255.70.60'
# ssh_port = 22
# ssh_user = 'JU97W'
# ssh_password = 'RamyHashtag@1989'

# # ------------------ CONFIG DB ------------------
# db_host = 'localhost'
# db_port = 3306
# db_user = 'reportings_user'
# db_password = '#Brochill34'
# db_name = 'besidedb'

# # ------------------ SQL ------------------
# query = """
# SELECT inj.id,
#        inj.data,
#        inj.created_at AS `Date Injection`,
#        rslt.response_data,
#        rslt.created_at AS `Date Traitement`,
#        rslt.form_id,
#        rslt.user_id, 
#        rslt.survey_data_id,
#        rslt.survey_schema_id
# FROM myapp_formresponse AS rslt
# INNER JOIN myapp_surveydata AS inj
#         ON rslt.survey_data_id = inj.id
# WHERE (rslt.survey_schema_id = 17 OR rslt.survey_schema_id = 28)
#   AND DATE(rslt.created_at) >= '2025-08-01'
#   AND JSON_UNQUOTE(JSON_EXTRACT(rslt.response_data, '$.Qualification')) = 'Formulaire complété';
# """
#  #


# # ------------------ FONCTION DE NORMALISATION ------------------
# def normalize_json_columns(df, columns_to_normalize):
#     for column_name in columns_to_normalize:
#         if column_name in df.columns and len(df) > 0:
#             if isinstance(df[column_name].iloc[0], str):
#                 try:
#                     df[column_name] = df[column_name].apply(json.loads)
#                     normalized_df = pd.json_normalize(df[column_name])
#                     suffix = "_response" if column_name == "response_data" else "_data"
#                     normalized_df = normalized_df.add_suffix(suffix)
#                     df = pd.concat([df.drop(column_name, axis=1), normalized_df], axis=1)
#                     print(f"Colonne {column_name} normalisée avec suffixe '{suffix}'")
#                 except Exception as e:
#                     print(f"Erreur normalisation {column_name}: {e}")
#             else:
#                 print(f"La colonne {column_name} ne contient pas de données JSON string")
#         else:
#             print(f"Colonne {column_name} non trouvée ou DataFrame vide")
#     return df

# # ------------------ CREATION ZIP ------------------
# def creer_zip(file_path):
#     try:
#         zip_path = file_path.replace(".xlsx", ".zip")
#         with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
#             zf.write(file_path, arcname=os.path.basename(file_path))
#         return zip_path
#     except Exception as e:
#         print(f"Erreur création ZIP: {e}")
#         return None

# # ------------------ ENVOI EMAIL ------------------
# def envoyer_email(destinataire, cc="", sujet="", corps_html="", piece_jointe=None):
#     try:
#         outlook = win32.Dispatch('Outlook.Application')
#         mail = outlook.CreateItem(0)
#         mail.To = destinataire
#         mail.CC = cc
#         mail.Subject = sujet
#         mail.HTMLBody = corps_html
#         if piece_jointe:
#             mail.Attachments.Add(piece_jointe)
#         mail.Send()
#         print("Email envoyé avec succès.")
#         return True
#     except Exception as e:
#         print(f"Erreur envoi email: {e}")
#         return False

# def envoyer_email_alerte(erreur_msg):
#     try:
#         outlook = win32.DispatchEx('Outlook.Application')
#         mail = outlook.CreateItem(0)
#         mail.Subject = "⚠️ ALERTE - BOT KYC Brutes"
#         mail.To = "ix41p@ningen-group.com "
#         #mail.CC = "fe48v@ningen-group.com;Ningen-pperformance@ningen-group.com;comd@ningen-group.com"
#         mail.Importance = 2
#         html_body = f"""
#         <html>
#         <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 0;">
#             <div style="max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e74c3c; border-radius: 8px; background-color: #fdf2f2; box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);">
#                 <div style="background-color: #e74c3c; color: white; padding: 15px; border-radius: 6px 6px 0 0; margin: -20px -20px 20px -20px;">
#                     <h2 style="margin: 0; font-size: 20px;">⚠️ ALERTE - Dysfonctionnement BOT KYC Brutes</h2>
#                 </div>
                
#                 <div style="margin-bottom: 20px;">
#                     <p style="font-size: 16px; color: #d35400; font-weight: bold;">
#                         Une erreur critique a été détectée lors de l'exécution du script :
#                     </p>
#                     <div style="background-color: #f8d7da; border: 1px solid #f5c6cb; border-radius: 4px; padding: 15px; margin: 15px 0;">
#                         <pre style="font-family: Consolas, Monaco, 'Andale Mono', monospace; font-size: 12px; color: #721c24; margin: 0; white-space: pre-wrap; word-wrap: break-word;">{erreur_msg}</pre>
#                     </div>
#                 </div>
                
#                 <div style="background-color: #f9f9f9; padding: 15px; border-radius: 4px; border: 1px solid #ddd;">
#                     <p style="font-size: 14px; margin: 0; color: #666;">
#                         <strong>Date :</strong> {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}<br>
#                         <strong>Action requise :</strong> Intervention technique nécessaire
#                     </p>
#                 </div>
                
#                 <div style="margin-top: 20px; padding-top: 15px; border-top: 1px solid #eee;">
#                     <p style="font-size: 12px; color: #999; text-align: center;">
#                         Ceci est un email automatique, merci de ne pas y répondre.
#                     </p>
#                 </div>
#             </div>
#         </body>
#         </html>
#         """
#         mail.HTMLBody = html_body
#         mail.Send()
#         print("Email d'alerte HTML envoyé avec succès.")
#     except Exception as e:
#         print(f"Impossible d'envoyer l'email d'alerte: {e}")

# # ------------------ EXECUTION ------------------
# try:
#     with SSHTunnelForwarder(
#         (ssh_host, ssh_port),
#         ssh_username=ssh_user,
#         ssh_password=ssh_password,
#         remote_bind_address=(db_host, db_port)
#     ) as tunnel:
#         local_port = tunnel.local_bind_port
#         engine = create_engine(f'mysql+pymysql://{db_user}:{db_password}@127.0.0.1:{local_port}/{db_name}')
#         with engine.connect() as conn:
#             df = pd.read_sql_query(text(query), conn)

#     print(f"Nombre de lignes récupérées : {len(df)}")

#     # Normalisation
#     df = normalize_json_columns(df, ['data', 'response_data'])
#     df.columns = df.columns.str.replace(r'(_data|_response)$', '', regex=True)


#     # Colonnes à conserver
#     KYC_to_keep = [
#         "Date Injection","Date Traitement", "survey_data_id", "Age", "Canal", "Source", "Secteur",
#         "phone number", "phone number 2", "Nb de relances suite au 1er contact",
#         "Nom", "Prénom", "Genre", "Date de naissance", "Nationalité", "Statut marital",
#         "Nom époux (se)", "Nombre d'enfants", "Carte ID", "Niveau d'étude",
#         "Adresse Professionnelle", "Code postal", "Adresse personnelle",
#         "Code postal (adresse personnelle)", "Nature", "Téléphone bureau",
#         "Profession exercée", "Matricule fiscale", "Employeur", "Nombre d'emplois à créer",
#         "Activité de l'entreprise", "Sous-secteurs", "En affaire depuis", "Nom bénéficiaire effectif",
#         "Prénom bénéficiaire effectif", "Care ID bénéficiaire effectif",
#         "Où avez-vous entendu parler de DAAM (source de sollicitation)?",
#         "Exercez-vous ou avez-vous exercé une fonction politique, juridictionnelle ou administrative importante?",
#         "Fonction exercée", "Dans quel pays", "Date fin de mission",
#         "Avez-vous une PPE qui est membre de votre famille ou une personne avec laquelle vous êtes en étroite relation d'affaire?",
#         "Nom et prénom de la personne", "Lien de parenté", "Fonction exercée du membre",
#         "Membre dans quel pays", "Date de la fin de mission du membre", "Objet du crédit",
#         "Type de produit DAAM", "Montant demandé", "Durée demandée",
#         "Date du versement mensuel souhaité (Le combien de chaque mois*)",
#         "Mensualité souhaitée", "Garantie(s) offerte(s)", "Valeur estimée garantie (s)",
#         "Date délivrance Carte ID", "Agency", "Gouvernaurat", "Délégation",
#         "Nombre d'employées", "Type de projet"  ,"user_id"
#     ]
#     KYC_existing = [col for col in KYC_to_keep if col in df.columns]
#     df = df.loc[:, ~df.columns.duplicated()].loc[:, KYC_existing]
#     df["Etat KYC"] = ""
#     df["Raison Rejet"] = ""
#     df["Infos à compléter"] = ""
#     # Nettoyage des montants
#     df['Montant demandé'] = df['Montant demandé'].apply(lambda val: ''.join(filter(str.isdigit, str(val))) if pd.notnull(val) else None)
#     fields = [
#     "survey_data_id",
#     "Date Traitement",
#     "Date Injection",
#     "Nom",
#     "Prénom",
#     "Genre",
#     "Date de naissance",
#     "Nationalité",
#     "Statut marital",
#     "Nom époux (se)",
#     "Nombre d'enfants",
#     "Carte ID",
#     "Date délivrance Carte ID",
#     "phone number",
#     "phone number 2",
#     "Gouvernaurat",
#     "Délégation",
#     "Agency",
#     "Adresse personnelle",
#     "Code postal (adresse personnelle)",
#     "Adresse Professionnelle",
#     "Code postal",
#     "Téléphone bureau",
#     "Profession exercée",
#     "Employeur",
#     "Activité de l'entreprise",
#     "Matricule fiscale",
#     "Nombre d'employées",
#     "Nombre d'emplois à créer",
#     "Sous-secteurs",
#     "En affaire depuis",
#     "Age",
#     "Secteur",
#     "Niveau d'étude",
#     "Nature",
#     "Canal",
#     "Source",
#     "Nb de relances suite au 1er contact",
#     "Date fin de mission",
#     "Care ID bénéficiaire effectif",
#     "Nom bénéficiaire effectif",
#     "Prénom bénéficiaire effectif",
#     "Exercez-vous ou avez-vous exercé une fonction politique, juridictionnelle ou administrative importante?",
#     "Fonction exercée",
#     "Dans quel pays",
#     "Avez-vous une PPE qui est membre de votre famille ou une personne avec laquelle vous êtes en étroite relation d'affaire?",
#     "Nom et prénom de la personne",
#     "Lien de parenté",
#     "Fonction exercée du membre",
#     "Membre dans quel pays",
#     "Date de la fin de mission du membre",
#     "Type de produit DAAM",
#     "Type de projet",
#     "Objet du crédit",
#     "Montant demandé",
#     "Durée demandée",
#     "Date du versement mensuel souhaité (Le combien de chaque mois*)",
#     "Mensualité souhaitée",
#     "Garantie(s) offerte(s)",
#     "Valeur estimée garantie (s)",
#     "Où avez-vous entendu parler de DAAM (source de sollicitation)?",
#     "Etat KYC",
#     "Raison Rejet",
#     "Infos à compléter",
#     "user_id"
# ]

#     existing_fields = [col for col in fields if col in df.columns]

#     # Réordonner le DataFrame
#     df = df[existing_fields]
#     df.rename(columns={'Gouvernaurat': 'Gouvernorat', 'Agency': 'Agence'}, inplace=True)


#     # Sauvegarde temporaire
#     temp_folder = os.path.join(os.getenv('TEMP'), 'kyc_attachments')
#     os.makedirs(temp_folder, exist_ok=True)
#     kyc_temp_file = os.path.join(temp_folder, 'KYC_Brutes.xlsx')

#     # Création Excel sans protection de feuille
#     with pd.ExcelWriter(kyc_temp_file, engine='xlsxwriter') as writer:
#         df.to_excel(writer, index=False, sheet_name='KYC')
#         worksheet = writer.sheets['KYC']

#         # Validation de données pour "Etat KYC"
#         col_idx = df.columns.get_loc("Etat KYC")
#         if len(df) > 0:
#             worksheet.data_validation(
#                 1, col_idx, len(df), col_idx,
#                 {'validate': 'list', 'source': ['Conforme', 'A Compléter', 'Rejetée']}
#             )

#     # Application du mot de passe pour ouvrir le fichier
#     try:
#         excel = win32.Dispatch('Excel.Application')
#         excel.Visible = False
#         excel.DisplayAlerts = False
        
#         wb = excel.Workbooks.Open(kyc_temp_file)
#         wb.SaveAs(kyc_temp_file, Password='DAAM@2025')
#         wb.Close()
#         excel.Quit()
        
#     except Exception as e:
#         print(f"Erreur lors de la protection du fichier: {e}")

#     # ZIP et envoi
#     zip_path = creer_zip(kyc_temp_file)
#     if zip_path:
#         formatted_start_date = (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y")
#         formatted_end_date = datetime.now().strftime("%d/%m/%Y")
#         logo_path = r"C:\Users\Administrateur\Desktop\WIMOVA\Post Prod_Wimova\logo-complet-color.png"
#         with open(logo_path, "rb") as img_file:
#             encoded_image = base64.b64encode(img_file.read()).decode("utf-8")
        
#         corps_email = f'''
#         <html>
#         <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
#             <div style="max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #ddd; border-radius: 8px; background-color: #f9f9f9;">
#                 <div style="position: relative; margin-bottom: 40px;">
#                     <img src="data:image/png;base64,{encoded_image}" alt="Logo" style="height:50px; width:100px; object-fit:contain; position:absolute; right:0; top:0;">
#                 </div>
#                 <p>Bonjour,</p>
#                 <p>Veuillez trouver en pièce jointe les fiches KYC Brutes complétées pour la période du
#                 <strong>{formatted_start_date}</strong> au <strong>{formatted_end_date}</strong>.</p>
#                 <ul style="list-style-type:none; padding:0;">
#                     <li style="margin-bottom:10px; padding:10px; background:#fff; border-radius:4px; border:1px solid #ddd;">
#                         <span style="font-weight:bold; color:#004986;">KYC Brutes :</span>
#                         <span>{len(df)} Fiches</span>
#                     </li>
#                 </ul>
#                 <p><strong>NINGEN Data Analytics</strong><br></p>
                    
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
#             </div>
#         </body>
#         </html>'''

#         envoyer_email(
#             #destinataire="ix41p@ningen-group.com",
#             #destinataire="amine.cherni@daam.tn;Ameni.Drissi@daam.tn",
#             #cc="sl06h@ningen-group.com ;td75s@ningen-group.com;aymen.ouertani@daam.tn;Ningen-Data-Management@ningen-group.com;Ningen-pperformance@ningen-group.com",
#             sujet=f"[Confidentiel] DAAM - Fiches KYC Brutes du {formatted_start_date} au {formatted_end_date}",
#             corps_html=corps_email,
#             piece_jointe=zip_path
#         )
#     else:
#         envoyer_email_alerte("Échec de la création de l'archive ZIP.")

# except Exception:
#     erreur_trace = traceback.format_exc()
#     print(f"Erreur lors du traitement des données : {erreur_trace}")
#     envoyer_email_alerte(erreur_trace)

# finally:
#     # Nettoyage fichiers temporaires
#     for f in [kyc_temp_file, zip_path if 'zip_path' in locals() else None]:
#         if f and os.path.exists(f):
#             try:
#                 os.remove(f)
#                 print(f"Fichier temporaire supprimé: {f}")
#             except Exception:
#                 pass



#==================================================================================================================================================

#=====================================================================================================================================================

#==========================================================================================================================================
# import pandas as pd
# from datetime import datetime, timedelta
# from sqlalchemy import create_engine, text
# import json
# import os
# import traceback
# import base64
# import win32com.client as win32
# import zipfile
# import warnings
# import sys

# # Configuration
# warnings.simplefilter("ignore", category=UserWarning)
# sys.stdout.reconfigure(encoding='utf-8')

# # ------------------ CONFIG AZURE DB ------------------
# DB_CONFIG = {
#     'host': 'dataserver.mysql.database.azure.com',
#     'port': 3306,
#     'user': 'Admach',
#     'password': 'NiGan10gEn',
#     'database': 'datawarehouse'  # ou le nom de votre base de données
# }

# # ------------------ SQL MODIFIÉ POUR AZURE ------------------
# query = """
# SELECT inj.surveydata_beside_id as id,
#        inj.data,
#        inj.created_at AS `Date Injection`,
#        rslt.response_data,
#        rslt.created_at AS `Date Traitement`,
#        rslt.form_beside_id,
#        rslt.beside_id, 
#        rslt.surveydata_beside_id,
#        rslt.surveyschema_beside_id
# FROM outbound_surveyresults AS rslt
# INNER JOIN outbound_surveydata AS inj
#         ON rslt.surveydata_beside_id = inj.id
# WHERE (rslt.surveyschema_beside_id = 17 OR rslt.surveyschema_beside_id = 28)
#   AND DATE(rslt.created_at) >= '2025-08-01'
#   AND JSON_UNQUOTE(JSON_EXTRACT(rslt.response_data, '$.Qualification')) = 'Formulaire complété';
# """

# # ------------------ FONCTION DE CONNEXION AZURE ------------------
# def connect_db():
#     """Établit une connexion à la base de données Azure"""
#     try:
#         # Création de l'URL de connexion pour SQLAlchemy
#         connection_string = f"mysql+pymysql://{DB_CONFIG['user']}:{DB_CONFIG['password']}@{DB_CONFIG['host']}:{DB_CONFIG['port']}/{DB_CONFIG['database']}"
#         engine = create_engine(connection_string, 
#                              connect_args={'ssl': {'ssl-mode': 'preferred'}})
#         print("Connexion réussie à la base de données Azure")
#         return engine
#     except Exception as e:
#         print(f"Erreur de connexion à la base de données Azure : {e}")
#         return None

# # ------------------ FONCTION DE NORMALISATION ------------------
# def normalize_json_columns(df, columns_to_normalize):
#     for column_name in columns_to_normalize:
#         if column_name in df.columns and len(df) > 0:
#             if isinstance(df[column_name].iloc[0], str):
#                 try:
#                     df[column_name] = df[column_name].apply(json.loads)
#                     normalized_df = pd.json_normalize(df[column_name])
#                     suffix = "_response" if column_name == "response_data" else "_data"
#                     normalized_df = normalized_df.add_suffix(suffix)
#                     df = pd.concat([df.drop(column_name, axis=1), normalized_df], axis=1)
#                     print(f"Colonne {column_name} normalisée avec suffixe '{suffix}'")
#                 except Exception as e:
#                     print(f"Erreur normalisation {column_name}: {e}")
#             else:
#                 print(f"La colonne {column_name} ne contient pas de données JSON string")
#         else:
#             print(f"Colonne {column_name} non trouvée ou DataFrame vide")
#     return df

# # ------------------ CREATION ZIP ------------------
# def creer_zip(file_path):
#     try:
#         zip_path = file_path.replace(".xlsx", ".zip")
#         with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
#             zf.write(file_path, arcname=os.path.basename(file_path))
#         return zip_path
#     except Exception as e:
#         print(f"Erreur création ZIP: {e}")
#         return None

# # ------------------ ENVOI EMAIL ------------------
# def envoyer_email(destinataire, cc="", sujet="", corps_html="", piece_jointe=None):
#     try:
#         outlook = win32.Dispatch('Outlook.Application')
#         mail = outlook.CreateItem(0)
#         mail.To = destinataire
#         mail.CC = cc
#         mail.Subject = sujet
#         mail.HTMLBody = corps_html
#         if piece_jointe:
#             mail.Attachments.Add(piece_jointe)
#         mail.Send()
#         print("Email envoyé avec succès.")
#         return True
#     except Exception as e:
#         print(f"Erreur envoi email: {e}")
#         return False

# def envoyer_email_alerte(erreur_msg):
#     try:
#         outlook = win32.DispatchEx('Outlook.Application')
#         mail = outlook.CreateItem(0)
#         mail.Subject = "⚠️ ALERTE - BOT KYC Brutes"
#         mail.To = "ix41p@ningen-group.com"
#         # mail.CC = "fe48v@ningen-group.com;Ningen-pperformance@ningen-group.com;comd@ningen-group.com"
#         mail.Importance = 2
#         html_body = f"""
#         <html>
#         <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 0;">
#             <div style="max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e74c3c; border-radius: 8px; background-color: #fdf2f2; box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);">
#                 <div style="background-color: #e74c3c; color: white; padding: 15px; border-radius: 6px 6px 0 0; margin: -20px -20px 20px -20px;">
#                     <h2 style="margin: 0; font-size: 20px;">⚠️ ALERTE - Dysfonctionnement BOT KYC Brutes</h2>
#                 </div>
                
#                 <div style="margin-bottom: 20px;">
#                     <p style="font-size: 16px; color: #d35400; font-weight: bold;">
#                         Une erreur critique a été détectée lors de l'exécution du script :
#                     </p>
#                     <div style="background-color: #f8d7da; border: 1px solid #f5c6cb; border-radius: 4px; padding: 15px; margin: 15px 0;">
#                         <pre style="font-family: Consolas, Monaco, 'Andale Mono', monospace; font-size: 12px; color: #721c24; margin: 0; white-space: pre-wrap; word-wrap: break-word;">{erreur_msg}</pre>
#                     </div>
#                 </div>
                
#                 <div style="background-color: #f9f9f9; padding: 15px; border-radius: 4px; border: 1px solid #ddd;">
#                     <p style="font-size: 14px; margin: 0; color: #666;">
#                         <strong>Date :</strong> {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}<br>
#                         <strong>Action requise :</strong> Intervention technique nécessaire
#                     </p>
#                 </div>
                
#                 <div style="margin-top: 20px; padding-top: 15px; border-top: 1px solid #eee;">
#                     <p style="font-size: 12px; color: #999; text-align: center;">
#                         Ceci est un email automatique, merci de ne pas y répondre.
#                     </p>
#                 </div>
#             </div>
#         </body>
#         </html>
#         """
#         mail.HTMLBody = html_body
#         mail.Send()
#         print("Email d'alerte HTML envoyé avec succès.")
#     except Exception as e:
#         print(f"Impossible d'envoyer l'email d'alerte: {e}")

# # ------------------ EXECUTION ------------------
# try:
#     # Connexion directe à Azure
#     engine = connect_db()
#     if engine is None:
#         raise Exception("Impossible de se connecter à la base de données Azure")
    
#     with engine.connect() as conn:
#         df = pd.read_sql_query(text(query), conn)

#     print(f"Nombre de lignes récupérées : {len(df)}")

#     # Normalisation
#     df = normalize_json_columns(df, ['data', 'response_data'])
#     df.columns = df.columns.str.replace(r'(_data|_response)$', '', regex=True)

#     # Renommer les colonnes pour correspondre à votre structure existante
#     df.rename(columns={
#         'beside_id': 'user_id',
#         'form_beside_id': 'form_id',
#         'surveydata_beside_id': 'survey_data_id',
#         'surveyschema_beside_id': 'survey_schema_id'
#     }, inplace=True)

#     # Colonnes à conserver
#     KYC_to_keep = [
#         "Date Injection","Date Traitement", "survey_data_id", "Age", "Canal", "Source", "Secteur",
#         "phone number", "phone number 2", "Nb de relances suite au 1er contact",
#         "Nom", "Prénom", "Genre", "Date de naissance", "Nationalité", "Statut marital",
#         "Nom époux (se)", "Nombre d'enfants", "Carte ID", "Niveau d'étude",
#         "Adresse Professionnelle", "Code postal", "Adresse personnelle",
#         "Code postal (adresse personnelle)", "Nature", "Téléphone bureau",
#         "Profession exercée", "Matricule fiscale", "Employeur", "Nombre d'emplois à créer",
#         "Activité de l'entreprise", "Sous-secteurs", "En affaire depuis", "Nom bénéficiaire effectif",
#         "Prénom bénéficiaire effectif", "Care ID bénéficiaire effectif",
#         "Où avez-vous entendu parler de DAAM (source de sollicitation)?",
#         "Exercez-vous ou avez-vous exercé une fonction politique, juridictionnelle ou administrative importante?",
#         "Fonction exercée", "Dans quel pays", "Date fin de mission",
#         "Avez-vous une PPE qui est membre de votre famille ou une personne avec laquelle vous êtes en étroite relation d'affaire?",
#         "Nom et prénom de la personne", "Lien de parenté", "Fonction exercée du membre",
#         "Membre dans quel pays", "Date de la fin de mission du membre", "Objet du crédit",
#         "Type de produit DAAM", "Montant demandé", "Durée demandée",
#         "Date du versement mensuel souhaité (Le combien de chaque mois*)",
#         "Mensualité souhaitée", "Garantie(s) offerte(s)", "Valeur estimée garantie (s)",
#         "Date délivrance Carte ID", "Agency", "Gouvernaurat", "Délégation",
#         "Nombre d'employées", "Type de projet", "user_id"
#     ]
    
#     KYC_existing = [col for col in KYC_to_keep if col in df.columns]
#     df = df.loc[:, ~df.columns.duplicated()].loc[:, KYC_existing]
#     df["Etat KYC"] = ""
#     df["Raison Rejet"] = ""
#     df["Infos à compléter"] = ""
    
#     # Nettoyage des montants
#     df['Montant demandé'] = df['Montant demandé'].apply(lambda val: ''.join(filter(str.isdigit, str(val))) if pd.notnull(val) else None)
    
#     fields = [
#         "survey_data_id",
#         "Date Traitement",
#         "Date Injection",
#         "Nom",
#         "Prénom",
#         "Genre",
#         "Date de naissance",
#         "Nationalité",
#         "Statut marital",
#         "Nom époux (se)",
#         "Nombre d'enfants",
#         "Carte ID",
#         "Date délivrance Carte ID",
#         "phone number",
#         "phone number 2",
#         "Gouvernaurat",
#         "Délégation",
#         "Agency",
#         "Adresse personnelle",
#         "Code postal (adresse personnelle)",
#         "Adresse Professionnelle",
#         "Code postal",
#         "Téléphone bureau",
#         "Profession exercée",
#         "Employeur",
#         "Activité de l'entreprise",
#         "Matricule fiscale",
#         "Nombre d'employées",
#         "Nombre d'emplois à créer",
#         "Sous-secteurs",
#         "En affaire depuis",
#         "Age",
#         "Secteur",
#         "Niveau d'étude",
#         "Nature",
#         "Canal",
#         "Source",
#         "Nb de relances suite au 1er contact",
#         "Date fin de mission",
#         "Care ID bénéficiaire effectif",
#         "Nom bénéficiaire effectif",
#         "Prénom bénéficiaire effectif",
#         "Exercez-vous ou avez-vous exercé une fonction politique, juridictionnelle ou administrative importante?",
#         "Fonction exercée",
#         "Dans quel pays",
#         "Avez-vous une PPE qui est membre de votre famille ou une personne avec laquelle vous êtes en étroite relation d'affaire?",
#         "Nom et prénom de la personne",
#         "Lien de parenté",
#         "Fonction exercée du membre",
#         "Membre dans quel pays",
#         "Date de la fin de mission du membre",
#         "Type de produit DAAM",
#         "Type de projet",
#         "Objet du crédit",
#         "Montant demandé",
#         "Durée demandée",
#         "Date du versement mensuel souhaité (Le combien de chaque mois*)",
#         "Mensualité souhaitée",
#         "Garantie(s) offerte(s)",
#         "Valeur estimée garantie (s)",
#         "Où avez-vous entendu parler de DAAM (source de sollicitation)?",
#         "Etat KYC",
#         "Raison Rejet",
#         "Infos à compléter",
#         "user_id"
#     ]

#     existing_fields = [col for col in fields if col in df.columns]

#     # Réordonner le DataFrame
#     df = df[existing_fields]
#     df.rename(columns={'Gouvernaurat': 'Gouvernorat', 'Agency': 'Agence'}, inplace=True)

#     # Sauvegarde temporaire
#     temp_folder = os.path.join(os.getenv('TEMP'), 'kyc_attachments')
#     os.makedirs(temp_folder, exist_ok=True)
#     kyc_temp_file = os.path.join(temp_folder, 'KYC_Brutes.xlsx')

#     # Création Excel sans protection de feuille
#     with pd.ExcelWriter(kyc_temp_file, engine='xlsxwriter') as writer:
#         df.to_excel(writer, index=False, sheet_name='KYC')
#         worksheet = writer.sheets['KYC']

#         # Validation de données pour "Etat KYC"
#         col_idx = df.columns.get_loc("Etat KYC")
#         if len(df) > 0:
#             worksheet.data_validation(
#                 1, col_idx, len(df), col_idx,
#                 {'validate': 'list', 'source': ['Conforme', 'A Compléter', 'Rejetée']}
#             )

#     # Application du mot de passe pour ouvrir le fichier
#     try:
#         excel = win32.Dispatch('Excel.Application')
#         excel.Visible = False
#         excel.DisplayAlerts = False
        
#         wb = excel.Workbooks.Open(kyc_temp_file)
#         wb.SaveAs(kyc_temp_file, Password='DAAM@2025')
#         wb.Close()
#         excel.Quit()
        
#     except Exception as e:
#         print(f"Erreur lors de la protection du fichier: {e}")

#     # ZIP et envoi
#     zip_path = creer_zip(kyc_temp_file)
#     if zip_path:
#         formatted_start_date = (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y")
#         formatted_end_date = datetime.now().strftime("%d/%m/%Y")
#         logo_path = r"C:\Users\Administrateur\Desktop\WIMOVA\Post Prod_Wimova\logo-complet-color.png"
#         with open(logo_path, "rb") as img_file:
#             encoded_image = base64.b64encode(img_file.read()).decode("utf-8")
        
#         corps_email = f'''
#         <html>
#         <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
#             <div style="max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #ddd; border-radius: 8px; background-color: #f9f9f9;">
#                 <div style="position: relative; margin-bottom: 40px;">
#                     <img src="data:image/png;base64,{encoded_image}" alt="Logo" style="height:50px; width:100px; object-fit:contain; position:absolute; right:0; top:0;">
#                 </div>
#                 <p>Bonjour,</p>
#                 <p>Veuillez trouver en pièce jointe les fiches KYC Brutes complétées pour la période du
#                 <strong>{formatted_start_date}</strong> au <strong>{formatted_end_date}</strong>.</p>
#                 <ul style="list-style-type:none; padding:0;">
#                     <li style="margin-bottom:10px; padding:10px; background:#fff; border-radius:4px; border:1px solid #ddd;">
#                         <span style="font-weight:bold; color:#004986;">KYC Brutes :</span>
#                         <span>{len(df)} Fiches</span>
#                     </li>
#                 </ul>
#                 <p><strong>NINGEN Data Analytics</strong><br></p>
                    
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
#             </div>
#         </body>
#         </html>'''

#         envoyer_email(
#             destinataire="ix41p@ningen-group.com",
#             #destinataire="amine.cherni@daam.tn;Ameni.Drissi@daam.tn",
#             #cc="sl06h@ningen-group.com ;td75s@ningen-group.com;aymen.ouertani@daam.tn;Ningen-Data-Management@ningen-group.com;Ningen-pperformance@ningen-group.com",
#             sujet=f"[Confidentiel] DAAM - Fiches KYC Brutes du {formatted_start_date} au {formatted_end_date}",
#             corps_html=corps_email,
#             piece_jointe=zip_path
#         )
#     else:
#         envoyer_email_alerte("Échec de la création de l'archive ZIP.")

# except Exception:
#     erreur_trace = traceback.format_exc()
#     print(f"Erreur lors du traitement des données : {erreur_trace}")
#     envoyer_email_alerte(erreur_trace)

# finally:
#     # Nettoyage fichiers temporaires
#     for f in [kyc_temp_file, zip_path if 'zip_path' in locals() else None]:
#         if f and os.path.exists(f):
#             try:
#                 os.remove(f)
#                 print(f"Fichier temporaire supprimé: {f}")
#             except Exception:
#                 pass



# #==================================================================================================================================================

# #=====================================================================================================================================================

# #==========================================================================================================================================


# import pandas as pd
# from datetime import datetime, timedelta
# from sqlalchemy import create_engine, text
# from sshtunnel import SSHTunnelForwarder
# import json
# import os
# import traceback
# import base64
# import win32com.client as win32
# import zipfile

# # ------------------ CONFIG SSH ------------------
# ssh_host = '169.255.70.60'
# ssh_port = 22
# ssh_user = 'JU97W'
# ssh_password = 'RamyHashtag@1989'

# # ------------------ CONFIG DB ------------------
# db_host = 'localhost'
# db_port = 3306
# db_user = 'reportings_user'
# db_password = '#Brochill34'
# db_name = 'besidedb'

# # Chemin vers le fichier Excel existant
# EXCEL_FILE_PATH = r"C:\Users\IX41P\OneDrive - Ningen Group\Bureau\DAAM_CAF\KYC_Brutes.xlsx"

# # ------------------ SQL ------------------
# query = """
# SELECT inj.id,
#        inj.data,
#        inj.created_at AS `Date Injection`,
#        rslt.response_data,
#        rslt.created_at AS `Date Traitement`,
#        rslt.form_id,
#        rslt.user_id, 
#        rslt.survey_data_id,
#        rslt.survey_schema_id
# FROM myapp_formresponse AS rslt
# INNER JOIN myapp_surveydata AS inj
#         ON rslt.survey_data_id = inj.id
# WHERE rslt.survey_schema_id = 28
#   AND DATE(rslt.created_at) >= '2025-08-01'
#   AND JSON_UNQUOTE(JSON_EXTRACT(rslt.response_data, '$.Qualification')) = 'Formulaire complété';
# """
# # ------------------ FONCTION DE NORMALISATION ------------------
# def normalize_json_columns(df, columns_to_normalize):
#     for column_name in columns_to_normalize:
#         if column_name in df.columns and len(df) > 0:
#             if isinstance(df[column_name].iloc[0], str):
#                 try:
#                     df[column_name] = df[column_name].apply(json.loads)
#                     normalized_df = pd.json_normalize(df[column_name])
#                     suffix = "_response" if column_name == "response_data" else "_data"
#                     normalized_df = normalized_df.add_suffix(suffix)
#                     df = pd.concat([df.drop(column_name, axis=1), normalized_df], axis=1)
#                     print(f"Colonne {column_name} normalisée avec suffixe '{suffix}'")
#                 except Exception as e:
#                     print(f"Erreur normalisation {column_name}: {e}")
#             else:
#                 print(f"La colonne {column_name} ne contient pas de données JSON string")
#         else:
#             print(f"Colonne {column_name} non trouvée ou DataFrame vide")
#     return df

# # ------------------ CREATION ZIP ------------------
# def creer_zip(file_path):
#     try:
#         zip_path = file_path.replace(".xlsx", ".zip")
#         with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
#             zf.write(file_path, arcname=os.path.basename(file_path))
#         return zip_path
#     except Exception as e:
#         print(f"Erreur création ZIP: {e}")
#         return None

# # ------------------ ENVOI EMAIL ------------------
# def envoyer_email(destinataire, cc="", sujet="", corps_html="", piece_jointe=None):
#     try:
#         outlook = win32.Dispatch('Outlook.Application')
#         mail = outlook.CreateItem(0)
#         mail.To = destinataire
#         mail.CC = cc
#         mail.Subject = sujet
#         mail.HTMLBody = corps_html
#         if piece_jointe:
#             mail.Attachments.Add(piece_jointe)
#         mail.Send()
#         print("Email envoyé avec succès.")
#         return True
#     except Exception as e:
#         print(f"Erreur envoi email: {e}")
#         return False

# def envoyer_email_alerte(erreur_msg):
#     try:
#         outlook = win32.DispatchEx('Outlook.Application')
#         mail = outlook.CreateItem(0)
#         mail.Subject = "⚠️ ALERTE - BOT KYC Brutes"
#         mail.To = "ix41p@ningen-group.com "
#         #mail.CC = "fe48v@ningen-group.com;Ningen-pperformance@ningen-group.com;comd@ningen-group.com"
#         mail.Importance = 2
#         html_body = f"""
#         <html>
#         <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 0;">
#             <div style="max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e74c3c; border-radius: 8px; background-color: #fdf2f2; box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);">
#                 <div style="background-color: #e74c3c; color: white; padding: 15px; border-radius: 6px 6px 0 0; margin: -20px -20px 20px -20px;">
#                     <h2 style="margin: 0; font-size: 20px;">⚠️ ALERTE - Dysfonctionnement BOT KYC Brutes</h2>
#                 </div>
                
#                 <div style="margin-bottom: 20px;">
#                     <p style="font-size: 16px; color: #d35400; font-weight: bold;">
#                         Une erreur critique a été détectée lors de l'exécution du script :
#                     </p>
#                     <div style="background-color: #f8d7da; border: 1px solid #f5c6cb; border-radius: 4px; padding: 15px; margin: 15px 0;">
#                         <pre style="font-family: Consolas, Monaco, 'Andale Mono', monospace; font-size: 12px; color: #721c24; margin: 0; white-space: pre-wrap; word-wrap: break-word;">{erreur_msg}</pre>
#                     </div>
#                 </div>
                
#                 <div style="background-color: #f9f9f9; padding: 15px; border-radius: 4px; border: 1px solid #ddd;">
#                     <p style="font-size: 14px; margin: 0; color: #666;">
#                         <strong>Date :</strong> {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}<br>
#                         <strong>Action requise :</strong> Intervention technique nécessaire
#                     </p>
#                 </div>
                
#                 <div style="margin-top: 20px; padding-top: 15px; border-top: 1px solid #eee;">
#                     <p style="font-size: 12px; color: #999; text-align: center;">
#                         Ceci est un email automatique, merci de ne pas y répondre.
#                     </p>
#                 </div>
#             </div>
#         </body>
#         </html>
#         """
#         mail.HTMLBody = html_body
#         mail.Send()
#         print("Email d'alerte HTML envoyé avec succès.")
#     except Exception as e:
#         print(f"Impossible d'envoyer l'email d'alerte: {e}")

# # ------------------ EXECUTION ------------------
# try:
#     # Étape 1: Charger les données existantes du fichier Excel
#     existing_data = pd.DataFrame()
#     if os.path.exists(EXCEL_FILE_PATH):
#         try:
#             # Charger le fichier Excel existant
#             existing_data = pd.read_excel(EXCEL_FILE_PATH)
#             print(f"Données existantes chargées : {len(existing_data)} lignes")
#         except Exception as e:
#             print(f"Erreur lors du chargement du fichier Excel: {e}")
#     else:
#         print("Fichier Excel non trouvé, création d'un nouveau fichier")

#     # Étape 2: Récupérer les nouvelles données depuis la base
#     with SSHTunnelForwarder(
#         (ssh_host, ssh_port),
#         ssh_username=ssh_user,
#         ssh_password=ssh_password,
#         remote_bind_address=(db_host, db_port)
#     ) as tunnel:
#         local_port = tunnel.local_bind_port
#         engine = create_engine(f'mysql+pymysql://{db_user}:{db_password}@127.0.0.1:{local_port}/{db_name}')
#         with engine.connect() as conn:
#             new_data = pd.read_sql_query(text(query), conn)

#     print(f"Nombre de nouvelles lignes récupérées depuis la DB : {len(new_data)}")

#     # Étape 3: Normaliser les nouvelles données
#     new_data = normalize_json_columns(new_data, ['data', 'response_data'])
#     new_data.columns = new_data.columns.str.replace(r'(_data|_response)$', '', regex=True)

#     # Colonnes à conserver
#     KYC_to_keep = [
#         "Date Injection","Date Traitement", "survey_data_id", "Age", "Canal", "Source", "Secteur",
#         "phone number", "phone number 2", "Nb de relances suite au 1er contact",
#         "Nom", "Prénom", "Genre", "Date de naissance", "Nationalité", "Statut marital",
#         "Nom époux (se)", "Nombre d'enfants", "Carte ID", "Niveau d'étude",
#         "Adresse Professionnelle", "Code postal", "Adresse personnelle",
#         "Code postal (adresse personnelle)", "Nature", "Téléphone bureau",
#         "Profession exercée", "Matricule fiscale", "Employeur", "Nombre d'emplois à créer",
#         "Activité de l'entreprise", "Sous-secteurs", "En affaire depuis", "Nom bénéficiaire effectif",
#         "Prénom bénéficiaire effectif", "Care ID bénéficiaire effectif",
#         "Où avez-vous entendu parler de DAAM (source de sollicitation)?",
#         "Exercez-vous ou avez-vous exercé une fonction politique, juridictionnelle ou administrative importante?",
#         "Fonction exercée", "Dans quel pays", "Date fin de mission",
#         "Avez-vous une PPE qui est membre de votre famille ou une personne avec laquelle vous êtes en étroite relation d'affaire?",
#         "Nom et prénom de la personne", "Lien de parenté", "Fonction exercée du membre",
#         "Membre dans quel pays", "Date de la fin de mission du membre", "Objet du crédit",
#         "Type de produit DAAM", "Montant demandé", "Durée demandée",
#         "Date du versement mensuel souhaité (Le combien de chaque mois*)",
#         "Mensualité souhaitée", "Garantie(s) offerte(s)", "Valeur estimée garantie (s)",
#         "Date délivrance Carte ID", "Agency", "Gouvernaurat", "Délégation",
#         "Nombre d'employées", "Type de projet", "user_id"
#     ]
    
#     # Filtrer les colonnes existantes
#     KYC_existing = [col for col in KYC_to_keep if col in new_data.columns]
#     new_data = new_data.loc[:, ~new_data.columns.duplicated()].loc[:, KYC_existing]
    
#     # Étape 4: Éviter les doublons basés sur survey_data_id ET Date Traitement
#     if not existing_data.empty and not new_data.empty:
#         # S'assurer que les colonnes nécessaires existent dans les deux DataFrames
#         if 'survey_data_id' in existing_data.columns and 'survey_data_id' in new_data.columns and \
#            'Date Traitement' in existing_data.columns and 'Date Traitement' in new_data.columns:
            
#             # Convertir Date Traitement en format datetime pour comparaison
#             existing_data['Date Traitement_temp'] = pd.to_datetime(existing_data['Date Traitement'], errors='coerce')
#             new_data['Date Traitement_temp'] = pd.to_datetime(new_data['Date Traitement'], errors='coerce')
            
#             # Convertir survey_data_id en string pour comparaison
#             existing_data['survey_data_id_temp'] = existing_data['survey_data_id'].astype(str)
#             new_data['survey_data_id_temp'] = new_data['survey_data_id'].astype(str)
            
#             # Créer une clé composite pour les deux DataFrames
#             existing_data['composite_key'] = existing_data['survey_data_id_temp'] + '_' + \
#                                             existing_data['Date Traitement_temp'].dt.strftime('%Y-%m-%d %H:%M:%S')
#             new_data['composite_key'] = new_data['survey_data_id_temp'] + '_' + \
#                                        new_data['Date Traitement_temp'].dt.strftime('%Y-%m-%d %H:%M:%S')
            
#             # Récupérer les clés existantes
#             existing_keys = set(existing_data['composite_key'].dropna())
            
#             # Filtrer pour garder seulement les nouvelles entrées
#             new_data = new_data[~new_data['composite_key'].isin(existing_keys)]
            
#             # Nettoyer les colonnes temporaires
#             new_data = new_data.drop(['Date Traitement_temp', 'survey_data_id_temp', 'composite_key'], axis=1)
#             existing_data = existing_data.drop(['Date Traitement_temp', 'survey_data_id_temp', 'composite_key'], axis=1)
            
#             print(f"Nouvelles entrées après suppression des doublons : {len(new_data)}")
        
#         else:
#             print("Colonnes 'survey_data_id' ou 'Date Traitement' manquantes pour la vérification des doublons")
    
#     # Étape 5: Ajouter les colonnes de traitement
#     if not new_data.empty:
#         new_data["Etat KYC"] = ""
#         new_data["Raison Rejet"] = ""
#         new_data["Infos à compléter"] = ""
        
#         # Nettoyage des montants
#         new_data['Montant demandé'] = new_data['Montant demandé'].apply(lambda val: ''.join(filter(str.isdigit, str(val))) if pd.notnull(val) else None)
    
#     fields = [
#         "survey_data_id",
#         "Date Traitement",
#         "Date Injection",
#         "Nom",
#         "Prénom",
#         "Genre",
#         "Date de naissance",
#         "Nationalité",
#         "Statut marital",
#         "Nom époux (se)",
#         "Nombre d'enfants",
#         "Carte ID",
#         "Date délivrance Carte ID",
#         "phone number",
#         "phone number 2",
#         "Gouvernaurat",
#         "Délégation",
#         "Agency",
#         "Adresse personnelle",
#         "Code postal (adresse personnelle)",
#         "Adresse Professionnelle",
#         "Code postal",
#         "Téléphone bureau",
#         "Profession exercée",
#         "Employeur",
#         "Activité de l'entreprise",
#         "Matricule fiscale",
#         "Nombre d'employées",
#         "Nombre d'emplois à créer",
#         "Sous-secteurs",
#         "En affaire depuis",
#         "Age",
#         "Secteur",
#         "Niveau d'étude",
#         "Nature",
#         "Canal",
#         "Source",
#         "Nb de relances suite au 1er contact",
#         "Date fin de mission",
#         "Care ID bénéficiaire effectif",
#         "Nom bénéficiaire effectif",
#         "Prénom bénéficiaire effectif",
#         "Exercez-vous ou avez-vous exercé une fonction politique, juridictionnelle ou administrative importante?",
#         "Fonction exercée",
#         "Dans quel pays",
#         "Avez-vous une PPE qui est membre de votre famille ou une personne avec laquelle vous êtes en étroite relation d'affaire?",
#         "Nom et prénom de la personne",
#         "Lien de parenté",
#         "Fonction exercée du membre",
#         "Membre dans quel pays",
#         "Date de la fin de mission du membre",
#         "Type de produit DAAM",
#         "Type de projet",
#         "Objet du crédit",
#         "Montant demandé",
#         "Durée demandée",
#         "Date du versement mensuel souhaité (Le combien de chaque mois*)",
#         "Mensualité souhaitée",
#         "Garantie(s) offerte(s)",
#         "Valeur estimée garantie (s)",
#         "Où avez-vous entendu parler de DAAM (source de sollicitation)?",
#         "Etat KYC",
#         "Raison Rejet",
#         "Infos à compléter",
#         "user_id"
#     ]

#     # Étape 6: Combiner les données existantes avec les nouvelles
#     if not new_data.empty:
#         # S'assurer que toutes les colonnes existent dans new_data
#         existing_fields = [col for col in fields if col in new_data.columns]
#         new_data = new_data[existing_fields]
#         new_data.rename(columns={'Gouvernaurat': 'Gouvernorat', 'Agency': 'Agence'}, inplace=True)
        
#         # Harmoniser les colonnes entre existing_data et new_data
#         if not existing_data.empty:
#             for col in existing_data.columns:
#                 if col not in new_data.columns and col not in ['Etat KYC', 'Raison Rejet', 'Infos à compléter']:
#                     new_data[col] = None
            
#             for col in new_data.columns:
#                 if col not in existing_data.columns:
#                     existing_data[col] = None
        
#         # Combiner les données
#         final_data = pd.concat([existing_data, new_data], ignore_index=True, sort=False)
#         print(f"Total après fusion : {len(final_data)} lignes ({len(new_data)} nouvelles)")
#     else:
#         final_data = existing_data
#         print("Aucune nouvelle donnée à ajouter")
    
#     # Réorganiser les colonnes selon l'ordre spécifié
#     final_fields = [col for col in fields if col in final_data.columns]
#     # Ajouter les colonnes de traitement si elles existent
#     for col in ["Etat KYC", "Raison Rejet", "Infos à compléter"]:
#         if col in final_data.columns and col not in final_fields:
#             final_fields.append(col)
    
#     final_data = final_data[final_fields]

#     # Étape 7: Sauvegarder dans le fichier Excel
#     if not final_data.empty:
#         try:
#             # Sauvegarder avec pandas (sans protection par mot de passe pour simplifier)
#             final_data.to_excel(EXCEL_FILE_PATH, index=False)
#             print(f"Fichier Excel sauvegardé avec succès : {EXCEL_FILE_PATH}")
            
#             # Optionnel: Ajouter la protection par mot de passe avec win32com
#             try:
#                 excel = win32.Dispatch('Excel.Application')
#                 excel.Visible = False
#                 excel.DisplayAlerts = False
                
#                 wb = excel.Workbooks.Open(EXCEL_FILE_PATH)
#                 wb.SaveAs(EXCEL_FILE_PATH, Password='DAAM@2025')
#                 wb.Close()
#                 excel.Quit()
#                 print("Protection par mot de passe appliquée")
#             except Exception as e:
#                 print(f"Erreur lors de l'application de la protection: {e}")
            
#         except Exception as e:
#             print(f"Erreur lors de la sauvegarde Excel: {e}")
#             raise
    
#     # Étape 8: ZIP et envoi
#     if len(new_data) > 0 and os.path.exists(EXCEL_FILE_PATH):  # Envoyer seulement s'il y a de nouvelles données
#         zip_path = creer_zip(EXCEL_FILE_PATH)
#         if zip_path:
#             formatted_start_date = (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y")
#             formatted_end_date = datetime.now().strftime("%d/%m/%Y")
#             logo_path = r"C:\Users\Administrateur\Desktop\WIMOVA\Post Prod_Wimova\logo-complet-color.png"
            
#             encoded_image = ""
#             if os.path.exists(logo_path):
#                 with open(logo_path, "rb") as img_file:
#                     encoded_image = base64.b64encode(img_file.read()).decode("utf-8")
            
#             corps_email = f'''
#             <html>
#             <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
#                 <div style="max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #ddd; border-radius: 8px; background-color: #f9f9f9;">
#                     <div style="position: relative; margin-bottom: 40px;">
#                         <img src="data:image/png;base64,{encoded_image}" alt="Logo" style="height:50px; width:100px; object-fit:contain; position:absolute; right:0; top:0;">
#                     </div>
#                     <p>Bonjour,</p>
#                     <p>Veuillez trouver en pièce jointe les fiches KYC Brutes mises à jour (schéma 25).</p>
#                     <ul style="list-style-type:none; padding:0;">
#                         <li style="margin-bottom:10px; padding:10px; background:#fff; border-radius:4px; border:1px solid #ddd;">
#                             <span style="font-weight:bold; color:#004986;">Nouvelles fiches ajoutées :</span>
#                             <span>{len(new_data)} fiches</span>
#                         </li>
#                         <li style="margin-bottom:10px; padding:10px; background:#fff; border-radius:4px; border:1px solid #ddd;">
#                             <span style="font-weight:bold; color:#004986;">Total fiches dans le fichier :</span>
#                             <span>{len(final_data)} fiches</span>
#                         </li>
#                     </ul>
#                     <p><strong>NINGEN Data Analytics</strong><br></p>
                        
#                         <div style="margin-top: 20px; font-size: 12px; color: #666;">
#                             <p>
#                                 Ceci est un message généré automatiquement. Merci de ne pas y répondre.<br> 
#                                 <strong>Besoin d'assistance ?</strong><br>
#                                 Veuillez contacter :
#                                 <a href="mailto:Ningen-Data-Management@ningen-group.com">
#                                     Ningen-Data-Management@ningen-group.com
#                                 </a>
#                             </p>
#                         </div>

#                         <div style="text-align: center; margin-top: 20px; font-size: 10px; color: #666;">
#                             <p>
#                                 Ce message et les éventuelles pièces jointes sont strictement confidentiels et destinés exclusivement au(x) destinataire(s) indiqué(s). Toute utilisation, diffusion ou reproduction non autorisée est interdite. Si vous avez reçu ce message par erreur, merci d'en avertir immédiatement l'expéditeur et de supprimer le courriel.
#                             </p>
#                         </div>
#                 </div>
#             </body>
#             </html>'''

#             envoyer_email(
#                 destinataire="ix41p@ningen-group.com",
#                 #destinataire="amine.cherni@daam.tn;Ameni.Drissi@daam.tn",
#                 #cc="sl06h@ningen-group.com ;td75s@ningen-group.com;aymen.ouertani@daam.tn;Ningen-Data-Management@ningen-group.com;Ningen-pperformance@ningen-group.com",
#                 sujet=f"[Confidentiel] DAAM - Mise à jour Fiches KYC Brutes ({len(new_data)} nouvelles, schéma 25)",
#                 corps_html=corps_email,
#                 piece_jointe=zip_path
#             )
            
#             # Nettoyer le fichier ZIP temporaire
#             if os.path.exists(zip_path):
#                 try:
#                     os.remove(zip_path)
#                     print(f"Archive ZIP temporaire supprimée: {zip_path}")
#                 except Exception:
#                     pass
#         else:
#             envoyer_email_alerte("Échec de la création de l'archive ZIP.")
#     elif len(new_data) == 0:
#         print("Aucune nouvelle donnée à ajouter. Email non envoyé.")

# except Exception:
#     erreur_trace = traceback.format_exc()
#     print(f"Erreur lors du traitement des données : {erreur_trace}")
#     envoyer_email_alerte(erreur_trace)


# #==================================================================================================================================================

# #=====================================partie3================================================================================================================

# #==========================================================================================================================================

import pandas as pd
from datetime import datetime, timedelta
from sqlalchemy import create_engine, text
from sshtunnel import SSHTunnelForwarder
import json
import os
import traceback
import base64
import win32com.client as win32
import zipfile
import requests
from io import BytesIO
import msal
from dotenv import load_dotenv

env_path = r"C:\Users\Administrateur\Desktop\Daam\DAAM_KYC\.env"
load_dotenv(dotenv_path=env_path)

# ------------------ CONFIG SSH ------------------
ssh_host = os.getenv("SSH_HOST")
ssh_port = int(os.getenv("SSH_PORT", "22"))
ssh_user = os.getenv("SSH_USER")
ssh_password = os.getenv("SSH_PASSWORD")

# ------------------ CONFIG DB ------------------
db_host = os.getenv("DB_HOST")
db_port = int(os.getenv("DB_PORT"))
db_user = os.getenv("DB_USER")
db_password = os.getenv("DB_PASSWORD")
db_name = os.getenv("DB_NAME")

# ------------------ CONFIG SHAREPOINT ------------------



TENANT_ID     = os.getenv("TENANT_ID")
CLIENT_ID     = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
SITE_DOMAIN = os.getenv("SITE_DOMAIN")
SITE_NAME     = os.getenv("SITE_NAME")



SHAREPOINT_FILE_LOG_PATH ="General/Autoreports Status"
# Log file name
log_filename = f"KYC_Brutes_{datetime.now().strftime('%Y%m%d_%H.%M')}.txt"




SHAREPOINT_FILE_PATH = "General/DAAM/C-KYC/KYC_Brutes.xlsx"

# ------------------ SQL ------------------
query = """
SELECT inj.id,
       inj.data,
       inj.created_at AS `Date Injection`,
       rslt.response_data,
       rslt.created_at AS `Date Traitement`,
       rslt.form_id,
       rslt.user_id, 
       rslt.survey_data_id,
       rslt.survey_schema_id
FROM myapp_formresponse AS rslt
INNER JOIN myapp_surveydata AS inj
        ON rslt.survey_data_id = inj.id
WHERE (rslt.survey_schema_id = 28 or rslt.survey_schema_id = 36) 
  AND DATE(rslt.created_at) >= '2025-08-01'
  AND JSON_UNQUOTE(JSON_EXTRACT(rslt.response_data, '$.Qualification')) = 'Formulaire complété';
"""
def write_log(message):
    global sharepoint_headers, drive_id
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    full_message = f"[{timestamp}] {message}\n"
    print(full_message)

    if sharepoint_headers is None or drive_id is None:
        return

    try:
        # Vérifier si fichier log existe
        response = requests.get(
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{SHAREPOINT_FILE_LOG_PATH}/{log_filename}:/content",
            headers=sharepoint_headers
        )
        log_stream = BytesIO(full_message.encode("utf-8"))
        if response.status_code == 200:
            existing_content = BytesIO(response.content)
            combined = existing_content.read() + log_stream.read()
            log_stream = BytesIO(combined)

        # Upload (create ou update)
        upload_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{SHAREPOINT_FILE_LOG_PATH}/{log_filename}:/content"
        upload_response = requests.put(upload_url, headers=sharepoint_headers, data=log_stream)
        upload_response.raise_for_status()
    except Exception as e:
        print(f"⚠️ Impossible de logger sur SharePoint : {e}")

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
    # D'abord, vérifier si le fichier existe et obtenir son ID
    try:
        response = requests.get(
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{file_path}",
            headers=headers
        )
        if response.status_code == 200:
            file_id = response.json()["id"]
            upload_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}/content"
        else:
            # Créer un nouveau fichier
            parent_path = "/".join(file_path.split("/")[:-1])
            file_name = file_path.split("/")[-1]
            
            # Obtenir l'ID du dossier parent
            folder_response = requests.get(
                f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{parent_path}",
                headers=headers
            )
            folder_id = folder_response.json()["id"]
            
            upload_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{folder_id}:/{file_name}:/content"
    except requests.exceptions.HTTPError:
        # Si le fichier n'existe pas, créer un nouveau
        parent_path = "/".join(file_path.split("/")[:-1])
        file_name = file_path.split("/")[-1]
        
        # Obtenir l'ID du dossier parent
        folder_response = requests.get(
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{parent_path}",
            headers=headers
        )
        folder_id = folder_response.json()["id"]
        
        upload_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{folder_id}:/{file_name}:/content"

    # Uploader le fichier
    with open(local_file_path, 'rb') as file:
        response = requests.put(upload_url, headers=headers, data=file)
    
    response.raise_for_status()
    return response.json()

# ------------------ FONCTION DE NORMALISATION ------------------
def normalize_json_columns(df, columns_to_normalize):
    for column_name in columns_to_normalize:
        if column_name in df.columns and len(df) > 0:
            if isinstance(df[column_name].iloc[0], str):
                try:
                    df[column_name] = df[column_name].apply(json.loads)
                    normalized_df = pd.json_normalize(df[column_name])
                    suffix = "_response" if column_name == "response_data" else "_data"
                    normalized_df = normalized_df.add_suffix(suffix)
                    df = pd.concat([df.drop(column_name, axis=1), normalized_df], axis=1)
                    print(f"Colonne {column_name} normalisée avec suffixe '{suffix}'")
                except Exception as e:
                    print(f"Erreur normalisation {column_name}: {e}")
            else:
                print(f"La colonne {column_name} ne contient pas de données JSON string")
        else:
            print(f"Colonne {column_name} non trouvée ou DataFrame vide")
    return df

# ------------------ CREATION ZIP ------------------
def creer_zip(file_path):
    try:
        zip_path = file_path.replace(".xlsx", ".zip")
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            zf.write(file_path, arcname=os.path.basename(file_path))
        return zip_path
    except Exception as e:
        print(f"Erreur création ZIP: {e}")
        return None

# ------------------ ENVOI EMAIL ------------------
def envoyer_email(destinataire, cc="", sujet="", corps_html="", piece_jointe=None):
    try:
        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.To = destinataire
        mail.CC = cc
        mail.Subject = sujet
        mail.HTMLBody = corps_html
        if piece_jointe:
            mail.Attachments.Add(piece_jointe)
        mail.Send()
        print("Email envoyé avec succès.")
        return True
    except Exception as e:
        print(f"Erreur envoi email: {e}")
        return False

def envoyer_email_alerte(erreur_msg):
    try:
        outlook = win32.DispatchEx('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.Subject = "⚠️ ALERTE - BOT KYC Brutes"
        mail.To = "ix41p@ningen-group.com "
        mail.Importance = 2
        html_body = f"""
        <html>
        <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 0;">
            <div style="max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e74c3c; border-radius: 8px; background-color: #fdf2f2; box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);">
                <div style="background-color: #e74c3c; color: white; padding: 15px; border-radius: 6px 6px 0 0; margin: -20px -20px 20px -20px;">
                    <h2 style="margin: 0; font-size: 20px;">⚠️ ALERTE - Dysfonctionnement BOT KYC Brutes</h2>
                </div>
                
                <div style="margin-bottom: 20px;">
                    <p style="font-size: 16px; color: #d35400; font-weight: bold;">
                        Une erreur critique a été détectée lors de l'exécution du script :
                    </p>
                    <div style="background-color: #f8d7da; border: 1px solid #f5c6cb; border-radius: 4px; padding: 15px; margin: 15px 0;">
                        <pre style="font-family: Consolas, Monaco, 'Andale Mono', monospace; font-size: 12px; color: #721c24; margin: 0; white-space: pre-wrap; word-wrap: break-word;">{erreur_msg}</pre>
                    </div>
                </div>
                
                <div style="background-color: #f9f9f9; padding: 15px; border-radius: 4px; border: 1px solid #ddd;">
                    <p style="font-size: 14px; margin: 0; color: #666;">
                        <strong>Date :</strong> {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}<br>
                        <strong>Action requise :</strong> Intervention technique nécessaire
                    </p>
                </div>
                
                <div style="margin-top: 20px; padding-top: 15px; border-top: 1px solid #eee;">
                    <p style="font-size: 12px; color: #999; text-align: center;">
                        Ceci est un email automatique, merci de ne pas y répondre.
                    </p>
                </div>
            </div>
        </body>
        </html>
        """
        mail.HTMLBody = html_body
        mail.Send()
        print("Email d'alerte HTML envoyé avec succès.")
    except Exception as e:
        print(f"Impossible d'envoyer l'email d'alerte: {e}")





#  Ajouter le préfixe 00216 aux numéros de téléphone
def add_prefix_to_phone_numbers(df):
    """Ajoute le préfixe 00216 aux numéros de téléphone si nécessaire"""
    phone_columns = ['phone number', 'phone number 2']
    
    for col in phone_columns:
        if col in df.columns:
            # Fonction pour ajouter le préfixe si le numéro n'a pas déjà le préfixe
            def format_phone_number(phone):
                if pd.isna(phone) or phone == '' or phone is None:
                    return phone
                
                # Convertir en string et nettoyer
                phone_str = str(phone).strip()
                
                # Supprimer les espaces, tirets, etc. pour normaliser
                phone_str = ''.join(filter(str.isdigit, phone_str))
                
                # Si le numéro n'est pas vide et ne commence pas par 00216
                if phone_str and not phone_str.startswith('00216'):
                    # Si le numéro commence par 0, enlever le 0 avant d'ajouter le préfixe
                    if phone_str.startswith('0'):
                        phone_str = phone_str[1:]
                    # Ajouter le préfixe
                    phone_str = '00216' + phone_str
                
                return phone_str
            
            df[col] = df[col].apply(format_phone_number)
            print(f"Préfixe 00216 ajouté à la colonne {col}")
    
    return df

# ------------------ EXECUTION ------------------
try:
    # Étape 1: Authentification SharePoint
    print("Authentification SharePoint...")
    sharepoint_headers = authenticate_sharepoint()
    drive_id = get_drive_id(sharepoint_headers)
    
    
    # Étape 2: Charger les données existantes depuis SharePoint
    write_log(f"Chargement du fichier depuis SharePoint: {SHAREPOINT_FILE_PATH}")
    existing_data = read_excel_from_sharepoint(sharepoint_headers, drive_id, SHAREPOINT_FILE_PATH)
    
    if existing_data.empty:
        write_log("Fichier SharePoint vide ou non trouvé, création de nouvelles données")
        existing_rows = 0
    else:
        existing_rows = len(existing_data)
        write_log(f"Données existantes chargées depuis SharePoint : {existing_rows} lignes")

    # Étape 3: Récupérer les nouvelles données depuis la base de données
    print("Connexion à la base de données via SSH...")
    with SSHTunnelForwarder(
        (ssh_host, ssh_port),
        ssh_username=ssh_user,
        ssh_password=ssh_password,
        remote_bind_address=(db_host, db_port)
    ) as tunnel:
        local_port = tunnel.local_bind_port
        engine = create_engine(f'mysql+pymysql://{db_user}:{db_password}@127.0.0.1:{local_port}/{db_name}')
        with engine.connect() as conn:
            new_data = pd.read_sql_query(text(query), conn)

    write_log(f"Nombre de nouvelles lignes récupérées depuis la DB : {len(new_data)}")

    # Étape 4: Normaliser les nouvelles données
    new_data = normalize_json_columns(new_data, ['data', 'response_data'])
    new_data.columns = new_data.columns.str.replace(r'(_data|_response)$', '', regex=True)
    # Étape 4.5: Filtrer les données déjà existantes
    if not existing_data.empty and 'survey_data_id' in existing_data.columns:
        # Obtenir les IDs déjà présents dans le fichier SharePoint
        existing_ids = set(existing_data['survey_data_id'].dropna().astype(str))
        
        # Filtrer les nouvelles données pour ne garder que celles qui n'existent pas déjà
        new_data['survey_data_id'] = new_data['survey_data_id'].astype(str)
        mask = ~new_data['survey_data_id'].isin(existing_ids)
        new_data = new_data[mask].copy()
        
        
    else:
        write_log("Pas de données existantes ou colonne survey_data_id manquante, toutes les données seront ajoutées")
    # Colonnes à conserver
    KYC_to_keep = [
        "Date Injection","Date Traitement", "survey_data_id", "Age", "Canal", "Source", "Secteur",
        "phone number", "phone number 2", "Nb de relances suite au 1er contact",
        "Nom", "Prénom", "Genre", "Date de naissance", "Nationalité", "Statut marital",
        "Nom époux (se)", "Nombre d'enfants", "Carte ID", "Niveau d'étude",
        "Adresse Professionnelle", "Code postal", "Adresse personnelle",
        "Code postal (adresse personnelle)", "Nature", "Téléphone bureau",
        "Profession exercée", "Matricule fiscale", "Employeur", "Nombre d'emplois à créer",
        "Activité de l'entreprise", "Sous-secteurs", "En affaire depuis", "Nom bénéficiaire effectif",
        "Prénom bénéficiaire effectif", "Care ID bénéficiaire effectif",
        "Où avez-vous entendu parler de DAAM (source de sollicitation)?",
        "Exercez-vous ou avez-vous exercé une fonction politique, juridictionnelle ou administrative importante?",
        "Fonction exercée", "Dans quel pays", "Date fin de mission",
        "Avez-vous une PPE qui est membre de votre famille ou une personne avec laquelle vous êtes en étroite relation d'affaire?",
        "Nom et prénom de la personne", "Lien de parenté", "Fonction exercée du membre",
        "Membre dans quel pays", "Date de la fin de mission du membre", "Objet du crédit",
        "Type de produit DAAM", "Montant demandé", "Durée demandée",
        "Date du versement mensuel souhaité (Le combien de chaque mois*)",
        "Mensualité souhaitée", "Garantie(s) offerte(s)", "Valeur estimée garantie (s)",
        "Date délivrance Carte ID", "Agency", "Gouvernaurat", "Délégation",
        "Nombre d'employées", "Type de projet", "user_id","IDu"
    ]
    
    # Filtrer les colonnes existantes
    KYC_existing = [col for col in KYC_to_keep if col in new_data.columns]
    new_data = new_data.loc[:, ~new_data.columns.duplicated()].loc[:, KYC_existing]
    
    # Étape 5: Préparer les nouvelles données
    if not new_data.empty:
        # Ajouter les colonnes de traitement vides pour les nouvelles données
        new_data["Etat KYC"] = ""
        new_data["Raison Rejet"] = ""
        new_data["Infos à compléter"] = ""
        
        # Nettoyage des montants
        new_data['Montant demandé'] = new_data['Montant demandé'].apply(lambda val: ''.join(filter(str.isdigit, str(val))) if pd.notnull(val) else None)
    
    # Liste des champs pour l'ordre final
    fields = [
        "survey_data_id",
        "Date Traitement",
        "Date Injection",
        "Nom",
        "Prénom",
        "Genre",
        "Date de naissance",
        "Nationalité",
        "Statut marital",
        "Nom époux (se)",
        "Nombre d'enfants",
        "Carte ID",
        "Date délivrance Carte ID",
        "phone number",
        "phone number 2",
        "Gouvernaurat",
        "Délégation",
        "Agency",
        "Adresse personnelle",
        "Code postal (adresse personnelle)",
        "Adresse Professionnelle",
        "Code postal",
        "Téléphone bureau",
        "Profession exercée",
        "Employeur",
        "Activité de l'entreprise",
        "Matricule fiscale",
        "Nombre d'employées",
        "Nombre d'emplois à créer",
        "Sous-secteurs",
        "En affaire depuis",
        "Age",
        "Secteur",
        "Niveau d'étude",
        "Nature",
        "Canal",
        "Source",
        "Nb de relances suite au 1er contact",
        "Date fin de mission",
        "Care ID bénéficiaire effectif",
        "Nom bénéficiaire effectif",
        "Prénom bénéficiaire effectif",
        "Exercez-vous ou avez-vous exercé une fonction politique, juridictionnelle ou administrative importante?",
        "Fonction exercée",
        "Dans quel pays",
        "Avez-vous une PPE qui est membre de votre famille ou une personne avec laquelle vous êtes en étroite relation d'affaire?",
        "Nom et prénom de la personne",
        "Lien de parenté",
        "Fonction exercée du membre",
        "Membre dans quel pays",
        "Date de la fin de mission du membre",
        "Type de produit DAAM",
        "Type de projet",
        "Objet du crédit",
        "Montant demandé",
        "Durée demandée",
        "Date du versement mensuel souhaité (Le combien de chaque mois*)",
        "Mensualité souhaitée",
        "Garantie(s) offerte(s)",
        "Valeur estimée garantie (s)",
        "Où avez-vous entendu parler de DAAM (source de sollicitation)?",
        "Etat KYC",
        "Raison Rejet",
        "Infos à compléter",
        "user_id","IDu"
    ]

    # Étape 6: Ajouter les nouvelles données aux données existantes
    if not new_data.empty:
        # Réorganiser les colonnes des nouvelles données
        existing_fields = [col for col in fields if col in new_data.columns]
        new_data = new_data[existing_fields]
        
        # Renommer les colonnes
        #new_data.rename(columns={'Gouvernaurat': 'Gouvernorat', 'Agency': 'Agence'}, inplace=True)
        
        # S'assurer que toutes les colonnes nécessaires existent
        for col in fields:
            if col not in new_data.columns:
                new_data[col] = None
        
        # Réorganiser dans l'ordre correct
        new_data = new_data[fields]
        
        write_log(f"Nouvelles données à ajouter : {len(new_data)} lignes")
        
        # Ajouter les nouvelles données aux données existantes
        if not existing_data.empty:
            # S'assurer que les colonnes sont dans le même ordre
            for col in fields:
                if col not in existing_data.columns:
                    existing_data[col] = None
            
            # Réorganiser les données existantes
            existing_data = existing_data[fields]
            
            # Concaténer les nouvelles données
            final_data = pd.concat([existing_data, new_data], ignore_index=True)
        else:
            final_data = new_data
        
        print(f"Total après ajout : {len(final_data)} lignes ({existing_rows} existantes + {len(new_data)} nouvelles)")
    else:
        final_data = existing_data
        write_log("Aucune nouvelle donnée à ajouter, fichier inchangé")

    # # Appliquer la fonction d'ajout de préfixe
    # if not final_data.empty:
    #     final_data = add_prefix_to_phone_numbers(final_data)
        

    # Étape 7: Sauvegarder temporairement le fichier localement
    if not final_data.empty:
        try:
            # Créer un dossier temporaire
            temp_folder = os.path.join(os.getenv('TEMP'), 'kyc_temp')
            os.makedirs(temp_folder, exist_ok=True)
            temp_file = os.path.join(temp_folder, 'KYC_Brutes_temp.xlsx')
            
            # Sauvegarder dans un fichier temporaire
            final_data.to_excel(temp_file, index=False)
            
            
            # Étape 8: Uploader vers SharePoint
            print(f"Upload du fichier vers SharePoint: {SHAREPOINT_FILE_PATH}")
            upload_result = upload_excel_to_sharepoint(
                sharepoint_headers, 
                drive_id, 
                SHAREPOINT_FILE_PATH, 
                temp_file
            )
            print(f"Fichier uploadé avec succès vers SharePoint: {upload_result.get('webUrl', 'URL non disponible')}")
            
            # Nettoyer le fichier temporaire
            if os.path.exists(temp_file):
                os.remove(temp_file)
                write_log("Fichier temporaire nettoyé")
                
        except Exception as e:
            write_log(f"Erreur lors de la sauvegarde/upload: {str(e)}")
            raise
    
    # Étape 9: ZIP et envoi (uniquement si nouvelles données)
    if len(new_data) > 0:
        # Créer un fichier local pour le ZIP
        local_zip_file = os.path.join(temp_folder, 'KYC_Brutes.xlsx')
        final_data.to_excel(local_zip_file, index=False)
        
        zip_path = creer_zip(local_zip_file)
        if zip_path:
            formatted_start_date = (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y")
            formatted_end_date = datetime.now().strftime("%d/%m/%Y")
            logo_path = r"C:\Users\Administrateur\Desktop\WIMOVA\Post Prod_Wimova\logo-complet-color.png"
            
            encoded_image = ""
            if os.path.exists(logo_path):
                with open(logo_path, "rb") as img_file:
                    encoded_image = base64.b64encode(img_file.read()).decode("utf-8")
            
            corps_email = f'''
            <html>
            <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
                <div style="max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #ddd; border-radius: 8px; background-color: #f9f9f9;">
                    <div style="position: relative; margin-bottom: 40px;">
                        <img src="data:image/png;base64,{encoded_image}" alt="Logo" style="height:50px; width:100px; object-fit:contain; position:absolute; right:0; top:0;">
                    </div>
                    <p>Bonjour,</p>
                    Veuillez trouver en pièce jointe les fiches KYC Brutes complétées pour la période du
                    <strong>{formatted_start_date}</strong> au <strong>{formatted_end_date}</strong>.</p>
                    <ul style="list-style-type:none; padding:0;">
                        
                         <li style="margin-bottom:10px; padding:10px; background:#fff; border-radius:4px; border:1px solid #ddd;">
                        <span style="font-weight:bold; color:#004986;">KYC Brutes :</span>
                        <span>{len(final_data)} fiches</span>
                        
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
            </body>
            </html>'''

            envoyer_email(
                #destinataire="ix41p@ningen-group.com",
                destinataire="amine.cherni@daam.tn;Ameni.Drissi@daam.tn",
                cc="sl06h@ningen-group.com ;td75s@ningen-group.com;aymen.ouertani@daam.tn;Ningen-Data-Management@ningen-group.com;Ningen-pperformance@ningen-group.com;pw39f@ningen-group.com",
                sujet=f"[Confidentiel] DAAM - Fiches KYC Brutes du {formatted_start_date} au {formatted_end_date}",
                corps_html=corps_email,
                piece_jointe=zip_path
            )
            
            # Nettoyer les fichiers temporaires
            for file_path in [local_zip_file, zip_path]:
                if os.path.exists(file_path):
                    try:
                        os.remove(file_path)
                        print(f"Fichier temporaire supprimé: {file_path}")
                    except Exception:
                        pass
        else:
            envoyer_email_alerte("Échec de la création de l'archive ZIP.")
    elif len(new_data) == 0:
        
        # Créer un fichier local pour le ZIP
        local_zip_file = os.path.join(temp_folder, 'KYC_Brutes.xlsx')
        final_data.to_excel(local_zip_file, index=False)
        
        zip_path = creer_zip(local_zip_file)
        if zip_path:
            formatted_start_date = (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y")
            formatted_end_date = datetime.now().strftime("%d/%m/%Y")
            logo_path = r"C:\Users\Administrateur\Desktop\WIMOVA\Post Prod_Wimova\logo-complet-color.png"
            
            encoded_image = ""
            if os.path.exists(logo_path):
                with open(logo_path, "rb") as img_file:
                    encoded_image = base64.b64encode(img_file.read()).decode("utf-8")
            
            corps_email = f'''
            <html>
            <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
                <div style="max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #ddd; border-radius: 8px; background-color: #f9f9f9;">
                    <div style="position: relative; margin-bottom: 40px;">
                        <img src="data:image/png;base64,{encoded_image}" alt="Logo" style="height:50px; width:100px; object-fit:contain; position:absolute; right:0; top:0;">
                    </div>
                    <p>Bonjour,</p>
                    Veuillez trouver en pièce jointe les fiches KYC Brutes complétées pour la période du
                    <strong>{formatted_start_date}</strong> au <strong>{formatted_end_date}</strong>.</p>
                    <ul style="list-style-type:none; padding:0;">
                        
                         <li style="margin-bottom:10px; padding:10px; background:#fff; border-radius:4px; border:1px solid #ddd;">
                        <span style="font-weight:bold; color:#004986;">KYC Brutes :</span>
                        <span>{len(final_data)} fiches</span>
                        
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
            </body>
            </html>'''

            envoyer_email(
                #destinataire="ix41p@ningen-group.com",
                destinataire="amine.cherni@daam.tn;Ameni.Drissi@daam.tn",
                cc="sl06h@ningen-group.com ;td75s@ningen-group.com;aymen.ouertani@daam.tn;Ningen-Data-Management@ningen-group.com;Ningen-pperformance@ningen-group.com",
                sujet=f"[Confidentiel] DAAM - Fiches KYC Brutes du {formatted_start_date} au {formatted_end_date}",
                corps_html=corps_email,
                piece_jointe=zip_path
            )
            
            # Nettoyer les fichiers temporaires
            for file_path in [local_zip_file, zip_path]:
                if os.path.exists(file_path):
                    try:
                        os.remove(file_path)
                        print(f"Fichier temporaire supprimé: {file_path}")
                    except Exception:
                        pass

except Exception :
    erreur_trace = traceback.format_exc()
    print(f"Erreur lors du traitement des données : {erreur_trace}")
    write_log(f"Erreur lors du traitement des données : {erreur_trace}")
    envoyer_email_alerte(erreur_trace)