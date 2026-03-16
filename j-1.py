# -*- coding: utf-8 -*-
import os
import tempfile
from io import BytesIO
from datetime import date, timedelta

import msal
import pandas as pd
import requests
from dotenv import load_dotenv

import win32com.client as win32
from datetime import date, datetime


# ============================================================
# LOAD ENV
# ============================================================
load_dotenv()

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

SITE_DOMAIN = os.getenv("SITE_DOMAIN")
SITE_NAME = os.getenv("SITE_NAME")
DRIVE_NAME = os.getenv("DRIVE_NAME", "Documents")

SP_J1_FOLDER = os.getenv("SP_J1_FOLDER")  # /General/Tunichèque/TEST/j-1  (optionnel, pas utilisé directement)
SP_KPI = os.getenv("SP_KPI")              # /General/Tunichèque/TEST/j-1/kpi_inbound_detail.xlsx
SP_OUTPUT_FOLDER = os.getenv("SP_OUTPUT_FOLDER")  # /General/Tunichèque/TEST/j-1
SP_INCIDENTS_FILE = os.getenv("SP_INCIDENTS_FILE")  # /General/Tunichèque/TEST/j-1/Tunichèque ...xlsx

SHAREPOINT_FILE_LOG_PATH = "General/Autoreports Status"
log_filename = f"tunicheque_intraday_client_{datetime.now().strftime('%Y%m%d_%H.%M')}.txt"

LOCAL_LOGO_PATH = os.getenv("LOCAL_LOGO_PATH", "")
MAIL_TO = os.getenv("MAIL_TO", "")
MAIL_CC = os.getenv("MAIL_CC", "")
MAIL_SUBJECT_PREFIX = os.getenv(
    "MAIL_SUBJECT_PREFIX",
    "Tunichèque - Reporting automatique {target_date}"
)

headers = None
drive_id = None




def require(v, name):
    if not v or str(v).strip() == "":
        write_log(f"Missing .env variable: {name}")
        raise ValueError(f"❌ Missing .env variable: {name}")
    return str(v).strip()


TENANT_ID = require(TENANT_ID, "TENANT_ID")
CLIENT_ID = require(CLIENT_ID, "CLIENT_ID")
CLIENT_SECRET = require(CLIENT_SECRET, "CLIENT_SECRET")
SITE_DOMAIN = require(SITE_DOMAIN, "SITE_DOMAIN")
SITE_NAME = require(SITE_NAME, "SITE_NAME")

SP_KPI = require(SP_KPI, "SP_KPI")
SP_OUTPUT_FOLDER = require(SP_OUTPUT_FOLDER, "SP_OUTPUT_FOLDER")
SP_INCIDENTS_FILE = require(SP_INCIDENTS_FILE, "SP_INCIDENTS_FILE")

MAIL_TO = require(MAIL_TO, "MAIL_TO")

LOCAL_LOGO_PATH = require(LOCAL_LOGO_PATH, "LOCAL_LOGO_PATH")
# if not os.path.isfile(LOCAL_LOGO_PATH):
#     raise FileNotFoundError(f"❌ Logo not found: {LOCAL_LOGO_PATH}")

def write_log(message: str) -> None:
    """
    Écrit le log dans la console + append dans un fichier log sur SharePoint.
    Crée un fichier log par exécution : log_filename (horodaté).
    """
    global headers, drive_id

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    full_message = f"[{timestamp}] {message}\n"
    print(full_message, end="")

    # Tant que Graph n'est pas prêt, on log seulement en console
    if headers is None or drive_id is None:
        return

    try:
        sp_log_path = f"{SHAREPOINT_FILE_LOG_PATH.strip().strip('/')}/{log_filename}"

        # 1) lire l'existant si présent
        get_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{sp_log_path}:/content"
        r_get = requests.get(get_url, headers=headers)

        combined_bytes = full_message.encode("utf-8")
        if r_get.status_code == 200:
            combined_bytes = r_get.content + combined_bytes

        # 2) upload (create/update)
        put_url = get_url  # même endpoint /content
        r_put = requests.put(put_url, headers=headers, data=combined_bytes)
        r_put.raise_for_status()

    except Exception as e:
        # Ne pas bloquer l'exécution si le log SharePoint échoue
        print(f"⚠️ Impossible de logger sur SharePoint : {e}")

# ============================================================
# GRAPH HELPERS
# ============================================================
def graph_headers() -> dict:
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET,
    )
    token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in token:
        raise RuntimeError(f"Graph token error: {token}")
    return {"Authorization": f"Bearer {token['access_token']}"}


def get_site_id(headers: dict) -> str:
    url = f"https://graph.microsoft.com/v1.0/sites/{SITE_DOMAIN}:/sites/{SITE_NAME}"
    r = requests.get(url, headers=headers)
    r.raise_for_status()
    return r.json()["id"]


def get_drive_id(headers: dict, site_id: str) -> str:
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
    r = requests.get(url, headers=headers)
    r.raise_for_status()
    for d in r.json().get("value", []):
        if (d.get("name") or "").lower() == DRIVE_NAME.lower():
            return d["id"]
    raise RuntimeError(f"Drive '{DRIVE_NAME}' not found")


def sp_download_bytes(headers: dict, drive_id: str, sp_path: str) -> bytes:
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:{sp_path}:/content"
    r = requests.get(url, headers=headers)
    r.raise_for_status()
    return r.content


def sp_read_excel(headers: dict, drive_id: str, sp_path: str) -> pd.DataFrame:
    content = sp_download_bytes(headers, drive_id, sp_path)
    df = pd.read_excel(BytesIO(content))
    df.columns = [str(c).strip() for c in df.columns]
    return df


def sp_upload_excel(headers: dict, drive_id: str, df: pd.DataFrame, sp_path: str) -> None:
    bio = BytesIO()
    df.to_excel(bio, index=False)
    bio.seek(0)
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:{sp_path}:/content"
    r = requests.put(url, headers=headers, data=bio.getvalue())
    if r.status_code not in (200, 201):
        raise RuntimeError(f"Upload failed ({r.status_code}): {r.text}")


# ============================================================
# WINDOW: FULL DAY J-1
# ============================================================
def j1_full_day_window() -> tuple[pd.Timestamp, pd.Timestamp]:
    # Yesterday 00:00:00 -> Today 00:00:00
    y = pd.Timestamp.now().normalize() - pd.Timedelta(days=1)
    start = y
    end_cap = y + pd.Timedelta(days=1)
    return start, end_cap


def format_window_full_day() -> str:
    return "00:00:00 - 23:59:59"


# ============================================================
# KPI CALC (Inbound hourly) - J-1 FULL DAY
# Output columns ONLY: Début | Fin | Reçus | Traités | QS
# ============================================================
def safe_div(n, d):
    if d in (0, None) or pd.isna(d):
        return None
    return float(n) / float(d)


def fmt_hour(h: int) -> str:
    return f"{h:02d}:00"


def normalize_call_type(x):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return None
    s = str(x).lower()
    if "inbound" in s or "entrant" in s:
        return "Inbound"
    if "outbound" in s or "sortant" in s:
        return "Outbound"
    return None


def build_hourly_inbound_for_date(df: pd.DataFrame, target_date: str, end_cap: pd.Timestamp) -> pd.DataFrame:
    required = {"datetime", "appels_recus", "appels_traites"}
    missing = required - set(df.columns)
    if missing:
        write_log(f"Colonnes manquantes dans KPI: {missing}")
        raise ValueError(f"❌ Colonnes manquantes dans KPI: {missing}")
        

    df2 = df.copy()
    df2["dt"] = pd.to_datetime(df2["datetime"], errors="coerce", dayfirst=False)
    df2 = df2.dropna(subset=["dt"])
    df2 = df2[df2["dt"].dt.strftime("%Y-%m-%d") == target_date]

    # inbound filter if exists
    if "call_type" in df2.columns:
        df2["call_type_norm"] = df2["call_type"].apply(normalize_call_type)
        df2 = df2[df2["call_type_norm"] == "Inbound"]

    df2["hour"] = df2["dt"].dt.floor("h")

    hourly = (
        df2.groupby("hour")
        .agg(
            recus=("appels_recus", "sum"),
            traites=("appels_traites", "sum"),
        )
        .reset_index()
    )

    start = pd.to_datetime(f"{target_date} 00:00:00")
    hours = pd.date_range(start=start, periods=24, freq="h")
    hourly = pd.DataFrame({"hour": hours}).merge(hourly, on="hour", how="left")

    hourly["recus"] = hourly["recus"].fillna(0).astype(int)
    hourly["traites"] = hourly["traites"].fillna(0).astype(int)

    hourly["debut"] = hourly["hour"].dt.hour.apply(fmt_hour)
    hourly["fin"] = ((hourly["hour"] + pd.Timedelta(hours=1)).dt.hour).map(fmt_hour)

    # QS only when recus > 0
    hourly["QS"] = hourly.apply(lambda r: safe_div(r["traites"], r["recus"]) if r["recus"] > 0 else None, axis=1)

    out = hourly[["hour", "debut", "fin", "recus", "traites", "QS"]].copy()

    # J-1 full day: end_cap = next day 00:00 => no future blanks, but we keep logic safe
    future_mask = out["hour"] >= end_cap
    out.loc[future_mask, ["recus", "traites", "QS"]] = None

    def qs_to_pct(x):
        if x is None or pd.isna(x):
            return ""
        return f"{float(x) * 100:.2f}%"

    out["QS"] = out["QS"].apply(qs_to_pct)
    out["recus"] = out["recus"].apply(lambda x: "" if x is None or pd.isna(x) else int(x))
    out["traites"] = out["traites"].apply(lambda x: "" if x is None or pd.isna(x) else int(x))

    reached = hourly[hourly["hour"] < end_cap].copy()
    total_recus = int(reached["recus"].sum()) if not reached.empty else 0
    total_traites = int(reached["traites"].sum()) if not reached.empty else 0
    total_qs = safe_div(total_traites, total_recus)
    total_qs_str = "" if total_qs is None else f"{total_qs * 100:.2f}%"

    # On n'a plus besoin de 'hour' dans le fichier final
    out = out.drop(columns=["hour"])

    total_row = pd.DataFrame([{
        "debut": "TOTAL",
        "fin": "",
        "recus": total_recus if total_recus != 0 else "",
        "traites": total_traites if total_traites != 0 else "",
        "QS": total_qs_str
    }])

    out = pd.concat([out, total_row], ignore_index=True)

    # ✅ Output final UNIQUEMENT 5 colonnes demandées
    out = out.rename(columns={
        "debut": "Début",
        "fin": "Fin",
        "recus": "Reçus",
        "traites": "Traités",
        "QS": "QS"
    })

    out = out[["Début", "Fin", "Reçus", "Traités", "QS"]]
    return out


# ============================================================
# INCIDENTS COUNT
# ============================================================
def count_incidents(df_inc: pd.DataFrame) -> int:
    # Simple: nombre de lignes
    return int(len(df_inc))


# ============================================================
# OUTLOOK EMAIL HTML
# ============================================================
def df_to_html_table(df: pd.DataFrame) -> str:
    cols = df.columns.tolist()

    th = "".join(
        f"<th style='padding:10px 12px;border-bottom:1px solid #e7e7ea;"
        f"text-align:center;background:#fafafa;font-weight:700;color:#1f2937;font-size:11pt;'>"
        f"{c}</th>"
        for c in cols
    )

    body_rows = []
    for _, row in df.iterrows():
        # ✅ FIX: TOTAL detection with 'Début'
        is_total = str(row.get("Début", "")).strip().upper() == "TOTAL"
        bg = "#fff7ed" if is_total else "#ffffff"
        fw = "700" if is_total else "400"
        bd = "1px solid #f0f0f2"

        tds = []
        for c in cols:
            v = row.get(c, "")
            if v is None or (isinstance(v, float) and pd.isna(v)):
                v = ""
            if v == "":
                v = "&nbsp;"
            tds.append(
                f"<td style='padding:9px 12px;border-bottom:{bd};text-align:center;"
                f"background:{bg};font-weight:{fw};color:#111827;font-size:11pt;'>{v}</td>"
            )
        body_rows.append("<tr>" + "".join(tds) + "</tr>")

    return f"""
<div style="border:1px solid #e7e7ea;border-radius:14px;overflow:hidden;background:#ffffff;">
  <table cellpadding="0" cellspacing="0" width="100%" style="border-collapse:collapse;">
    <tr>{th}</tr>
    {''.join(body_rows)}
  </table>
</div>
""".strip()


def build_html_pro(target_date: str, window_str: str, hourly_table_html: str, nb_incidents: int) -> str:
    logo_cid = "COMPANY_LOGO"
    year = date.today().year

    return f"""
<html>
  <body style="margin:0;padding:0;background:#f6f6f7;font-family:Calibri, Arial, sans-serif;font-size:11pt;color:#111827;">
    <table width="100%" cellpadding="0" cellspacing="0" style="background:#f6f6f7;padding:24px 0;">
      <tr>
        <td align="center">
          <table width="780" cellpadding="0" cellspacing="0" style="background:#ffffff;border-radius:16px;overflow:hidden;border:1px solid #e7e7ea;">
            <tr>
              <td style="background:#111827;padding:18px 22px;color:#ffffff;">
                <table width="100%" cellpadding="0" cellspacing="0">
                  <tr>
                    <td style="vertical-align:middle;">
                      <div style="font-size:16pt;font-weight:700;letter-spacing:0.2px;">Tunichèque — Reporting automatique</div>
                      <div style="margin-top:6px;font-size:10pt;opacity:0.9;">Date : <b>{target_date}</b></div>
                    </td>
                    <td align="right" style="vertical-align:middle;">
                      <img src="cid:{logo_cid}" alt="Logo" style="height:44px;display:block;border-radius:8px;background:#ffffff;padding:5px;" />
                    </td>
                  </tr>
                </table>
              </td>
            </tr>

            <tr>
              <td style="padding:18px 22px;">

                <p style="margin:0 0 10px 0;">Bonjour,</p>

                <div style="margin:0 0 10px 0;padding:12px 14px;border:1px solid #e7e7ea;border-radius:14px;background:#fbfbfc;">
                  <div style="font-weight:700;color:#111827;">Intraday {target_date}</div>
                  <div style="margin-top:4px;color:#374151;">
                    Intervalle analysé : <b>{window_str}</b>
                  </div>
                </div>

                <div style="height:10px;"></div>

                <div style="font-weight:700;font-size:12pt;margin:10px 0 10px 0;">Rapport d'Appels</div>
                {hourly_table_html}

                <div style="height:14px;"></div>

                <div style="margin:0;padding:12px 14px;border:1px solid #e7e7ea;border-radius:14px;background:#fbfbfc;">
                  <div style="font-weight:700;color:#111827;">Suivi des incidents techniques {target_date}</div>
                  <div style="margin-top:4px;color:#374151;">
                    Intervalle : <b>{window_str}</b>
                  </div>
                  <div style="margin-top:6px;color:#374151;">
                    Nombre d'incidents escaladés Support Technique Total : <b>{nb_incidents}</b>
                  </div>
                </div>

                <div style="margin:16px 0 12px 0;border-top:1px solid #e7e7ea;"></div>

                <div style="color:#6b7280;font-size:10.5pt;line-height:1.6;">
                  Ce message a été généré automatiquement par notre système. Merci de ne pas répondre à cet e-mail.<br/>
                  Pour toute question ou complément d'information, veuillez contacter l'équipe Data Management :
                  <b>Ningen-Data-Management@ningen-group.com</b>.
                </div>
              </td>
            </tr>

            <tr>
              <td style="background:#f3f4f6;padding:12px 22px;color:#6b7280;font-size:10pt;">
                © {year} Ningen Group — Data Management
              </td>
            </tr>
          </table>
          <div style="height:18px;"></div>
        </td>
      </tr>
    </table>
  </body>
</html>
""".strip()


def send_outlook_mail_with_inline_logo(to_email: str, cc: str, subject: str, html_body: str,
                                      logo_path: str, attachments: list) -> None:
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    mail.To = to_email
    mail.CC = cc or ""
    mail.Subject = subject
    mail.HTMLBody = html_body

    cid = "COMPANY_LOGO"
    att_logo = mail.Attachments.Add(logo_path)
    att_logo.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", cid)

    for p in attachments:
        mail.Attachments.Add(p)

    mail.Send()


# ============================================================
# MAIN
# ============================================================
def main():
    # ✅ target date = yesterday
    
    global headers,drive_id
    write_log("START J-1 process")
    try :
      headers = graph_headers()
      site_id = get_site_id(headers)
      drive_id = get_drive_id(headers, site_id)
      # full day window for J-1
      _, w_end_cap = j1_full_day_window()
      window_str = format_window_full_day()
      target_date = (date.today() - timedelta(days=1)).strftime("%Y-%m-%d")
      
      # KPI read
      
      df_kpi = sp_read_excel(headers, drive_id, SP_KPI)

      write_log("Building hourly INBOUND table (J-1 full day)...")
      df_hourly = build_hourly_inbound_for_date(df_kpi, target_date, end_cap=w_end_cap)

      # Upload result
      result_name = f"Hourly_Inbound_J-1_{target_date}.xlsx"
      sp_result_path = f"{SP_OUTPUT_FOLDER}/{result_name}"
      
      sp_upload_excel(headers, drive_id, df_hourly, sp_result_path)

      # incidents file (attach)
      write_log(" Downloading incidents file :")
      incidents_bytes = sp_download_bytes(headers, drive_id, SP_INCIDENTS_FILE)

      tmp_dir = tempfile.mkdtemp(prefix="tunicheque_mail_j1_")
      local_incidents = os.path.join(tmp_dir, os.path.basename(SP_INCIDENTS_FILE))
      with open(local_incidents, "wb") as f:
          f.write(incidents_bytes)

      # Count incidents rows
      try:
          df_inc = pd.read_excel(BytesIO(incidents_bytes))
          df_inc.columns = [str(c).strip() for c in df_inc.columns]
          nb_incidents = count_incidents(df_inc)
      except Exception as e :
          nb_incidents = 0
          write_log(f"WARNING incidents count failed: {type(e).__name__}: {e}")

      # Build HTML
      hourly_table_html = df_to_html_table(df_hourly)
      html = build_html_pro(
          target_date=target_date,
          window_str=window_str,
          hourly_table_html=hourly_table_html,
          nb_incidents=nb_incidents,
      )

      # Subject
      subject_prefix = MAIL_SUBJECT_PREFIX.format(target_date=target_date)
      subject = f"{subject_prefix} | {target_date}"

      write_log("Sending Outlook email to:")
      send_outlook_mail_with_inline_logo(
          to_email=MAIL_TO,
          cc=MAIL_CC,
          subject=subject,
          html_body=html,
          logo_path=LOCAL_LOGO_PATH,
          attachments=[local_incidents],
      )

      write_log(" DONE")
      

    except Exception as e:
        # best effort: log failed
        try:
            write_log(f"FAILED ❌ | {type(e).__name__}: {e}")
        except Exception:
            pass
        raise
if __name__ == "__main__":
    main()
