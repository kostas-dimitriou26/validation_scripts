import win32com.client as win32
import pandas as pd
import json
import zipfile
import os
from datetime import date, datetime

# ============================================================
# Config
# ============================================================
TARGET_SUBJECT = "Long Term Suspension Report - PRD - long-term-suspension-task"
OUTPUT_FILE = os.path.join("json_blue", "status_long_term_suspension_check.json")
CHECK_NAME = "Long_Term_Suspension_Check"
SHAREPOINT_LOCAL_PATH = r"C:\Users\kdimitriou\OneDrive - MORNINGSTAR INC\Indexes Global Operations-Daily Operations - Daily Operations\Index Operations\kostas_tests"
TEMP_EXTRACT_PATH = os.path.join(os.environ["TEMP"], "validation_temp_attachments")
# ============================================================
# Step 1 — open Outlook and look at the inbox
# ============================================================
outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)  # newest email first

# ============================================================
# Step 2 — find today's email with the matching subject
# (subject has a dynamic date at the end, so we only match the static part)
# ============================================================
today_email = None
today = date.today()

for msg in messages:
    if TARGET_SUBJECT in msg.Subject and msg.ReceivedTime.date() == today:
        today_email = msg
        break

# ============================================================
# Step 3 — unzip the attachment, read the CSV, check it
# ============================================================
saved_html_filename = None

if today_email is None:
    suspension_status = "ERROR: report email not found for today"
else:
    received_date = today_email.ReceivedTime.strftime("%d/%m/%Y %H:%M")

    # --- Save and unzip the attachment ---
    os.makedirs(TEMP_EXTRACT_PATH, exist_ok=True)
    zip_path = None
    for attachment in today_email.Attachments:
        if attachment.FileName.lower().endswith(".zip"):
            zip_path = os.path.join(TEMP_EXTRACT_PATH, attachment.FileName)
            attachment.SaveAsFile(zip_path)
            break

    if zip_path is None:
        suspension_status = "ERROR: no zip attachment found in email"
        securities_table_html = "<p>No zip attachment found.</p>"
    else:
        with zipfile.ZipFile(zip_path, "r") as z:
            z.extractall(TEMP_EXTRACT_PATH)
            data_filename = [n for n in z.namelist() if n.lower().endswith(".xlsx")][0]

        data_path = os.path.join(TEMP_EXTRACT_PATH, data_filename)
        df = pd.read_excel(data_path)

        if df.empty:
            suspension_status = f"clear: no securities to be dropped (received {received_date})"
            securities_table_html = "<p>No securities to be dropped.</p>"
        else:
            tickers = df["ticker"].tolist() if "ticker" in df.columns else df.iloc[:, 3].tolist()
            suspension_status = f"CHECK: {len(df)} security(ies) to be dropped — {tickers} (received {received_date})"
            securities_table_html = df.to_html(index=False)

    # --- Build the saved HTML report (subject + sent time + securities table) ---
    header_html = f"""
    <div style='font-family:Calibri,Arial,sans-serif; margin-bottom:16px; padding-bottom:8px; border-bottom:2px solid #333;'>
        <h2 style='margin:0;'>{today_email.Subject}</h2>
        <p style='margin:4px 0 0 0; color:#555;'>Sent: {received_date}</p>
    </div>
    """

    body_html = f"""
    <div style='font-family:Calibri,Arial,sans-serif;'>
        <h3>Securities to be dropped</h3>
        {securities_table_html}
    </div>
    """

    saved_html_filename = f"Long_Term_Suspension_Report_{today_email.ReceivedTime.strftime('%Y-%m-%d')}.html"
    html_path = os.path.join(SHAREPOINT_LOCAL_PATH, saved_html_filename)
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(header_html + body_html)

# ============================================================
# Step 4 — write the result, same format as your other checks
# ============================================================
status_data = {
    "check_name": CHECK_NAME,
    "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    "checks": {
        "Securities to be dropped": suspension_status
    },
    "saved_html_filename": saved_html_filename
}

os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)

with open(OUTPUT_FILE, "w") as f:
    json.dump(status_data, f, indent=2)

print(json.dumps(status_data, indent=2))