import smtplib
import time
from datetime import datetime
import sys
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# --- CONFIGURATION ---
GMAIL_USER = "kostasdim520@gmail.com"
GMAIL_APP_PASSWORD = "skfy qcwt khna ocnc"   # 16-char Google App Password
EXCEL_FILE = "ECB_Heads_of_Section_FullList.xlsx"
SHEET_NAME = "ECB Heads of Section"
PDF_ATTACHMENT = "CV DIMITRIOU SF.pdf"
LOG_FILE = "send_log.csv"

EMAIL_SUBJECT = "Expression of interest for short-term contract opportunities"

EMAIL_BODY_TEMPLATE = """Dear {first_name},

I hope this message finds you well.
 
I recently completed my traineeship at the DG Monetary Policy (MAY) and am looking for a short-term contract opportunity, ideally running through the end of the year. I wanted to reach out regarding {division} .
 
My profile combines two things I'd like to keep doing: working with granular financial data, and building automation around it. Currently i am employed at Morningstar, working on an AI-agent pipeline that autonomously monitors daily news for regulatory milestones on pending corporate actions. It ingests analysts' notes as context, conducts the research, and produces a summarised email report that replaces a manual daily check. I also recently automated a list of daily manual data validation checks, generating recurring reports without human intervention, and setting up automated email dispatch of those outputs to stakeholders.

During my traineeship I worked extensively with security holdings (SHSS), bank balance sheets (Orbis, Bank Lending Survey)  and multiple internal datasets, cleansing, joining, aggregating and visualising. This work was targeted at the maintenance and extension of the division's dashboard infrastructure through Python and Git-based collaboration.
 
This kind of work, data pipelines and process automation, is specifically what I want to continue doing, and it is the type of contribution I believe I could deliver quickly within your section over a short-term contract. My CV is attached.
 
Thank you for your time and consideration.
 
Kind regards,
Kostas Dimitriou

"""

DELAY_SECONDS = 4       # pause between sends, avoids looking like a spam burst
DRY_RUN = False          # flip to False only after checking the printed output
MAX_SENDS = None        # e.g. 5, to test on a slice before running the full list
# ---------------------

# --- LOAD CONTACTS ---
#df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)

df = pd.DataFrame([{
    "First Name": "Kostas",
    "Division / Section": "Test Division",
    "Email Address": "kostas.dimitriou@morningstar.com",
}])

required_cols = {"First Name", "Division / Section", "Email Address"}
missing = required_cols - set(df.columns)
if missing:
    sys.exit(f"Excel file is missing required columns: {missing}")

if MAX_SENDS:
    df = df.head(MAX_SENDS)

# --- READ ATTACHMENT ONCE (same file for every email, no need to re-read per send) ---
with open(PDF_ATTACHMENT, "rb") as f:
    pdf_bytes = f.read()

# --- CONNECT TO GMAIL (skipped entirely in dry-run mode) ---
server = None
if not DRY_RUN:
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(GMAIL_USER, GMAIL_APP_PASSWORD)

# --- SEND LOOP ---
results = []

for index, row in df.iterrows():
    recipient_email = str(row["Email Address"]).strip()
    first_name = str(row["First Name"]).strip()
    division = str(row["Division / Section"]).strip()
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")


    status = "DRY_RUN"
    error = ""

    try:
        body = EMAIL_BODY_TEMPLATE.format(first_name=first_name, division=division)

        msg = MIMEMultipart()
        msg["From"] = GMAIL_USER
        msg["To"] = recipient_email
        msg["Subject"] = EMAIL_SUBJECT
        msg.attach(MIMEText(body, "plain"))

        part = MIMEBase("application", "octet-stream")
        part.set_payload(pdf_bytes)
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename={PDF_ATTACHMENT}")
        msg.attach(part)

        if not DRY_RUN:
            server.send_message(msg)
            status = "SENT"

    except Exception as e:
        status = "FAILED"
        error = str(e)

    print(f"[{index + 1}/{len(df)}] {status}: {first_name} <{recipient_email}> {error}")

    results.append({
        "Timestamp": timestamp,
        "First Name": first_name,
        "Email Address": recipient_email,
        "Division / Section": division,
        "Status": status,
        "Error": error,
    })

    if not DRY_RUN and status == "SENT":
        time.sleep(DELAY_SECONDS)

if server is not None:
    server.quit()

# --- SUMMARY ---
pd.DataFrame(results).to_csv(LOG_FILE, index=False)

sent = sum(1 for r in results if r["Status"] == "SENT")
failed = sum(1 for r in results if r["Status"] == "FAILED")
dry = len(results) - sent - failed

print(f"\nDone. Sent: {sent}, Failed: {failed}, Dry-run rows: {dry}")
print(f"Full log written to {LOG_FILE}")