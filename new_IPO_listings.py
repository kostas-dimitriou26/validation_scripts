import win32com.client as win32
import pandas as pd
import json
import io
from datetime import date, datetime,timedelta
import os

# ============================================================
# Config
# ============================================================
TARGET_SUBJECT = "IPOs Listed Today and De-SPAC Transactions Anticipated Tomorrow"
OUTPUT_FILE = "status_ipo_despac_check.json"
CHECK_NAME = "IPO_DeSPAC_Ticker_Check"
SHAREPOINT_LOCAL_PATH = r"C:\Users\kdimitriou\OneDrive - MORNINGSTAR INC\Indexes Global Operations-Daily Operations - Daily Operations\Index Operations\kostas_tests"


# ============================================================
# Step 1 — open Outlook and look at the inbox
# ============================================================
outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)  # newest email first

# ============================================================
# Step 2 — find today's email with the matching subject
# ============================================================
today_email = None
yesterday = date.today() - timedelta(days=1)

for msg in messages:
    if TARGET_SUBJECT in msg.Subject and msg.ReceivedTime.date() == yesterday:
        today_email = msg
        break

# ============================================================
# Step 3 — read the IPO table, check it, and save a shared copy
# ============================================================
saved_html_filename = None

if today_email is None:
    ipo_status = "ERROR: report email not found for today"
else:
    tables = pd.read_html(io.StringIO(today_email.HTMLBody))
    ipo_table = tables[0]  # first table in the email = IPO Ticker table
    first_cell = str(ipo_table.iloc[1, 0])
    received_date = today_email.ReceivedTime.strftime("%d/%m/%Y")

    # Build a small header block with subject + sent time, then the original email content
    header_html = f"""
       <div style='font-family:Calibri,Arial,sans-serif; margin-bottom:16px; padding-bottom:8px; border-bottom:2px solid #333;'>
           <h2 style='margin:0;'>{today_email.Subject}</h2>
           <p style='margin:4px 0 0 0; color:#555;'>Sent: {received_date}</p>
       </div>
       """

    # Save the email body as a readable HTML file in the shared folder
    saved_html_filename = f"IPO_DeSPAC_Report_{today_email.ReceivedTime.strftime('%Y-%m-%d')}.html"
    html_path = os.path.join(SHAREPOINT_LOCAL_PATH, saved_html_filename)
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(header_html + today_email.HTMLBody)

    if "none" in first_cell.lower():
        ipo_status = f"clear: no new IPO tickers (received {received_date})"
    else:
        ipo_status = f"CHECK: new IPO ticker reported — {first_cell} (received {received_date})"

# ============================================================
# Step 4 — write the result, same format as your other checks
# ============================================================
status_data = {
    "check_name": CHECK_NAME,
    "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    "checks": {
        "New IPO ticker": ipo_status
    },
    "saved_html_filename": saved_html_filename
}

with open(OUTPUT_FILE, "w") as f:
    json.dump(status_data, f, indent=2)

print(json.dumps(status_data, indent=2))