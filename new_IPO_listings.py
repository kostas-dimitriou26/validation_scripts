import win32com.client as win32
import pandas as pd
import json
import io
from datetime import date, datetime,timedelta

# ============================================================
# Config
# ============================================================
TARGET_SUBJECT = "IPOs Listed Today and De-SPAC Transactions Anticipated Tomorrow"
OUTPUT_FILE = "status_ipo_despac_check.json"
CHECK_NAME = "IPO_DeSPAC_Ticker_Check"

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

entry_id = today_email.EntryID
store_id = today_email.Parent.StoreID  # the folder's StoreID
email_link = f'<a href="outlook:{entry_id}">Open email</a>'
#print("Subject:", today_email.Subject)
#print("Received:", today_email.ReceivedTime)
#print("HTMLBody length:", len(today_email.HTMLBody))

# ============================================================
# Step 3 — read the IPO table from inside the email and check it
# ============================================================
received_date = today_email.ReceivedTime.strftime("%d/%m/%Y")
if today_email is None:
    ipo_status = "ERROR: report email not found for today"
else:
    tables = pd.read_html(io.StringIO(today_email.HTMLBody))
    ipo_table = tables[0]  # first table in the email = IPO Ticker table
    first_cell = str(ipo_table.iloc[1, 0])

    checked_date = today_email.ReceivedTime.strftime("%d/%m/%Y")

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
    "email_entry_id": today_email.EntryID if today_email else None
}

with open(OUTPUT_FILE, "w") as f:
    json.dump(status_data, f, indent=2)

print(json.dumps(status_data, indent=2))