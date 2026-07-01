import pandas as pd
from pathlib import Path
import os
import mariadb
import datetime
import time
import sys
from pandas.tseries.offsets import *
import win32com.client as win32
import numpy as np
from pandas.tseries.offsets import CustomBusinessDay
from datetime import date
from datetime import datetime, timedelta, date
import json
import win32com.client as win32
import subprocess


CATEGORY = "ORANGE"
RECIPIENTS = "kostas.dimitriou@morningstar.com"
# ============================================================
# STEP 1 — Run the DCAF script and wait for it to finish
# ============================================================
print("Running DCAF check...")
subprocess.run([sys.executable, "CBOE_DCAF_kostas.py"])

# ============================================================
# STEP 2 — Read the JSON status file that DCAF script produced
# ============================================================
with open("status_dcaf_check.json", "r") as f:
    dcaf_result = json.load(f)

# dcaf_result["checks"] looks like:
# { "new_CAs": "OK", "send_notice_today": "CHECK" }

# ============================================================
# STEP 3 — Build the email table rows from the JSON
# ============================================================
all_results = [dcaf_result]

table_rows = ""
for result in all_results:
    script_name = result["check_name"]  # e.g. "DCAF_CA_Check"
    for check, status in result["checks"].items():
        color = "red" if status == "CHECK" else "green"
        table_rows += f"""
        <tr>
            <td style='padding:8px;'>{script_name}</td>
            <td style='padding:8px;'>{check}</td>
            <td style='padding:8px; color:{color};'><b>{status}</b></td>
        </tr>"""

timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")

email_body = f"""
<h3>[{CATEGORY}] Daily Validation Report - {timestamp}</h3>
<table border='1' cellpadding='8' style='border-collapse:collapse;'>
    <tr style='background-color:#e0e0e0;'>
        <th style='padding:8px;'>Task</th>
        <th style='padding:8px;'>Check</th>
        <th style='padding:8px;'>Status</th>
    </tr>
    {table_rows}
</table>
"""

# ============================================================
# STEP 4 — Send the email via Outlook
# Subject contains [ORANGE] so Outlook rule can sort it automatically
# ============================================================
outlook = win32.gencache.EnsureDispatch('Outlook.Application')
mail = outlook.CreateItem(0)
mail.To = RECIPIENTS
mail.Subject = f"[{CATEGORY}] Daily Validation Report - {timestamp}"
mail.HTMLBody = email_body
mail.Send()

print(f"Done! [{CATEGORY}] summary email sent to {RECIPIENTS}")