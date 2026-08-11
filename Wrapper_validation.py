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

print("Running CBOE RBICS comparison check...")
subprocess.run([sys.executable, "CBOE_RBICS_comparison_Kostas.py"])

print("Running duplicate open timespans check...")
subprocess.run([sys.executable, "Duplicate_open_timespans_Kostas.py"])

print("Running day shift QC check...")
subprocess.run([sys.executable, "Day_shift_QC_check_Kostas.py"])

# ============================================================
# STEP 2 — Read the JSON status file that DCAF script produced
# ============================================================
with open("status_dcaf_check.json", "r") as f:
    dcaf_result = json.load(f)

with open("status_CBOE_RBICS_Comparison.json", "r") as f:
    rbics_result = json.load(f)

with open("status_duplicate_open_timespans.json", "r") as f:
    duplicate_timespans_result = json.load(f)

with open("status_Day_shift_QC_checks.json", "r") as f:
    day_shift_qc_result = json.load(f)



# ============================================================
# STEP 3 — Build the email table rows from the JSON
# ============================================================
all_results = [dcaf_result, rbics_result,duplicate_timespans_result, day_shift_qc_result]

table_rows = ""
for result in all_results:
    script_name = result["check_name"]  # e.g. "DCAF_CA_Check"
    for check, status in result["checks"].items():
        color = "red" if "CHECK" in status else "green"
        table_rows += f"""
        <tr>
            <td style='padding:8px;'>{script_name}</td>
            <td style='padding:8px; color:{color};'><b>{status}</b></td>
        </tr>"""

timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")

email_body = f"""
<h3>[{CATEGORY}] Daily Validation Report - {timestamp}</h3>
<table border='1' cellpadding='8' style='border-collapse:collapse;'>
    <tr style='background-color:#e0e0e0;'>
        <th style='padding:8px;'>Task</th>
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