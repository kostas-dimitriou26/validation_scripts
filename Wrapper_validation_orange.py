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
# Helper — run a check script, and safely load its JSON status
# ============================================================
def run_check(script_name, status_file, display_name):
    print(f"Running {display_name}...")
    result = subprocess.run([sys.executable, script_name])

    if result.returncode != 0:
        print(f"  -> {display_name} FAILED (exit code {result.returncode})")
        return {
            "check_name": display_name,
            "checks": {display_name: "ERROR"}
        }

    try:
        with open(status_file, "r") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError) as e:
        print(f"  -> {display_name} status file unreadable: {e}")
        return {
            "check_name": display_name,
            "checks": {display_name: "ERROR"}
        }

# ============================================================
# STEP 1 & 2  — run each check, load its status safely
# ============================================================
dcaf_result = run_check("CBOE_DCAF_kostas.py", "status_dcaf_check.json", "DCAF_CA_Check")
rbics_result = run_check("CBOE_RBICS_comparison_Kostas.py", "status_CBOE_RBICS_Comparison.json", "CBOE_RBICS_Comparison")
duplicate_timespans_result = run_check("Duplicate_open_timespans_Kostas.py", "status_duplicate_open_timespans.json", "Duplicate_Open_Timespans")
day_shift_qc_result = run_check("Day_shift_QC_check_Kostas.py", "status_Day_shift_QC_checks.json", "Day_Shift_QC_Check")
IPO_listings = run_check("new_IPO_listings.py", "status_ipo_despac_check.json", "IPO_Listings")

# ============================================================
# STEP 3 — Build the email table rows from the JSON
# ============================================================
all_results = [dcaf_result, rbics_result, duplicate_timespans_result, day_shift_qc_result, IPO_listings]

table_rows = ""
for result in all_results:
    script_name = result["check_name"]

    entry_id = result.get("email_entry_id")
    link_html = f' <a href="outlook:{entry_id}">[Open email]</a>' if entry_id else ""

    for check, status in result["checks"].items():
        if status == "ERROR":
            color = "orange"
        elif "CHECK" in status:
            color = "red"
        else:
            color = "green"
        table_rows += f"""
        <tr>
            <td style='padding:8px;'>{script_name}</td>
            <td style='padding:8px; color:{color};'><b>{status}</b>{link_html}</td>
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
# ============================================================
outlook = win32.gencache.EnsureDispatch('Outlook.Application')
mail = outlook.CreateItem(0)
mail.To = RECIPIENTS
mail.Subject = f"[{CATEGORY}] Daily Validation Report - {timestamp}"
mail.HTMLBody = email_body
mail.Send()

print(f"Done! [{CATEGORY}] summary email sent to {RECIPIENTS}")