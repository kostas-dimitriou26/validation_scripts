import sys
from datetime import datetime, timedelta, date
import json
import win32com.client as win32
import subprocess


CATEGORY = "BLUE"
RECIPIENTS = "kostas.dimitriou@morningstar.com"

RT_MISMATCH_HOUR = 16
RT_MISMATCH_MINUTE = 30
TOLERANCE_MINUTES = 20

def is_due(hh, mm, now):
    scheduled = now.replace(hour=hh, minute=mm, second=0, microsecond=0)
    return abs((now - scheduled).total_seconds()) <= TOLERANCE_MINUTES * 60

def run_check(script_name, status_file, display_name):
    print(f"Running {display_name}...")
    result = subprocess.run([sys.executable, script_name])
    if result.returncode != 0:
        print(f"  -> {display_name} FAILED (exit code {result.returncode})")
        return {"check_name": display_name, "checks": {display_name: "ERROR"}}
    try:
        with open(status_file, "r") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError) as e:
        print(f"  -> {display_name} status file unreadable: {e}")
        return {"check_name": display_name, "checks": {display_name: "ERROR"}}

# ============================================================
# STEP 1 & 2 — CA_backend always runs; RT check only at 16:30
# ============================================================
now = datetime.now()

CA_backend_result = run_check("CA_backend_kostas.py", "status_corporate_action_import_run.json", "CA_backend_Check")

if is_due(RT_MISMATCH_HOUR, RT_MISMATCH_MINUTE, now):
    RT_mismatch = run_check("RT loading qc check_kostas.py", "RT_loading_qc_check.json", "RT_loading_qc_Check")
else:
    print("Skipping RT_loading_qc_Check — not due until 16:30")
    RT_mismatch = {"check_name": "RT_loading_qc_Check", "checks": {"Skipping": "to be run at 16:30"}}

# ============================================================
# STEP 3 — Build the email table rows from the JSON
# ============================================================
all_results = [CA_backend_result, RT_mismatch]

table_rows = ""
for result in all_results:
    script_name = result["check_name"]

    status_lines = []
    for check, status in result["checks"].items():
        if status == "ERROR":
            color = "blue"
        elif status=='to be run at 16:30':
            color = "black"
        elif "CHECK" in status:
            color = "red"
        else:
            color = "green"
        status_lines.append(f"<span style='color:{color};'><b>{check}:</b> {status}</span>")

    combined_status_html = "<br>".join(status_lines)

    table_rows += f"""
        <tr>
            <td style='padding:8px; vertical-align:top;'>{script_name}</td>
            <td style='padding:8px;'>{combined_status_html}</td>
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