import mariadb
import sys
import pandas as pd
import json
from datetime import date, datetime, timedelta
import os

# To display all columns
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 1000)

# ============================================================
# Config
# ============================================================
LOOKBACK_DAYS = 2            # T, T-1, T-2 -> 3 calendar days total
START_END_MAX_MINUTES = 10    # max allowed start_time -> end_time duration
GAP_MAX_MINUTES = 20          # max allowed gap between consecutive end_time entries

OUTPUT_FILE = os.path.join("json_blue", "status_corporate_action_import_run.json")
CHECK_NAME = "Corporate_Action_Import_Run_Check"

today = date.today()
window_start = today - timedelta(days=LOOKBACK_DAYS)   # T-2, 00:00:00

# ============================================================
# Connecting to ticks DB
# ============================================================
try:
    conn = mariadb.connect(
        user='ops_pending_ca_tracker',
        password='zmUFk8Xxcc8WrycCrMVe0plI84KkQAmqaOwoKHMn',
        host='mb-prd-db.us-east-1.mif0286.eas.morningstar.com',
        port=3306,
        database='ticks'
    )
except mariadb.Error as e:
    print(f"Error connecting to MariaDB Platform: {e}")
    sys.exit(1)

cur = conn.cursor()

# ============================================================
# Query 1 — last 3 calendar days of import runs (T, T-1, T-2)
# ============================================================
sql_recent = """
    SELECT id, start_time, end_time, total_securities, done_securities,
           no_of_cas, import_after, status
    FROM corporate_action_backend.corporate_action_import_run
    WHERE end_time >= %s
    ORDER BY end_time DESC
"""
cur.execute(sql_recent, (window_start,))
rows = cur.fetchall()
cols = [d[0] for d in cur.description]
df = pd.DataFrame(rows, columns=cols)

# MariaDB connector can return numeric columns as decimal.Decimal, which
# breaks pandas/numpy arithmetic (std, mean, subtraction, etc.) -> cast to float
for c in ["total_securities", "done_securities", "no_of_cas"]:
    if c in df.columns:
        df[c] = df[c].astype(float)


conn.close()

# ============================================================
# Guard: no data at all in the lookback window
# ============================================================
if df.empty:
    status_data = {
        "check_name": CHECK_NAME,
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "checks": {
            "Latest run status": "CHECK: no rows found in last 3 days",
        }
    }
    with open(OUTPUT_FILE, "w") as f:
        json.dump(status_data, f, indent=2)
    print(json.dumps(status_data, indent=2))
    sys.exit(0)

df = df.sort_values("end_time", ascending=False).reset_index(drop=True)
df["start_time"] = pd.to_datetime(df["start_time"])
df["end_time"] = pd.to_datetime(df["end_time"])

# ============================================================
# CHECK 1 — status of the latest run (row 0, table already DESC)
# ============================================================
latest_row = df.iloc[0]
if latest_row["status"] == "success":
    latest_status = f"CLEAR: id {latest_row['id']} status=success"
else:
    latest_status = f"CHECK: id {latest_row['id']} status={latest_row['status']}"

# ============================================================
# CHECK 2 — any entries with status != 'success' -> flag ids
# ============================================================
bad_status_ids = df.loc[df["status"] != "success", "id"].tolist()
non_success_status = (
    f"CHECK: ids {bad_status_ids}" if bad_status_ids
    else "CLEAR: all runs status=success"
)

# ============================================================
# CHECK 3 — Daily sum of CAs
# ============================================================
today_cas = df.loc[df["end_time"].dt.date == today, "no_of_cas"].sum()
volume_status = f"Today's CA count: {int(today_cas)}"

# ============================================================
# CHECK 4 — total_securities == done_securities for every entry
# NaN-safe: two NULLs are not treated as a mismatch, but a NULL
# paired with a real value IS flagged (that's a genuine anomaly)
# ============================================================
both_null = df["total_securities"].isna() & df["done_securities"].isna()
differs = df["total_securities"].ne(df["done_securities"])  # NaN != NaN -> True here, hence the exclusion below
mismatch_mask = differs & ~both_null
mismatch_ids = df.loc[mismatch_mask, "id"].tolist()
securities_status = (
    f"CHECK: ids {mismatch_ids}" if mismatch_ids
    else "CLEAR: total_securities == done_securities for all runs"
)

# ============================================================
# CHECK 5 — start_time -> end_time duration > 10 minutes -> flag id
# ============================================================
df["duration_min"] = (df["end_time"] - df["start_time"]).dt.total_seconds() / 60
long_run_ids = df.loc[df["duration_min"] > START_END_MAX_MINUTES, "id"].tolist()
duration_status = (
    f"CHECK: ids {long_run_ids} (run duration > {START_END_MAX_MINUTES} min)" if long_run_ids
    else f"CLEAR: all runs completed within {START_END_MAX_MINUTES} min"
)

# ============================================================
# CHECK 6 — gap between consecutive end_time entries > 20 minutes
# (id flagged = the later entry of the pair with the large gap)
# ============================================================
df_sorted = df.sort_values("end_time", ascending=True).reset_index(drop=True)
df_sorted["gap_min"] = df_sorted["end_time"].diff().dt.total_seconds() / 60
df_sorted["prev_id"] = df_sorted["id"].shift(1)

gap_pairs_df = df_sorted.loc[df_sorted["gap_min"] > GAP_MAX_MINUTES, ["prev_id", "id"]]
gap_pairs = [f"{int(r.prev_id)}->{int(r.id)}" for r in gap_pairs_df.itertuples()]

gap_status = (
    f"CHECK: id pairs {gap_pairs}" if gap_pairs
    else f"CLEAR"
)

# ============================================================
# Write status JSON (same format as the other check scripts)
# ============================================================status_data = {
status_data = {
    "check_name": CHECK_NAME,
    "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    "checks": {
        "Latest run status": latest_status,
        "Non-success entries (last 3 days)": non_success_status,
        "Today's CA volume": volume_status,
        "Total vs done securities mismatch": securities_status,
        "Start/End duration > 10min": duration_status,
        "Gap between runs > 20min": gap_status,
    }
}

with open(OUTPUT_FILE, "w") as f:
    json.dump(status_data, f, indent=2)

print(json.dumps(status_data, indent=2))