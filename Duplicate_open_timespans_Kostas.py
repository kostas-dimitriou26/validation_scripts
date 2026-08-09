import mariadb
import sys
import os
import pandas as pd
import json
from datetime import date, datetime




#To display all columns
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 1000)


# Connecting to ticks DB
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

# Fetch SQL dates
sql_dates = """SELECT
    CURDATE() AS t_date,
    CASE
        WHEN DAYOFWEEK(CURDATE()) = 2 THEN CURDATE() - INTERVAL 3 DAY
        ELSE CURDATE() - INTERVAL 1 DAY
    END AS t_1_date;"""

cur.execute(sql_dates)
t_date_actual, t_1_date_actual = cur.fetchone()

# --- Query 1: duplicate variation_ids on current date ---
sql_t = """select variation_id, count(variation_id) as ct from index_timespan it 
where valid_from = CURDATE()
and `type` = 'Open'
group by variation_id
having count(variation_id) > 1
ORDER by ct desc;"""

cur.execute(sql_t)
result_t = cur.fetchall()
t_date = pd.DataFrame(result_t, columns=["variation_id", "ct"])

# --- Query 2: duplicate variation_ids on t-1 bd ---
sql_t_1 = """SELECT
    variation_id,
    COUNT(*) AS ct
FROM index_timespan it
WHERE valid_from =
    CASE
        WHEN DAYOFWEEK(CURDATE()) = 2 THEN CURDATE() - INTERVAL 3 DAY
        ELSE CURDATE() - INTERVAL 1 DAY
    END
AND `type` = 'Close'
GROUP BY variation_id
HAVING count(*) > 1
ORDER BY ct DESC;"""

cur.execute(sql_t_1)
result_t_1 = cur.fetchall()
t_1_date = pd.DataFrame(result_t_1, columns=["variation_id", "ct"])

conn.close()

t_status = 'T:CHECK' if not t_date.empty else 'clear'
t_1_status = 'T:CHECK' if not t_1_date.empty else 'clear'

combined_status = f"T: {t_date_actual.strftime('%d/%m/%Y')} {t_status} | T-1: {t_1_date_actual.strftime('%d/%m/%Y')} {t_1_status}"

status_data = {
    "check_name": "Index_Timespan_Duplicate_Check",
    "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    "checks": {
        "Duplicate open timespans": combined_status,
    }
}
with open("status_duplicate_open_timespans.json", "w") as f:
    json.dump(status_data, f, indent=2)