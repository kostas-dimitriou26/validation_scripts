import os
import pandas as pd
import mariadb
import psycopg2
import sys
import json
import numpy as np
from datetime import date, timedelta,datetime

#Set dates
today = date.today()

# Next business day: skip Sat/Sun only (no holiday calendar)
next_bd = today + timedelta(days=1)
while next_bd.weekday() >= 5:  # 5=Saturday, 6=Sunday
    next_bd += timedelta(days=1)

eod_date = f"""'{today.strftime('%Y-%m-%d')}'"""#today
open_date = f"""'{next_bd.strftime('%Y-%m-%d')}'"""#Tomorrow / next BD

OUTPUT_FILE = os.path.join("json_blue", "RT_loading_qc_check.json")
CHECK_NAME = "RT_loading_qc_Check"

#To display all columns
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 1000)



try:
    conn1= psycopg2.connect(
        host='mi-prd-db.us-west-2.mif0286.eas.morningstar.com',
        database='mi_prd_db',
        user='kostas_dimitriou',
        password='aFcXw0aPcE4PTrF2',
        port=5432

    )
except psycopg2.Error as e:
    print(f"Error connecting to Cirrus Platform: {e}")
    sys.exit(1)

# Get Cursor for cirrus
cur1 = conn1.cursor()

helix_close="""select distinct pc.performance_id, vpd.val, vpd3.val
from market_data.security_of_interest pc
join refdata.id_mapping im on pc.performance_id=im.performance_id
join refdata.vendor_prioritized_data vpd on pc.performance_id =vpd.entity_id and vpd.attribute_code ='ric' and vpd.start_date<="""+open_date+""" and vpd.end_date >= """+eod_date+""" 
join refdata.vendor_prioritized_data vpd3 on pc.performance_id =vpd3.entity_id and vpd3.attribute_code ='Currency' and vpd3.start_date<="""+open_date+""" and vpd3.end_date >= """+eod_date+""" 
where pc.end_date>=now()
and pc.priority in (1,2)
and im.universe_code in ('FC', 'ST')
and pc.performance_id not like ('%DUMMY%')
and NOT EXISTS (
    SELECT 1
    FROM cirrus_corporate_action.proforma_intermediate s
    WHERE s.performance_id = pc.performance_id
    and s.portfolio_key in ('3460199','3460354', '3501351','3501354','3501357','3501367','3501369','3501371')
    and s.effective_date >='2026-06-22'
);"""

cur1.execute(helix_close)
result1=cur1.fetchall()
mi_prod_data=pd.DataFrame(result1)
mi_prod_data.columns = ['pid', 'mi_ric', 'mi_currency']
mi_prod_data['mi_currency'] = mi_prod_data['mi_currency'].replace({'CNY': 'CNH'})
conn1.close()

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

# Get Cursor for ticks
cur = conn.cursor()

# to get lc table values

sql_ticks = """select distinct lc.listing_id,lc2.string_value,lc3.currency 
from listing_characteristic lc 
join listing l ON  lc.listing_id =l.listing_id
join `security` s on s.security_id=l.security_id 
join listing_characteristic lc2 on lc.listing_id=lc2.listing_id and lc2.name='RIC' and lc2.valid_to>="""+eod_date+"""
join listing_characteristic lc3 on lc.listing_id=lc3.listing_id and lc3.name='Currency' and lc3.valid_to>="""+eod_date+"""
where s.security_type in ('Common Stock', 'DR','Unit', 'ETF');"""

cur.execute(sql_ticks)

result = cur.fetchall()

ticks_data = pd.DataFrame(result)

conn.close()

ticks_data.columns = ['listing_id', 'ticks_ric', 'ticks_currency']
ticks_data['ticks_currency'] = ticks_data['ticks_currency'].replace({'ILA': 'ILS', 'GBX': 'GBP', 'CNY': 'CNH','KWF':'KWD','ZAC': 'ZAR'})


#lookup for listing ids using pi ids
full_file=pd.merge(mi_prod_data,ticks_data, how='left',left_on='mi_ric', right_on='ticks_ric')

full_file['ric_check']=full_file['mi_ric']==full_file['ticks_ric']
full_file['currency_check']=full_file['mi_currency']==full_file['ticks_currency']

#filter false entries
filtered_df = full_file[~(full_file['ric_check'])]
#filtered_df.to_excel('missing_ric.xlsx')
#print(filtered_df)

if filtered_df.notnull().any().any():
    print('')
    print('Please verify the mismatches using pid. You can use useful queries > security characteristics using pd or directly query in the vpd table.\n')
    print(filtered_df)
else:
    print('')
    print('No mismatches for today')

mismatch_status = 'CHECK' if not filtered_df.empty else 'clear'

status_data = {
    "check_name": CHECK_NAME,
    "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    "checks": {
        "RT loading mismatch": mismatch_status
    }
}

with open(OUTPUT_FILE, "w") as f:
    json.dump(status_data, f, indent=2)