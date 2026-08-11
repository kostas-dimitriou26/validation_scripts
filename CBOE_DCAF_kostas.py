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
from datetime import datetime, timedelta
import json

pd.set_option('display.max_columns', None)
pd.set_option('display.width', 1000)

###### SET DATE ######
#date = '2026-06-29'  #insert date for T-1

# Get t-1 BD
date_aux=date.today()

while True:
    date_aux = date_aux - timedelta(days=1)
    if date_aux.weekday() < 5:  # Monday=0, Sunday=6
        break
date_aux=date_aux.strftime("%Y-%m-%d")
date=date_aux
del date_aux

#Dates for T
currdate= time.strftime("%Y-%m-%d")

#Path to DCAF file!Q" (with exception)
path1 = os.path.expanduser("~\\OneDrive - MORNINGSTAR INC\\Indexes Global Operations-Daily Operations - Daily Operations\\Index Operations\\CBOE\\Checks\\Moorgate CA Files\\")
path2 = os.path.expanduser("~\\OneDrive - MORNINGSTAR INC\\Daily Operations\\Index Operations\\CBOE\\Checks\\Moorgate CA Files\\")

if os.path.exists(path1):
    path = path1
elif os.path.exists(path2):
    path = path2
else:
    raise FileNotFoundError("Neither path exists")

path_mgfile=path +'corporateactions-'+date+'.csv'

#Connecting to DB
try:
    conn = mariadb.connect(
    user='devashish.gantayat',
    password='+Pq;yiGgCCY1~8GuVKVvCPRMDITd3dW`oN|pS@Sm',
    host='mb-prd-db.us-east-1.mif0286.eas.morningstar.com',
    port=3306,
    database='ticks'
    )
except mariadb.Error as e:
    print(f"Error connecting to MariaDB Platform: {e}")
    sys.exit(1)

df_CA = pd.read_csv(path_mgfile, sep=';', dayfirst=False)

DCAF_size = os.path.getsize(path_mgfile)*0.000977
if DCAF_size< 100:
    print("File size should be around 100 kb, open and check the contents if not")

print("DACF File Size is :", round(DCAF_size,2), "kilo bytes")

print(" ")

#custom business day calcs using calendar id 134

date1 = pd.to_datetime(date) - BusinessDay(30)
date2 = pd.to_datetime(date) + BusinessDay(30)

date1_str = f"'{date1.strftime('%Y-%m-%d')}'"

date2_str = f"'{date2.strftime('%Y-%m-%d')}'"

cur2 = conn.cursor()

calendar_134 = """select date from calendar_day cd where calendar_id =134 and `date` between """ + date1_str + """ and """ + date2_str + """;"""

cur2.execute(calendar_134)

result1 = cur2.fetchall()

calendar = pd.DataFrame(result1)

calendar.columns = ["holidays"]

holidays = pd.to_datetime(calendar['holidays'])

# Create a CustomBusinessDay object

custom_bd = CustomBusinessDay(holidays = holidays)

df_CA = df_CA[df_CA['Status'] == 'confirmed'].copy()
df_CA['EffectiveDate'] = pd.to_datetime(df_CA['EffectiveDate'], format='%Y%m%d')
df_CA['notice_date'] = df_CA['EffectiveDate'].apply(lambda x: x - 3 * custom_bd)
df_CA['Send_notice']=df_CA['notice_date']==currdate
df_CA['Send_notice_today']=df_CA['Send_notice'].replace({True:'Yes',False:'No'})

#To get the details of mergers, delistings

merger = df_CA.CorpActionType.str.contains('merger')
takeover = df_CA.CorpActionType.str.contains('takeover')
delist = df_CA.CorpActionCategory.str.contains('delist')
spinf_off=df_CA.CorpActionType.str.contains('demerger')


df_CAs_filtered=df_CA[(df_CA['Status']=='confirmed') & (merger | takeover| delist|spinf_off) ]

final_df=df_CAs_filtered[['ISIN','RIC','ConstituentName','CorpActionType','EffectiveDate','notice_date','Send_notice_today']]
final_df=final_df.drop_duplicates()

#saving the output file
path_string = os.path.expanduser('~') + \
              '\\OneDrive - MORNINGSTAR INC\\Daily Operations\\Index Operations\\CBOE\\Checks\\'
localpathlog = Path(path_string + 'CA validation')
outputfilename = localpathlog / 'cafile-comparison-{0}.xlsx'.format(date)

# Get Cursor
cur = conn.cursor()

#Get the CA details from DB

sql= """with cte2 as (select
case 
when weekday(CURRENT_DATE())= 0 then date_add(CURRENT_DATE(), INTERVAL 3 day)
when weekday(CURRENT_DATE())= 1 then date_add(CURRENT_DATE(), INTERVAL 3 day)
when weekday(CURRENT_DATE())= 2 then date_add(CURRENT_DATE(), INTERVAL 5 day)
when weekday(CURRENT_DATE())= 3 then date_add(CURRENT_DATE(), INTERVAL 5 day)
when weekday(CURRENT_DATE())= 4 then date_add(CURRENT_DATE(), INTERVAL 5 day)
end as date3),

cte as (select its.security_id, its.listing_id, sc.string_value as ISIN
from index_timespan_security its
join index_timespan it on it.timespan_id = its.timespan_id
join index_timespan_security_aspect itsa on its.timespan_id =itsa.timespan_id and its.listing_id =itsa.listing_id and itsa.name ='Fraction'
join security_characteristic sc on its.security_id = sc.security_id and sc.name = 'ISIN' and CURRENT_DATE() between sc.valid_from and sc.valid_to  
where it.valid_from = CURRENT_DATE() and it.`type` ='Open'
and itsa.double_value <>0
and it.variation_id in (select DISTINCT iv.variation_id from index_variation iv 
join index_masterdata im on im.masterdata_id = iv.masterdata_id 
and im.client_id = 3))

select distinct cab.isin,sc2.string_value, ab.action_name, ab.effective_date, cad.gross_distribution_rate as gross_amount, cad.non_taxable_amount, cad.pid, cad.non_pid,cad.currency, cab.last_updated_at  
from corporate_action_backend.corporate_action_base cab
join security_characteristic sc on cab.isin = sc.string_value and sc.name = 'ISIN' and CURRENT_DATE() between sc.valid_from and sc.valid_to
join security_characteristic sc2 on sc.security_id = sc2.security_id and sc2.name = 'Name' and CURRENT_DATE() between sc2.valid_from and sc2.valid_to
join ticks.action_base ab on sc.security_id = ab.security_id and cab.provider_event_id = ab.provider_action_id  
left join ticks.corporate_action_dividend cad on ab.action_id = cad.action_id
#join cte on cab.isin = cte.ISIN
#join listing_characteristic lc on cte.listing_id=lc.listing_id and lc.name ='RIC' and CURRENT_DATE() between lc.valid_from and lc.valid_to  
where cast(cab.last_updated_at as date) >= CURRENT_DATE() and  cab.status_id= 2 and cab.provider_id = 7 
and cab.isin in (select distinct ISIN from cte)
and ab.effective_date BETWEEN CURRENT_DATE() and (select date3 from cte2)
order by ab.effective_date ;"""

cur.execute(sql)

result=cur.fetchall()

CAs=pd.DataFrame(result)

try:
    CAs.columns = ["ISIN", "security_name","action_name", "effective_date", "gross_amount", "non_taxable_amount", "pid", "non_pid",
                   "currency", "last_updated_at"]
except BaseException as ex:
    print('No new CAs confirmed today for next 3 Bds ')

try:
    cols=['gross_amount','non_taxable_amount','pid','non_pid']
    CAs[cols] = CAs[cols].apply(pd.to_numeric, errors='ignore', axis=1)
except BaseException as ex:
    print(' ')


if not CAs.empty:
    print("Details of late CAs")
    CAs.drop_duplicates(subset=['ISIN', 'action_name','effective_date'], inplace=True)
    print(CAs)

conn.close()


print('')

print('details for confirmed takeovers/mergers/delistings')
print(final_df.sort_values(by=['Send_notice_today','EffectiveDate'],ascending=[False,True]))

if 'last_updated_at' in CAs.columns:
    CAs = CAs.drop('last_updated_at', axis=1)

CAs= CAs.replace(np.nan, '', regex=True)
email_table= CAs.to_html(index=False)
email_date_sub=currdate

#Sending email
if not CAs.empty:
    outlook = win32.gencache.EnsureDispatch('Outlook.Application')
    mail_item = outlook.CreateItem(0)
    mail_item.To = 'IndexTeamEU@cboe.com'
    mail_item.CC = 'IndexesDO@morningstar.com'
    body = "<br>Hi Team,</h1>Please note that today we have confirmed/updated the below mentioned corporate actions that have ex-date in next 3 business days: <br><br>   "+email_table+"   <br><br>Thanks"
    mail_item.HTMLBody = (body)
    mail_item.Subject = "Daily Report {}: List of CAs confirmed after DCAF publication".format(email_date_sub)
    mail_item.Display()
else:
    print(" ")

#produce JSON output
new_cas_status='Late CAs:CHECK' if not CAs.empty else 'Late CAs:clear'
send_notice_status='ECA notice:CHECK' if (final_df['Send_notice_today']=='Yes').any() else 'ECA notice:clear'

combined_status = f"Last bd: {date} | {new_cas_status} | {send_notice_status}"

status_data = {
    "check_name": "DCAF_CA_Check",
    "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    "checks": {
        "DCAF_summary": combined_status
    }
}
with open("status_dcaf_check.json", "w") as f:
    json.dump(status_data, f, indent=2)