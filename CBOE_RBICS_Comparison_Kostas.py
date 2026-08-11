import mariadb
from pathlib import Path
import sys
import os
import pandas as pd
from datetime import date

#To display all columns
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 1000)


path1 = os.path.expanduser("~/OneDrive - MORNINGSTAR INC/Indexes Global Operations-Daily Operations - Daily Operations/Index Operations/CBOE/Checks/")
path2 = os.path.expanduser("~/OneDrive - MORNINGSTAR INC/Daily Operations/Index Operations/CBOE/Checks/")

if os.path.exists(path1):
    path = path1
elif os.path.exists(path2):
    path = path2
else:
    raise FileNotFoundError("Neither path exists")

localpathlog = Path(path + 'RBICS validation')

#current date
currdate = date.today()


#Connecting to ticks DB
try:
    conn= mariadb.connect(
        user='ops_pending_ca_tracker',
        password='zmUFk8Xxcc8WrycCrMVe0plI84KkQAmqaOwoKHMn',
        host='mb-prd-db.us-east-1.mif0286.eas.morningstar.com',
        port=3306,
        database='ticks'
    )
except mariadb.Error as e:
    print(f"Error connecting to MariaDB Platform: {e}")
    sys.exit(1)

# Get Cursor
cur = conn.cursor()


#to get CBOE members and RBICS details from security characteristics table

sql_ticks="""Select distinct sa.security_id, sb.string_value as ISIN, sd.int_value as RBICS_code, se.string_value as RBICS_name from security_characteristic sa
left join security_characteristic sb on sa.security_id =sb.security_id  and sb.name ='ISIN' and CURRENT_DATE() between sb.valid_from and sb.valid_to
left join security_characteristic sd on sa.security_id =sd.security_id  and sd.name ='RBICSEconomyId' and CURRENT_DATE() between sd.valid_from and sd.valid_to
left join security_characteristic se on sa.security_id =se.security_id  and se.name ='RBICSEconomyName' and CURRENT_DATE() between se.valid_from and se.valid_to
where sa.security_id in (select distinct its.security_id
from index_timespan_security its
join index_timespan it on it.timespan_id = its.timespan_id
join ticks.index_timespan_security_aspect itsa on its.timespan_id =itsa.timespan_id and its.listing_id =itsa.listing_id and itsa.name ='Fraction'
where it.valid_from = CURRENT_DATE() and it.`type` in ('Open','Rebalance','Fixing')
and itsa.double_value <>0
and it.variation_id in (select DISTINCT iv.variation_id from index_variation iv 
join index_masterdata im on im.masterdata_id = iv.masterdata_id 
and im.client_id = 3))"""

cur.execute(sql_ticks)

result1=cur.fetchall()

ticks_data=pd.DataFrame(result1)

conn.close()

ticks_data.columns=["security_id","ISIN","RBICS_ECON_NUM_MSTAR","RBICS_ECON_NAME_MSTAR"]


ticks_data['ISIN'] = '"' + ticks_data['ISIN'].astype(str) + '"'
cboe_members=','.join(map(str,ticks_data['ISIN']))

#print(cboe_members)

#to connect to raw DB

try:
    conn1= mariadb.connect(
        user='kostas.dimitriou',
        password='1efp93nwb8ROfi3eXxtZRMG0_yYfDmKJjcjZmF3u',
        host='mb-prd-raw-db.us-west-2.mif0286.eas.morningstar.com',
        port=3306,
        database='raw_data'
    )
except mariadb.Error as e:
    print(f"Error connecting to MariaDB Platform: {e}")
    sys.exit(1)

# Get Cursor
cur1 = conn1.cursor()


#to get CBOE members and RBICS details from security characteristics table, using window function we ware getting latest details for each ISIN

sql_provider=""" with cte as (select distinct isin , rbics_econ_num , rbics_econ_name , `date`,ROW_NUMBER() over (PARTITION by isin order by `date` desc) num
from cboe_facts cf
where  `date`>= DATE_ADD(CURRENT_DATE(), INTERVAL -4 day) and
cf.isin in ("""+cboe_members+"""))
select isin , rbics_econ_num , rbics_econ_name , `date`
from cte
where num=1 """

cur1.execute(sql_provider)

result2=cur1.fetchall()

provider_data=pd.DataFrame(result2)

conn1.close()

provider_data.columns=["ISIN","RBICS_ECON_NUM_CBOE","RBICS_ECON_NAME_CBOE","CBOE_Date"]


#Removing double quotes added to ISINs
ticks_data['ISIN'] = ticks_data['ISIN'].str.strip('"')


#Merging ticks data with CBOE facts data, we are using left join here to get all the avlues from ticks data and matching values form cboe_facts
Comparison=pd.merge(ticks_data,provider_data, on='ISIN', how='left')

#to remove spaces (if any) to avoid any trailing spaces
#Comparison.map(lambda x: x.strip() if isinstance(x, str) else x), #if it gives error, replace map with applymap
# Comparison = Comparison.map(lambda x: x.strip() if isinstance(x, str) else x)
# compatible with old and new pandas versions
if hasattr(Comparison, 'applymap'):
    Comparison = Comparison.applymap(lambda x: x.strip() if isinstance(x, str) else x)
else:
    Comparison = Comparison.map(lambda x: x.strip() if isinstance(x, str) else x)


#compariosn of RBICS codes and names
Comparison['RBICS_ECON_NAME_Match'] = Comparison.apply(lambda Comparison: Comparison['RBICS_ECON_NAME_MSTAR'] == Comparison['RBICS_ECON_NAME_CBOE'], axis=1)

Comparison['RBICS_ECON_NUM_Match'] = Comparison.apply(lambda Comparison: Comparison['RBICS_ECON_NUM_MSTAR'] == Comparison['RBICS_ECON_NUM_CBOE'], axis=1)


#Re-ordering of columns in output file
Comparison = Comparison[['CBOE_Date','security_id', 'ISIN','RBICS_ECON_NUM_MSTAR','RBICS_ECON_NAME_MSTAR','RBICS_ECON_NUM_CBOE','RBICS_ECON_NAME_CBOE','RBICS_ECON_NAME_Match','RBICS_ECON_NUM_Match']]
Comparison=Comparison.drop('security_id',axis=1)

#exporting outfile to onedrive folder
outputfilename = localpathlog / 'RBICSfile-comparison-{0}.xlsx'.format(currdate)
#Comparison.to_excel(outputfilename, sheet_name='RBICS Details Check', index = False)

#printing various alerts
print('There is/are'+" " +str((~Comparison['RBICS_ECON_NAME_Match']).sum())+" "+' value/s not matching in RBICS_name')
print('There is/are'+" " +str((~Comparison['RBICS_ECON_NUM_Match']).sum())+" "+' value/s not matching in RBICS_num')

result_df=Comparison[~Comparison.RBICS_ECON_NAME_Match | ~Comparison.RBICS_ECON_NUM_Match]

print('\nISIN/s where there is a mismatch :\n', result_df)






