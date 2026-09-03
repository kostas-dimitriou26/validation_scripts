import mariadb
import sys
import os
import pandas as pd
import json
from datetime import date, datetime


OUTPUT_FILE = os.path.join("json_orange", "status_Day_shift_QC_checks.json")


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

sql_4 = """select distinct cs.config_id,case when cs.stock_exchange_id in (72,64,83,79,90,95,88,105,8,70) then 38
else 6 end as pqid,'REALTIME',cs.listing_id,case when cs.currency in ('CNH','GBP','KWD','ILS','ZAR')
then cs.currency else null end as override,1
from index_variation_currency_characteristic ivc
join calc_security cs on ivc.index_variation_id = cs.variation_id and cs.name = 'Open'
join index_timespan it on it.timespan_id=cs.timespan_id
join index_variation iv on ivc.index_variation_id = iv.variation_id
left join index_config_security ics on cs.config_id = ics.config_id and cs.listing_id = ics.listing_id and ics.price_type = 'REALTIME'
where it.valid_from=CURDATE()
and it.`type` ='Open'
and ivc.name= 'IsEnabledForRealtime' and ivc.bool_value = 1 and curdate() between ivc.valid_from and ivc.valid_to
and cs.the_date =CURDATE()
and ics.price_type is NULL;"""

cur.execute(sql_4)
result_4 = cur.fetchall()
check_4 = pd.DataFrame(
    result_4,
    columns=["config_id", "pqid", "price_type", "listing_id", "override", "flag"]
)

# --- CHECK 4.1: constituents in SOI index missing an RT symbol in provider query listing characteristic ---
sql_4_1 = """select distinct
	cs.security_id ,
	ics.listing_id,
	ics.provider_query_id,
	lc.string_value as ric,
	lc2.int_value as TradingStatus
from
	calc_security cs
join index_timespan it on
	it.timespan_id = cs.timespan_id
left join ticks.listing_characteristic lc on cs.listing_id =lc.listing_id and lc.name ='RIC' and lc.valid_to >=now()
left join ticks.listing_characteristic lc2 on cs.listing_id =lc2.listing_id and lc2.name ='TradingStatus' and lc2.valid_to >=now() and lc2.int_value <>0
join index_config_security ics on
	cs.config_id = ics.config_id
	and cs.listing_id = ics.listing_id
	and ics.price_type = 'REALTIME'
	and it.valid_from =CURDATE()
	and it.`type` = 'Open'
	where
	not exists (
	select 1 from provider_query_listing_characteristic pqlc
	where
		pqlc.listing_id = ics.listing_id
		and pqlc.provider_query_id = ics.provider_query_id
		and pqlc.valid_to >= now())
and ics.provider_query_id in (6,38);"""

cur.execute(sql_4_1)
result_4_1 = cur.fetchall()
check_4_1 = pd.DataFrame(
    result_4_1,
    columns=["security_id", "listing_id", "provider_query_id", "ric", "TradingStatus"]
)


# --- CHECK 5: minor currency securities missing override (CLOSE price type) ---
sql_5_close = """select * from index_config_security a
join listing_characteristic k on a.listing_id = k.listing_id and k.name = 'Currency'
join exchange_rate_constant z on k.currency = z.currency_from
where a.price_type = 'CLOSE'
and k.currency in ('CNY','GBX','ILA','KWF','ZAC')
and k.valid_to >= CURDATE()
and listing_currency_override is null;"""

cur.execute(sql_5_close)
result_5_close = cur.fetchall()
check_5_close = pd.DataFrame(result_5_close)


# --- CHECK 5: minor currency securities missing override (REALTIME price type) ---
sql_5_realtime = """select * from index_config_security a
join listing_characteristic k on a.listing_id = k.listing_id and k.name = 'Currency'
join exchange_rate_constant z on k.currency = z.currency_from
where a.price_type = 'REALTIME'
and k.currency in ('CNY','GBX','ILA','KWF','ZAC')
and k.valid_to >= CURDATE()
and listing_currency_override is null;"""

cur.execute(sql_5_realtime)
result_5_realtime = cur.fetchall()
check_5_realtime = pd.DataFrame(result_5_realtime)

# --- CHECK 6: check CBOE identifiers ---
sql_6 = """select  distinct f.security_id, a.listing_id,h.string_value as country, z.string_value as ISIN, g.string_value as name,  b.stock_exchange_id, c.currency,
       d.string_value as RIC, e.string_value as MorningStarId, k.string_value as BloombergSymbol,l.string_value as ExchangeTicker, n.string_value as 'Sedol',
       i.string_value as MIC, m.int_value as RBICS_Id, x.string_value as RBICS_Name, sh.string_value as 'MSName'
from index_timespan it
    join index_timespan_security its on its.timespan_id = it.timespan_id and it.config_id = its.config_id
    join index_variation iv on it.variation_id = iv.variation_id 
    join index_masterdata im on iv.masterdata_id = im.masterdata_id 
    join listing a on a.listing_id = its.listing_id
    join security f on a.security_id = f.security_id and f.security_type = 'Common Stock'
left join security_characteristic z on f.security_id = z.security_id and z.name = 'ISIN' and curdate() between z.valid_from and z.valid_to
left join security_characteristic g on f.security_id = g.security_id and g.name = 'Name' and curdate() between g.valid_from and g.valid_to
left join security_characteristic h on f.security_id = h.security_id and h.name = 'Country' and curdate() between h.valid_from and h.valid_to
left join security_characteristic sh on f.security_id = sh.security_id and sh.name = 'MSName' and curdate() between sh.valid_from and sh.valid_to
left join security_characteristic m on f.security_id = m.security_id and m.provider_query_id = 1 and m.name = 'RBICSEconomyId' and curdate() between m.valid_from and m.valid_to
left join security_characteristic x on f.security_id = x.security_id and x.provider_query_id = 1 and x.name = 'RBICSEconomyName' and curdate() between x.valid_from and x.valid_to
left join listing_characteristic b on a.listing_id = b.listing_id and b.name = 'StockExchangeId' and curdate() between b.valid_from and b.valid_to
left join listing_characteristic c on a.listing_id = c.listing_id and c.name = 'Currency' and curdate() between c.valid_from and c.valid_to
left join listing_characteristic d on a.listing_id = d.listing_id and d.name = 'RIC' and curdate() between d.valid_from and d.valid_to
left join listing_characteristic e on a.listing_id = e.listing_id and e.name = 'MorningstarSymbol' and curdate() between e.valid_from and e.valid_to
left join listing_characteristic k on a.listing_id = k.listing_id and k.provider_query_id = 22 and k.name = 'BloombergTicker' and curdate() between k.valid_from and k.valid_to
left join listing_characteristic l on a.listing_id = l.listing_id and l.name = 'ExchangeTicker' and curdate() between l.valid_from and l.valid_to
left join listing_characteristic n on a.listing_id = n.listing_id and n.name = 'Sedol' and curdate() between n.valid_from and n.valid_to
left join stock_exchange_characteristic i on b.stock_exchange_id = i.stock_exchange_id and i.name = 'MIC'  and curdate() between i.valid_from and i.valid_to
where curdate() = it.valid_from  and type in ('Open','Rebalance','Fixing')
and im.client_id = 3
and (h.string_value is null or z.string_value is null or g.string_value is null or  b.stock_exchange_id is null or c.currency is null or
       d.string_value is null or  k.string_value is null or l.string_value is null or n.string_value is null or n.string_value = ''  or  m.int_value is null or x.string_value is null or
       i.string_value is null or sh.string_value is null
       )
order by f.security_id;"""

cur.execute(sql_6)
result_6 = cur.fetchall()
colnames_6 = [desc[0] for desc in cur.description]
check_6 = pd.DataFrame(result_6, columns=colnames_6)


# --- CHECK 6.1: check Morningstar identifiers ---
sql_6_1 = """select distinct  c2.name, rfs.listing_id, rfs.security_id, s.issuer_id, ic.string_value as 'MSCompanyId' , ic1.string_value as 'MSCountryOfClassification',
ic2.string_value as 'MSCountryOfDomicile', ic3.string_value as 'MSMarketSegment', ic4.string_value as 'MSRegion',
sc.string_value as 'MSName', lc.string_value as 'MSShareClassId', x.string_value as 'GICSLevel4Code',
b.stock_exchange_id, c.currency, d.string_value as 'RIC', e.string_value as 'MorningstarSymbol', k.string_value as 'BloombergTicker',
l.string_value as 'ExchangeTicker', n.string_value as 'Sedol'
from rebalance r 
join rebalance_fixing_security rfs on r.rebalance_id = rfs.rebalance_id 
join `security` s on s.security_id = rfs.security_id 
join index_variation ivc on r.variation_id = ivc.variation_id 
join index_masterdata im on ivc.masterdata_id = im.masterdata_id 
join client c2 on im.client_id = c2.client_id 
left join issuer_characteristic ic on s.issuer_id = ic.issuer_id and ic.name = 'MSCompanyId' and CURDATE() between ic.valid_from and ic.valid_to
left join issuer_characteristic ic1 on s.issuer_id = ic1.issuer_id and ic1.name = 'MSCountryOfClassification' and CURDATE() between ic1.valid_from and ic1.valid_to 
left join issuer_characteristic ic2 on s.issuer_id = ic2.issuer_id and ic2.name = 'MSCountryOfDomicile' and CURDATE() between ic2.valid_from and ic2.valid_to 
left join issuer_characteristic ic3 on s.issuer_id = ic3.issuer_id and ic3.name = 'MSMarketSegment' and CURDATE() between ic3.valid_from and ic3.valid_to 
left join issuer_characteristic ic4 on s.issuer_id = ic4.issuer_id and ic4.name = 'MSRegion' and CURDATE() between ic4.valid_from and ic4.valid_to 
left join security_characteristic sc on rfs.security_id = sc.security_id and sc.name = 'MSName' and CURDATE() between sc.valid_from and sc.valid_to
left join security_characteristic sc1 on rfs.security_id = sc1.security_id and sc1.name = 'MSName' and CURDATE() between sc1.valid_from and sc1.valid_to
left join listing_characteristic lc on rfs.listing_id = lc.listing_id and lc.name = 'MSShareClassId' and CURDATE() between lc.valid_from and lc.valid_to
left join security_characteristic x on rfs.security_id = x.security_id and x.name = 'GICSLevel4Code' and curdate() between x.valid_from and x.valid_to
left join listing_characteristic b on rfs.listing_id = b.listing_id and b.name = 'StockExchangeId' and curdate() between b.valid_from and b.valid_to
left join listing_characteristic c on rfs.listing_id = c.listing_id and c.name = 'Currency' and curdate() between c.valid_from and c.valid_to
left join listing_characteristic d on rfs.listing_id = d.listing_id and d.name = 'RIC' and curdate() between d.valid_from and d.valid_to
left join listing_characteristic e on rfs.listing_id = e.listing_id and e.name = 'MorningstarSymbol' and curdate() between e.valid_from and e.valid_to
left join listing_characteristic k on rfs.listing_id = k.listing_id and k.name = 'BloombergTicker' and k.provider_query_id = 22  and curdate() between k.valid_from and k.valid_to
left join listing_characteristic l on rfs.listing_id = l.listing_id and l.name = 'ExchangeTicker' and curdate() between l.valid_from and l.valid_to
left join listing_characteristic n on rfs.listing_id = n.listing_id and n.name = 'Sedol' and curdate() between n.valid_from and n.valid_to
where r.fixing_date >= '2026-07-01' and r.variation_id not in  (1441, 2152,2039) and b.stock_exchange_id not in  (180,174) and im.client_id in ('9','12','20','32','31')
and (ic.string_value is null or ic1.string_value is null or ic2.string_value is null or ic3.string_value is null or ic4.string_value is null or
sc.string_value is null or sc1.string_value is null or lc.string_value is null or x.string_value is null or b.stock_exchange_id is null or c.currency is null or 
d.string_value is null  or k.string_value is null or l.string_value is null or n.string_value is null)
and s.security_type = 'Common Stock';"""

cur.execute(sql_6_1)
result_6_1 = cur.fetchall()
colnames_6_1 = [desc[0] for desc in cur.description]
check_6_1 = pd.DataFrame(result_6_1, columns=colnames_6_1)

# --- CHECK 12: RIC in listing_characteristic vs provider_query_listing_characteristic mismatch ---
sql_12 = """select * from (select distinct its.security_id,
sc.string_value as ISIN,
its.listing_id,
lc.string_value as ric,
lc.provider_query_id,
lc.valid_from as ric_valid_from,
pqlc.string_value as pqid_56,
pqlc.valid_from as pqid_56_valid_from,
case when lc.string_value=pqlc.string_value then 'Ok' else 'Check' end as matching
from index_timespan it
join index_timespan_security its on it.timespan_id = its.timespan_id
join index_timespan_security_aspect itsa on its.timespan_id =itsa.timespan_id and its.listing_id =itsa.listing_id and itsa.name ='Fraction'
join security_characteristic sc on its.security_id = sc.security_id and sc.name = 'ISIN' and CURRENT_DATE() between sc.valid_from and sc.valid_to
join listing_characteristic lc on its.listing_id =lc.listing_id and lc.name='RIC' and CURDATE() between lc.valid_from and lc.valid_to
join provider_query_listing_characteristic pqlc on its.listing_id =pqlc.listing_id and pqlc.provider_query_id =56 and CURDATE() between pqlc.valid_from and pqlc.valid_to
join ticks.security s on its.security_id=s.security_id
where it.valid_from = CURRENT_DATE()  and it.`type` = 'Open'
and abs(itsa.double_value) <> 0
and s.security_type='Common Stock'
order by matching ) as ric
where ric.matching='Check';"""

cur.execute(sql_12)
result_12 = cur.fetchall()
colnames_12 = [desc[0] for desc in cur.description]
check_12 = pd.DataFrame(result_12, columns=colnames_12)


# --- CHECK 14: names for pq id 42 and pq id 1 matching ---
sql_14 = """select * from (select distinct its.security_id,
sc.string_value as ISIN, 
sc2.string_value as name_pqid_1,
sc3.string_value as name_pqid_42,
case when lower(sc2.string_value)=lower(sc3.string_value) then 'Ok' else 'Check' end as matching
from index_timespan it
join index_timespan_security its on it.timespan_id = its.timespan_id
join calc_security cs on its.timespan_id =cs.timespan_id and it.`type` =cs.name and it.valid_from =cs.the_date 
join index_timespan_security_aspect itsa on its.timespan_id =itsa.timespan_id and its.listing_id =itsa.listing_id and itsa.name ='Fraction'
join security_characteristic sc on its.security_id = sc.security_id and sc.name = 'ISIN' and CURRENT_DATE() between sc.valid_from and sc.valid_to
join security_characteristic sc2 on its.security_id = sc2.security_id and sc2.name = 'Name' and CURRENT_DATE() between sc2.valid_from and sc2.valid_to and sc2.provider_query_id=1
join security_characteristic sc3 on its.security_id = sc3.security_id and sc3.name = 'MSName' and CURRENT_DATE() between sc3.valid_from and sc3.valid_to and sc3.provider_query_id=42
join `security` s on its.security_id=s.security_id
where it.valid_from = CURRENT_DATE()  and it.`type` in ('Open,''Rebalance','Fixing')
and abs(itsa.double_value) <> 0
and s.security_type='Common Stock'
and cs.client_id =3
order by matching ) as name
where name.matching='Check';"""

cur.execute(sql_14)
result_14 = cur.fetchall()
colnames_14 = [desc[0] for desc in cur.description]
check_14 = pd.DataFrame(result_14, columns=colnames_14)


# --- CHECK 15: duplicate provider_query_id 22 entries ---
sql_15 = """select string_value ,count(*)
from provider_query_listing_characteristic pqlc
where valid_to>=now()
and pqlc.provider_query_id in (22)
group by pqlc.string_value
having count(*)>1;"""

cur.execute(sql_15)
result_15 = cur.fetchall()
colnames_15 = [desc[0] for desc in cur.description]
check_15 = pd.DataFrame(result_15, columns=colnames_15)

status_4 = 'CHECK' if not check_4.empty else 'clear'
status_4_1 = 'CHECK' if not check_4_1.empty else 'clear'
status_5_close = 'CHECK' if not check_5_close.empty else 'clear'
status_5_realtime = 'CHECK' if not check_5_realtime.empty else 'clear'
status_6 = 'CHECK' if not check_6.empty else 'clear'
status_6_1 = 'CHECK' if not check_6_1.empty else 'clear'
status_12 = 'CHECK' if not check_12.empty else 'clear'
status_14 = 'CHECK' if not check_14.empty else 'clear'
status_15 = 'CHECK' if not check_15.empty else 'clear'

combined_status = (
    f"4: {status_4} | "
    f"4.1: {status_4_1} | "
    f"5 CLOSE: {status_5_close} | "
    f"5 REALTIME: {status_5_realtime} | "
    f"6: {status_6} | "
    f"6.1: {status_6_1} | "
    f"12: {status_12} | "
    f"14: {status_14} | "
    f"15: {status_15}"
)

status_data = {
    "check_name": "Day_shift_QC_checks",
    "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    "checks": {
        "Day_shift_QC_checks": combined_status,
    }
}
with open(OUTPUT_FILE, "w") as f:
    json.dump(status_data, f, indent=2)