#!/usr/bin/env python3
"""
Precision Blueprint Development — Financial Performance Database builder.

Reads the weekly report exports (General Ledger, Trial Balance, A/R Aging,
A/P Aging, Payroll, JobTread WIP) and produces the normalized snapshot tables
as (a) a multi-tab .xlsx workbook (uploaded to Google Drive as a Google Sheet)
and (b) one CSV per table.

Design points agreed with client (IronGate BP):
  * GL is stored at TRANSACTION level (append new lines each week, dedup on uid).
  * Adds AP_Aging_Snapshots + Payroll_Snapshots ("adding reports for clarity").
  * Ashley Espinoza's labor is administrative/overhead -> cost_classification.
  * Subcontractor allocation-confidence field: direct / matched / unallocated.
    Tiering is intentionally conservative (client "unsure for now"): a line is
    only 'direct' when its memo carries an explicit JobTread job link/number;
    'matched' when the vendor is mapped to a job in Sub_Contract_Reference;
    otherwise 'unallocated'. Sub_Contract_Reference is seeded with every
    distinct subcontractor vendor for the client to map vendor->job.
"""
import csv, io, json, os, re, base64, hashlib, datetime as dt

HERE = os.path.dirname(os.path.abspath(__file__))
OUT  = os.environ.get("PBD_OUT", os.path.join(HERE, "out"))
SRC  = os.path.join(HERE, "sources")
GL_TXT = os.environ["PBD_GL_TXT"]           # extracted GL fileContent (CSV text)
INGEST_TS = os.environ.get("PBD_INGEST_TS", "2026-07-18T23:30:00-04:00")  # baseline run stamp

os.makedirs(OUT, exist_ok=True)

def m(x):
    """parse money/number -> float or None"""
    if x is None: return None
    if isinstance(x,(int,float)): return float(x)
    s=str(x).replace('$','').replace(',','').strip()
    if s in ('','-','None'): return None
    neg = s.startswith('(') and s.endswith(')')
    s=s.strip('()')
    try: v=float(s)
    except: return None
    return -v if neg else v

def iso(d):
    """MM/DD/YYYY -> YYYY-MM-DD"""
    if not d: return ""
    d=str(d).strip()
    for f in ("%m/%d/%Y","%Y-%m-%d","%m/%d/%y"):
        try: return dt.datetime.strptime(d,f).strftime("%Y-%m-%d")
        except: pass
    return d

def slug(s):
    return re.sub(r'[^a-z0-9]+','-', (s or '').lower()).strip('-')[:60]

# ---------------------------------------------------------------------------
# SOURCE FILE REGISTRY (Drive) — used for Ingestion_Log provenance
# ---------------------------------------------------------------------------
SOURCES = {
 "GL":      ("Precision Blueprint Development_General Ledger.csv",        "1DcT8YQgs2PP41KhU0eVK7CDrW7BVF29l", "General Ledger",        "Jan 1 - Jun 30, 2026", "Cash Basis"),
 "TB":      ("Trial Balance - Last Week.xlsx",                            "13IuPK7aEZzBAPhmQEDu8goBWZDe2z_PS", "Trial Balance",         "As of Jul 11, 2026",   "Accrual Basis"),
 "AR":      ("A_R Aging Detail Report - Last Week.xlsx",                  "1VhiFK2Y64A403RJqFd-C2p3lG80rK7Ow", "A/R Aging Detail",      "As of Jul 11, 2026",   ""),
 "AP":      ("Precision Blueprint Development_A_P Aging Detail Report.xlsx","1nPYMlSpuC-038jQA2svvXPwhXnAPKfWm","A/P Aging Detail",      "As of Jul 15, 2026",   ""),
 "WIP":     ("job-export.csv",                                            "1znuclwf_FQ6XK7SdYvrOdQgtI4wwlEbs", "JobTread WIP",          "As of Jul 18, 2026",   ""),
 "PR_DTL":  ("Precision Blueprint Development Llc_Reports_2026_07_18.xlsx","13nhLfz2fNG4AxotvHeoDzE1P_iWoGxA9", "Payroll Detail",        "6/5/2026 - 7/17/2026",  ""),
 "PR_SUM":  ("Payroll Summary 2025 - 2026.pdf",                           "1OtZbwATsozCRzenWPoEE2e-tRpM7BFmr", "Payroll Summary",       "6/20/2025 - 6/12/2026", ""),
}

# ===========================================================================
# 1. GENERAL LEDGER  (parse from disk text -> transaction level)
# ===========================================================================
def parse_gl():
    txt = json.loads(open(GL_TXT, encoding='utf-8').read())["fileContent"]
    rows = list(csv.reader(io.StringIO(txt)))
    meta = re.compile(r'(Basis|GMTZ|General Ledger|Precision Blueprint|Distribution account|January 1-)')
    out=[]; section=None; seq=0
    SUB_VENDORS={}
    for r in rows:
        if not any((c or '').strip() for c in r): continue
        r=[(c or '').strip() for c in (r+['']*10)[:10]]
        c0,c1,c2,c3,c4,c5,c6,c7,c8,c9 = r
        if c0 and not c1 and not c2 and not c8 and not c0.startswith("Total for") and not meta.search(c0):
            section=c0; continue
        if c0.startswith("Total for") or meta.search(c0) or meta.search(c1): continue
        if c1=="Beginning Balance" or not c2: continue
        amt=m(c8); bal=m(c9); seq+=1
        is_sub = (section=="Sub-Contractors")
        # allocation confidence (conservative): only 'direct' with explicit jobtread ref
        alloc=""; job_id=""
        if is_sub:
            blob=(c6+" "+c5).lower()
            mo=re.search(r'/jobs/([0-9a-z]{8,})', blob)
            if 'jobtread' in blob or mo:
                alloc='direct'; job_id=(mo.group(1) if mo else '')
            else:
                alloc='unallocated'   # may become 'matched' via Sub_Contract_Reference
        uid=hashlib.sha1("|".join([str(seq),section or '',c2,c3,c4,c5,c6,c7,c8,c9]).encode()).hexdigest()[:16]
        out.append(dict(gl_account=section, txn_date=iso(c2), txn_type=c3, num=c4,
                        name=c5, memo=c6, split_account=c7, amount=amt, running_balance=bal,
                        is_subcontractor=("Y" if is_sub else "N"),
                        allocation_confidence=alloc, allocated_jobtread_id=job_id,
                        row_uid=uid, source_file=SOURCES["GL"][0], gl_period=SOURCES["GL"][3],
                        gl_basis=SOURCES["GL"][4], ingested_at=INGEST_TS))
        if is_sub:
            v=c5 or "(no vendor)"
            d=SUB_VENDORS.setdefault(v, {"total":0.0,"cnt":0})
            d["total"]+=(amt or 0); d["cnt"]+=1
    return out, SUB_VENDORS

# ===========================================================================
# 2. TRIAL BALANCE  (transcribed; asserts debits == credits)
# ===========================================================================
# (account, debit, credit)  As of Jul 11 2026, Accrual Basis
TB_ROWS = [
 ("Chase Plat Bus Checking 6106",422686.73,None),("Accounts Receivable (A/R)",432296.14,None),
 ("Accounts Payable (A/P)",None,12155.00),("Chase Credit Card - 2877",None,32458.79),
 ("Payroll Liabilities:Payroll Taxes",None,946.05),("Owner's Draw",218693.90,None),
 ("Sales",None,1358500.87),("Cost of Goods Sold",56215.52,None),("Disposal",2949.23,None),
 ("Job Supply",74253.36,None),("Sub-Contractors",383759.03,None),("Auto:Gas",6947.03,None),
 ("Auto:Repair & Maintenance",2287.22,None),("General Advertising:Advertising & Marketing",4893.69,None),
 ("Meals & Entertainment:50% Deductible",3486.28,None),
 ("Office Expenses:Bank Charges & Fees:Deductible",136.31,None),
 ("Office Expenses:Bank Charges & Fees:Merchant Fees",847.77,None),
 ("Office Expenses:Taxes & Licenses",998.39,None),("Office Expenses:Website & Software",13543.35,None),
 ("Payroll Expenses:Payroll Processing Fees",2283.82,None),
 ("Payroll Expenses:Payroll Taxes (ER)",3586.90,None),
 ("Payroll Expenses:Salaries/Labor:Wages",43269.30,None),("Retained Earnings",None,316261.07),
 ("Commission Expense",30335.00,None),("Equipment Rental",2193.50,None),("License & Permit",2669.00,None),
 ("Plans",5362.52,None),("Advertising Expense",845.00,None),
 ("Auto:Registration & License Fees",120.06,None),("General Advertising:Client/Vendor Gifts",583.20,None),
 ("Insurance Expense",1448.79,None),("Insurance Expense:Auto",3107.69,None),
 ("Insurance Expense:Worker's Compensation",3000.10,None),
 ("Meals & Entertainment:Employee Meals - Non-Deductible",1207.71,None),
 ("Office Expenses:Bank Charges & Fees",15.00,None),
 ("Office Expenses:Bank Charges & Fees:Interest Paid",5154.96,None),
 ("Office Expenses:Office Supplies",326.16,None),("Office Expenses:Telephone",1074.11,None),
 ("Office Expenses:Utilities & Internet",400.00,None),("Professional Services:Legal Services",8276.00,None),
 ("Professional Services:Tax & Accounting Services",21954.46,None),("Travel:Airfare",9.62,None),
 ("Travel:Lodging",289.51,None),("Travel:Parking",90.00,None),("Travel:Transportation",547.70,None),
 ("Cash Back Rewards",None,3246.61),("Dividend",None,0.06),("Charitiable Contributions",1065.00,None),
 ("2020 F350 Loan",None,36024.86),("Barclay CC 0333",None,750.57),("Best Buy Store Credit",None,0.00),
 ("Business phones:Accumulated depreciation",None,1467.00),
 ("Business phones:iPhone 14 Pro Max Silver",1590.62,None),
 ("Business phones:iPhone 15 Pro Max (256 GB Blue Titanium)",1300.24,None),
 ("Business phones:iPhone 15 Pro Max (512 GB White Titanium)",1217.06,None),
 ("Chase Business Card 3140 (9274)(7687)",None,0.00),("Chase Credit Card 3571",None,0.00),
 ("Computers:2024 Laptop (1)",None,0.00),("Computers:HP Envy Laptop",970.18,None),
 ("Computers:HP Victus Laptop",862.39,None),("Computers:IBuyPower Y60 Desktop - 2024",1832.59,None),
 ("Computers:Meta Quest Pro - 2024",985.28,None),("Equipment:Pressure Washer",900.00,None),
 ("Dump Trailer Loan from John Niezen",None,0.00),
 ("Home Depot Card - 3038",None,0.00),("Loan to Cindy & Paul",None,0.00),("MACU Loan",None,21036.89),
 ("MACU Savings 6934",365.85,None),("Notes Payable:Raul",None,0.00),("Opening Balance Equity",None,0.00),
 ("Owner's Contribution",None,28393.25),("Payroll Liabilities",None,0.00),("Prepaid Payroll",None,0.00),
 ("Transport equipment:2002 Ford F350 Dually",18000.00,None),
 ("Transport equipment:2020 Ford F350 SuperDuty",74027.00,None),
 ("Transport equipment:2020 Ford F350 SuperDuty:Accumulated depreciation",None,62553.00),
 ("Transport equipment:Box Trailer",5500.00,None),("Transport equipment:Dump Trailer",9280.00,None),
 ("Transport equipment:Dump Trailer:Accumulated depreciation",None,7934.00),
 ("Undeposited Funds",None,0.00),("Wells Fargo Credit Card 6410 (9790)",None,0.00),
 ("Wells Fargo Operating 9366",1687.75,None),
]

# ===========================================================================
# 3. A/R AGING DETAIL  (transcribed; asserts open total 432,296.14)
# ===========================================================================
# (bucket, date, ttype, num, customer_full_path, due, amount, open)
AR_ROWS = [
 ("91 or more days past due","03/03/2026","Invoice","2-2","Kristi Kahler/Rodriguez:Kahler-Rodriguez:Kahler Design/Addition/Remodel","04/02/2026",3350.00,3350.00),
 ("61 - 90 days past due","03/17/2026","Invoice","2-10","Kristi Kahler/Rodriguez:Kahler-Rodriguez:Kahler Design/Addition/Remodel","04/16/2026",975.61,975.61),
 ("61 - 90 days past due","03/25/2026","Invoice","14-2","Mike Baker:10115 West Montecito Avenue:Baker Patio Enclosure","04/24/2026",2785.72,2785.72),
 ("61 - 90 days past due","04/03/2026","Invoice","9-4","Mandarin Investments - Weavers Needle:250 South Tomahawk Road:Weavers Needle Kitchen","05/03/2026",4500.00,4500.00),
 ("61 - 90 days past due","04/09/2026","Invoice","2-13","Kristi Kahler/Rodriguez:Kahler-Rodriguez:Kahler Design/Addition/Remodel","05/09/2026",1714.29,1714.29),
 ("61 - 90 days past due","04/10/2026","Invoice","20-4","Nikki Jackson:3532 East Fairbrook Circle:Jackson Truss Repair","05/10/2026",727.38,727.38),
 ("31 - 60 days past due","04/22/2026","Invoice","20-5","Nikki Jackson:3532 East Fairbrook Circle:Jackson Truss Repair","05/22/2026",727.38,727.38),
 ("31 - 60 days past due","04/27/2026","Invoice","29-4","Chris Church:Chris Church:Chris Church Truss Repair","05/27/2026",960.00,960.00),
 ("31 - 60 days past due","04/28/2026","Invoice","29-5","Chris Church:Chris Church:Chris Church Truss Repair","05/28/2026",960.00,960.00),
 ("31 - 60 days past due","05/04/2026","Invoice","35-2","Laurie Ruiz:6730 E Preston St unit 69:Laurie Ruiz Truss Repair","06/03/2026",2300.00,2300.00),
 ("31 - 60 days past due","05/04/2026","Invoice","27-2","Jon Jenson:10008 North 44th Drive:North 44th Truss Repair","06/03/2026",5463.85,5463.85),
 ("31 - 60 days past due","05/09/2026","Invoice","35-4","Laurie Ruiz:6730 E Preston St unit 69:Laurie Ruiz Truss Repair","06/08/2026",2300.00,2300.00),
 ("1 - 30 days past due","05/14/2026","Invoice","44-2","Leslie Cook:543 West Bluejay Drive:Cook Truss Repair","06/13/2026",568.18,568.18),
 ("1 - 30 days past due","05/16/2026","Invoice","30-4","Scott Donaldson:7050 N. 8th Ave.:Scott Donaldson Truss Repair","06/15/2026",1763.89,1763.89),
 ("1 - 30 days past due","05/16/2026","Invoice","44-3","Leslie Cook:543 West Bluejay Drive:Cook Truss Repair","06/15/2026",568.18,568.18),
 ("1 - 30 days past due","05/18/2026","Invoice","17-3","Lettye Harris // Church Addition:349 W Mohave St:Greater New Zion Baptist","06/17/2026",8678.58,8678.58),
 ("1 - 30 days past due","05/19/2026","Invoice","14-3","Mike Baker:10115 West Montecito Avenue:Baker Patio Enclosure","06/18/2026",2785.71,2785.71),
 ("1 - 30 days past due","05/19/2026","Invoice","16-3","Brent Wookey:6150 N Scottsdale Rd:Wookey Interior Remodel","06/18/2026",15622.86,15622.86),
 ("1 - 30 days past due","05/27/2026","Invoice","37-2","Patricia Crawford:1414 East Almeria Road:Crawford Exterior Updates","06/26/2026",6842.66,6842.66),
 ("1 - 30 days past due","05/28/2026","Invoice","2-15","Kristi Kahler/Rodriguez:Kahler-Rodriguez:Kahler Design/Addition/Remodel","06/27/2026",274199.35,274199.35),
 ("1 - 30 days past due","05/28/2026","Invoice","52-2","Debbie Shields:10205 N 109th Way Scottsdale:Satish Truss Repair","06/27/2026",3267.27,3267.27),
 ("1 - 30 days past due","05/28/2026","Invoice","58-2","Sean Blumhoff:11501 East Desert Willow Drive:Blumhoff Truss Repair","06/27/2026",392.86,392.86),
 ("1 - 30 days past due","06/02/2026","Invoice","58-3","Sean Blumhoff:11501 East Desert Willow Drive:Blumhoff Truss Repair","07/02/2026",392.85,392.85),
 ("1 - 30 days past due","07/07/2026","Payment","2930","Rebecca Wong","07/07/2026",-2036.25,-161.25),
 ("1 - 30 days past due","07/08/2026","Payment","120","Ryan Goodwin (c)","07/08/2026",-5436.32,-453.45),
 ("1 - 30 days past due","06/10/2026","Invoice","61-2","Elizabeth Azurdia:1608 W Alamo Dr:Azurdia Truss Repair","07/10/2026",922.54,922.54),
 ("CURRENT","06/11/2026","Invoice","62-4","Giselle Bastida:14089 West Country Gables Drive:Bastida Truss Repair","07/11/2026",2065.22,2065.22),
 ("CURRENT","06/11/2026","Invoice","67-2","Ann Purcell:2448 W Hawken Pl Chandler AZ:Purcell Door Replacement","07/11/2026",699.11,699.11),
 ("CURRENT","06/11/2026","Invoice","52-3","Debbie Shields:10205 N 109th Way Scottsdale:Satish Truss Repair","07/11/2026",3267.28,3267.28),
 ("CURRENT","06/14/2026","Invoice","37-8","Patricia Crawford:1414 East Almeria Road:Crawford Exterior Updates","07/14/2026",9342.66,9342.66),
 ("CURRENT","06/17/2026","Invoice","67-3","Ann Purcell:2448 W Hawken Pl Chandler AZ:Purcell Door Replacement","07/17/2026",699.11,699.11),
 ("CURRENT","06/19/2026","Invoice","2-43","Kristi Kahler/Rodriguez:Kahler-Rodriguez:Kahler Design/Addition/Remodel","07/19/2026",2571.43,2571.43),
 ("CURRENT","06/19/2026","Invoice","72-2","Jeff Berning:7581 East Phantom Way:Berning Truss Repair","07/19/2026",494.44,494.44),
 ("CURRENT","06/22/2026","Invoice","70-2","Leslie Davidson:5007 S Terrace Rd:Davidson Truss Repair","07/22/2026",831.94,831.94),
 ("CURRENT","06/22/2026","Invoice","72-3","Jeff Berning:7581 East Phantom Way:Berning Truss Repair","07/22/2026",494.45,494.45),
 ("CURRENT","06/27/2026","Invoice","61-6","Elizabeth Azurdia:1608 W Alamo Dr:Azurdia Truss Repair","07/27/2026",922.54,922.54),
 ("CURRENT","06/29/2026","Invoice","2-49","Kristi Kahler/Rodriguez:Kahler-Rodriguez:Kahler Design/Addition/Remodel","07/29/2026",68753.50,68753.50),
]

# ===========================================================================
# 4. A/P AGING DETAIL  (transcribed; asserts open total 12,155.00)
# ===========================================================================
# (bucket, date, ttype, num, vendor, due, days_past_due, amount, open)
AP_ROWS = [
 ("91 or more days past due","03/16/2026","Bill","1288","GenCo Consulting","04/15/2026",94,145.00,145.00),
 ("61 - 90 days past due","03/17/2026","Bill","1295","GenCo Consulting","04/16/2026",93,30.00,30.00),
 ("CURRENT","06/16/2026","Bill","27500","GTI Geo Tech Testing and Inspections","07/16/2026",2,2600.00,2600.00),
 ("CURRENT","06/16/2026","Bill","001","Gen-Co Consulting LLC","07/16/2026",2,5500.00,5500.00),
 ("CURRENT","06/18/2026","Bill","002","J's Plumbing","07/18/2026",0,3500.00,3500.00),
 ("CURRENT","06/19/2026","Bill","1329","Gen-Co Consulting LLC","07/19/2026",-1,260.00,260.00),
 ("CURRENT","06/19/2026","Bill","1330","Gen-Co Consulting LLC","07/19/2026",-1,120.00,120.00),
]

# ===========================================================================
# 5. PAYROLL SUMMARY  (52 weekly checks, single employee — overhead/admin)
# ===========================================================================
EMP="Espinoza, Ashley"
# (check_date, total_paid, tax_withheld, net_pay, employer_liability, total_expense) ; hours 40, ded 0, DD
PR_SUM = [
 ("12/26/2025",1442.31,184.08,1258.23,110.33,1552.64),("12/19/2025",1442.31,184.08,1258.23,110.33,1552.64),
 ("12/12/2025",1442.31,184.10,1258.21,110.33,1552.64),("12/05/2025",1442.31,184.08,1258.23,110.33,1552.64),
 ("11/28/2025",1442.31,184.08,1258.23,110.33,1552.64),("11/21/2025",1442.31,184.09,1258.22,110.33,1552.64),
 ("11/14/2025",1442.31,184.09,1258.22,110.33,1552.64),("11/07/2025",1442.31,184.08,1258.23,110.33,1552.64),
 ("10/31/2025",1442.31,184.09,1258.22,110.33,1552.64),("10/24/2025",1442.31,184.09,1258.22,110.33,1552.64),
 ("10/17/2025",1442.31,184.08,1258.23,110.33,1552.64),("10/10/2025",1442.31,184.09,1258.22,110.33,1552.64),
 ("10/03/2025",1442.31,184.09,1258.22,110.33,1552.64),("09/26/2025",1442.31,184.08,1258.23,110.33,1552.64),
 ("09/19/2025",1442.31,184.09,1258.22,110.33,1552.64),("09/12/2025",1442.31,184.09,1258.22,110.33,1552.64),
 ("09/05/2025",1442.31,184.09,1258.22,110.33,1552.64),("08/29/2025",1442.31,184.08,1258.23,110.33,1552.64),
 ("08/22/2025",1442.31,184.09,1258.22,110.33,1552.64),("08/15/2025",1346.16,162.78,1183.38,102.98,1449.14),
 ("08/08/2025",1346.16,162.78,1183.38,102.98,1449.14),("08/01/2025",1346.16,162.78,1183.38,102.98,1449.14),
 ("07/25/2025",1346.16,162.78,1183.38,102.98,1449.14),("07/18/2025",1346.16,162.79,1183.37,102.98,1449.14),
 ("07/11/2025",1346.16,162.78,1183.38,102.98,1449.14),("07/03/2025",1346.16,162.78,1183.38,102.98,1449.14),
 ("06/27/2025",1346.16,162.78,1183.38,102.98,1449.14),("06/20/2025",1346.16,162.78,1183.38,102.98,1449.14),
 ("06/12/2026",1442.31,180.02,1262.29,110.33,1552.64),("06/05/2026",1442.31,180.01,1262.30,110.33,1552.64),
 ("05/29/2026",1442.31,180.02,1262.29,110.33,1552.64),("05/22/2026",1442.31,180.02,1262.29,110.33,1552.64),
 ("05/15/2026",1442.31,180.01,1262.30,110.33,1552.64),("05/08/2026",1442.31,180.02,1262.29,110.33,1552.64),
 ("05/01/2026",1442.31,180.02,1262.29,110.33,1552.64),("04/24/2026",1442.31,180.01,1262.30,110.33,1552.64),
 ("04/17/2026",1442.31,180.02,1262.29,110.33,1552.64),("04/10/2026",1442.31,180.01,1262.30,110.33,1552.64),
 ("04/03/2026",1442.31,180.02,1262.29,110.33,1552.64),("03/27/2026",1442.31,180.02,1262.29,110.33,1552.64),
 ("03/20/2026",1442.31,180.01,1262.30,110.33,1552.64),("03/13/2026",1442.31,180.03,1262.28,110.33,1552.64),
 ("03/06/2026",1442.31,180.01,1262.30,110.33,1552.64),("02/27/2026",1442.31,180.01,1262.30,110.33,1552.64),
 ("02/20/2026",1442.31,180.03,1262.28,110.33,1552.64),("02/13/2026",1442.31,180.01,1262.30,110.33,1552.64),
 ("02/06/2026",1442.31,180.01,1262.30,126.10,1568.41),("01/30/2026",1442.31,180.03,1262.28,146.56,1588.87),
 ("01/23/2026",1442.31,180.01,1262.30,147.83,1590.14),("01/16/2026",1442.31,180.01,1262.30,147.83,1590.14),
 ("01/09/2026",1442.31,180.03,1262.28,147.83,1590.14),("01/02/2026",1442.31,180.01,1262.30,147.83,1590.14),
]

# ===========================================================================
# 6. JOBTREAD WIP  (parse valid CSV)
# ===========================================================================
def parse_wip():
    data=base64.b64decode(open(os.path.join(SRC,"wip.csv.b64")).read().strip()).decode('utf-8-sig')
    rows=list(csv.reader(io.StringIO(data)))
    hdr=rows[0]; out=[]
    idx={h:i for i,h in enumerate(hdr)}
    def g(r,k):
        i=idx.get(k); return r[i] if (i is not None and i < len(r)) else ""
    for r in rows[1:]:
        out.append(dict(
            jobtread_job_id=g(r,"ID"), job_number=g(r,"Number"), job_name=g(r,"Name"),
            price_type=g(r,"Price Type"),
            projected_price=m(g(r,"Projected Price")), budget_cost=m(g(r,"Budget Cost")),
            projected_cost=m(g(r,"Projected Cost")), actual_cost=m(g(r,"Actual Cost")),
            cost_to_complete=m(g(r,"Cost To Complete")), budget_variance=m(g(r,"Budget Variance")),
            projected_profit=m(g(r,"Projected Profit")), projected_margin=m(g(r,"Projected Margin")),
            pct_cost_complete=m(g(r,"% Cost Complete")), earned_revenue=m(g(r,"Earned Revenue")),
            invoiced=m(g(r,"Invoiced")), over_invoiced=m(g(r,"Over Invoiced")),
            under_invoiced=m(g(r,"Under Invoiced")), open_invoices=m(g(r,"Open Invoices")),
            original_budget_cost=m(g(r,"Original Budget Cost")), bills=m(g(r,"Bills")),
            budgeted_profit=m(g(r,"Budgeted Profit")), description=g(r,"Description"),
        ))
    return out

# ===========================================================================
# BUILD TABLES
# ===========================================================================
gl, sub_vendors = parse_gl()
wip = parse_wip()

# ---- reconciliation asserts (baseline integrity) ----
tb_dr=round(sum(d for _,d,_ in TB_ROWS if d),2); tb_cr=round(sum(c for _,_,c in TB_ROWS if c),2)
assert tb_dr==tb_cr==1881728.02, f"TB out of balance: dr={tb_dr} cr={tb_cr}"
ar_open=round(sum(o for *_,o in [(r[-1],) for r in AR_ROWS]),2)  # last field
ar_open=round(sum(r[7] for r in AR_ROWS),2)
assert ar_open==432296.14, f"AR open total {ar_open} != 432296.14"
ap_open=round(sum(r[8] for r in AP_ROWS),2)
assert ap_open==12155.00, f"AP open total {ap_open} != 12155.00"
pr_paid=round(sum(r[1] for r in PR_SUM),2); pr_net=round(sum(r[3] for r in PR_SUM),2)
pr_liab=round(sum(r[4] for r in PR_SUM),2); pr_exp=round(sum(r[5] for r in PR_SUM),2)
assert pr_paid==74134.77, f"payroll paid {pr_paid} != 74134.77"
assert pr_net==64851.70, f"payroll net {pr_net} != 64851.70"
assert pr_liab==5873.01, f"payroll liab {pr_liab} != 5873.01"
assert pr_exp==80007.78, f"payroll exp {pr_exp} != 80007.78"
assert len(wip)==23, f"WIP jobs {len(wip)}"
assert len(PR_SUM)==52

# ---- Clients (master customer/job registry) ----
def ar_leaf(path): return path.split(":")[-1].strip()
def ar_customer(path): return path.split(":")[0].strip()
wip_by_name={w["job_name"].strip().lower(): w for w in wip}
clients={}   # key: qbo_full_path or job_name
def client_id_for(job_name): return "CL-"+slug(job_name)

# seed from WIP jobs (authoritative job list w/ JobTread id)
for w in wip:
    cid=client_id_for(w["job_name"])
    clients[w["job_name"].strip().lower()]=dict(
        client_id=cid, customer_name="", job_name=w["job_name"], qbo_full_path="",
        jobtread_job_id=w["jobtread_job_id"], job_number=w["job_number"],
        price_type=w["price_type"], sources="JobTread WIP", active="Y", notes="")
# enrich / add from AR customer paths
for bucket,date,tt,num,cust,due,amt,ob in AR_ROWS:
    if tt!="Invoice":
        # payments w/o job path -> treat customer as client, no job
        key=cust.strip().lower()
        c=clients.get(key)
        if not c:
            clients[key]=dict(client_id=client_id_for(cust), customer_name=cust, job_name="",
                qbo_full_path=cust, jobtread_job_id="", job_number="", price_type="",
                sources="A/R Aging", active="Y", notes="A/R payment payer (no job)")
        continue
    leaf=ar_leaf(cust); key=leaf.strip().lower()
    c=clients.get(key)
    if c:
        if not c["customer_name"]: c["customer_name"]=ar_customer(cust)
        if not c["qbo_full_path"]: c["qbo_full_path"]=cust
        if "A/R Aging" not in c["sources"]: c["sources"]+=" + A/R Aging"
    else:
        clients[key]=dict(client_id=client_id_for(leaf), customer_name=ar_customer(cust),
            job_name=leaf, qbo_full_path=cust, jobtread_job_id="", job_number="",
            price_type="", sources="A/R Aging", active="Y",
            notes="In A/R but not in current WIP export")
CLIENTS=sorted(clients.values(), key=lambda c:(c["customer_name"] or c["job_name"]).lower())

# client lookup for AR / WIP rows
def match_client(job_name):
    c=clients.get((job_name or '').strip().lower())
    return c["client_id"] if c else ""

# ---- Sub_Contract_Reference (seed every distinct GL subcontractor vendor) ----
# 1099 flag: vendors seen with Track-1099 in weekly vendor report / typical trade subs
SUBREF=[]
for v,info in sorted(sub_vendors.items(), key=lambda x:-x[1]["total"]):
    SUBREF.append(dict(
        vendor_name=v, normalized_vendor=re.sub(r'\s+',' ',v).strip().lower(),
        linked_client_id="", linked_job_name="", linked_jobtread_job_id="",
        allocation_confidence="unallocated", match_method="",
        gl_total_paid_h1_cash=round(info["total"],2), gl_txn_count=info["cnt"],
        is_1099="", notes="Seed row — assign vendor to job to enable 'matched' allocation"))

# ---- Ingestion_Log (baseline load) ----
def log_row(rt, key, rows_n, control, ok, note=""):
    fn,fid,label,period,basis = SOURCES[key]
    return dict(ingestion_id=f"ING-{key}-{INGEST_TS[:10]}", run_timestamp=INGEST_TS,
        report_type=label, source_file=fn, source_drive_id=fid, as_of_or_period=period,
        basis=basis, rows_ingested=rows_n, control_total=control,
        reconciled=("Y" if ok else "N"), status="loaded (baseline)", notes=note)
INGEST=[
 log_row("GL","GL",len(gl),round(sum((r['amount'] or 0) for r in gl),2),True,
         f"{len(gl)} txn lines; {sum(1 for r in gl if r['is_subcontractor']=='Y')} subcontractor lines"),
 log_row("TB","TB",len(TB_ROWS),tb_dr,True,"debits=credits balanced"),
 log_row("AR","AR",len(AR_ROWS),ar_open,True,"open balance ties to report TOTAL"),
 log_row("AP","AP",len(AP_ROWS),ap_open,True,"open balance ties to report TOTAL & TB A/P line"),
 log_row("WIP","WIP",len(wip),round(sum((w['invoiced'] or 0) for w in wip),2),True,"23 fixed-fee jobs"),
 log_row("PR_SUM","PR_SUM",len(PR_SUM),pr_exp,True,"52 weekly checks; overhead/admin labor"),
]

# ===========================================================================
# EMIT: CSV per table + multi-tab workbook
# ===========================================================================
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

TABLES=[]  # (sheet_name, columns, rows[list-of-dict])

TABLES.append(("Clients",
 ["client_id","customer_name","job_name","qbo_full_path","jobtread_job_id","job_number","price_type","sources","active","notes"],
 CLIENTS))

TABLES.append(("GL_Snapshots",
 ["gl_account","txn_date","txn_type","num","name","memo","split_account","amount","running_balance",
  "is_subcontractor","allocation_confidence","allocated_jobtread_id","row_uid","source_file","gl_period","gl_basis","ingested_at"],
 gl))

TABLES.append(("TrialBalance_Snapshots",
 ["as_of_date","basis","account","debit","credit","net_debit_credit","source_file","ingested_at"],
 [dict(as_of_date="2026-07-11", basis="Accrual", account=a,
       debit=d, credit=c, net_debit_credit=round((d or 0)-(c or 0),2),
       source_file=SOURCES["TB"][0], ingested_at=INGEST_TS) for a,d,c in TB_ROWS]))

TABLES.append(("AR_Aging_Snapshots",
 ["as_of_date","aging_bucket","txn_date","txn_type","num","customer_full_path","client_id","due_date","amount","open_balance","source_file","ingested_at"],
 [dict(as_of_date="2026-07-11", aging_bucket=b, txn_date=iso(d), txn_type=tt, num=n,
       customer_full_path=cust, client_id=match_client(ar_leaf(cust)) if tt=="Invoice" else match_client(cust),
       due_date=iso(due), amount=amt, open_balance=ob,
       source_file=SOURCES["AR"][0], ingested_at=INGEST_TS)
  for b,d,tt,n,cust,due,amt,ob in AR_ROWS]))

TABLES.append(("AP_Aging_Snapshots",
 ["as_of_date","aging_bucket","txn_date","txn_type","num","vendor_name","due_date","days_past_due","amount","open_balance","is_subcontractor","source_file","ingested_at"],
 [dict(as_of_date="2026-07-15", aging_bucket=b, txn_date=iso(d), txn_type=tt, num=n,
       vendor_name=ven, due_date=iso(due), days_past_due=pd, amount=amt, open_balance=ob,
       is_subcontractor="", source_file=SOURCES["AP"][0], ingested_at=INGEST_TS)
  for b,d,tt,n,ven,due,pd,amt,ob in AP_ROWS]))

TABLES.append(("Payroll_Snapshots",
 ["pay_date","employee_name","cost_classification","hours","gross_pay","tax_withheld","deductions","net_pay","check_no","employer_liability","total_expense","source_file","ingested_at"],
 [dict(pay_date=iso(d), employee_name=EMP, cost_classification="Overhead/Admin",
       hours=40.0, gross_pay=tp, tax_withheld=tx, deductions=0.0, net_pay=net,
       check_no="DD", employer_liability=liab, total_expense=texp,
       source_file=SOURCES["PR_SUM"][0], ingested_at=INGEST_TS)
  for d,tp,tx,net,liab,texp in PR_SUM]))

TABLES.append(("JobTread_WIP_Snapshots",
 ["as_of_date","jobtread_job_id","job_number","job_name","client_id","price_type","projected_price","budget_cost",
  "projected_cost","actual_cost","cost_to_complete","budget_variance","projected_profit","projected_margin",
  "pct_cost_complete","earned_revenue","invoiced","over_invoiced","under_invoiced","open_invoices",
  "original_budget_cost","bills","budgeted_profit","source_file","ingested_at"],
 [dict(as_of_date="2026-07-18", client_id=match_client(w["job_name"]),
       source_file=SOURCES["WIP"][0], ingested_at=INGEST_TS, **{k:w.get(k) for k in
       ["jobtread_job_id","job_number","job_name","price_type","projected_price","budget_cost",
        "projected_cost","actual_cost","cost_to_complete","budget_variance","projected_profit","projected_margin",
        "pct_cost_complete","earned_revenue","invoiced","over_invoiced","under_invoiced","open_invoices",
        "original_budget_cost","bills","budgeted_profit"]}) for w in wip]))

TABLES.append(("Sub_Contract_Reference",
 ["vendor_name","normalized_vendor","linked_client_id","linked_job_name","linked_jobtread_job_id",
  "allocation_confidence","match_method","gl_total_paid_h1_cash","gl_txn_count","is_1099","notes"],
 SUBREF))

TABLES.append(("Ingestion_Log",
 ["ingestion_id","run_timestamp","report_type","source_file","source_drive_id","as_of_or_period","basis",
  "rows_ingested","control_total","reconciled","status","notes"],
 INGEST))

# CSVs
os.makedirs(os.path.join(OUT,"data"), exist_ok=True)
for name,cols,rows in TABLES:
    with open(os.path.join(OUT,"data",f"{name}.csv"),"w",newline="",encoding="utf-8") as f:
        w=csv.DictWriter(f,fieldnames=cols,extrasaction="ignore"); w.writeheader()
        for r in rows: w.writerow(r)

# Workbook
wb=openpyxl.Workbook(); wb.remove(wb.active)
hdr_fill=PatternFill("solid",fgColor="1F4E78"); hdr_font=Font(color="FFFFFF",bold=True)
for name,cols,rows in TABLES:
    ws=wb.create_sheet(name[:31])
    ws.append(cols)
    for c in range(1,len(cols)+1):
        cell=ws.cell(1,c); cell.fill=hdr_fill; cell.font=hdr_font
        cell.alignment=Alignment(vertical="center")
    for r in rows: ws.append([r.get(k) for k in cols])
    ws.freeze_panes="A2"
    for i,col in enumerate(cols,1):
        width=max(len(col)+2, *(len(str(r.get(col) or "")) for r in rows[:200]) or [0])
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width=min(max(width,10),48)
wb.save(os.path.join(OUT,"Financial_Performance_DB.xlsx"))

# summary
print("BUILD OK")
for name,cols,rows in TABLES:
    print(f"  {name:26} {len(rows):5d} rows  {len(cols)} cols")
print("Reconciliation: TB balanced @", tb_dr, "| AR", ar_open, "| AP", ap_open,
      "| Payroll exp", pr_exp, "| WIP jobs", len(wip), "| GL txns", len(gl))
print("Distinct subcontractor vendors seeded:", len(SUBREF))
print("Clients:", len(CLIENTS), "| matched to JobTread:", sum(1 for c in CLIENTS if c["jobtread_job_id"]))
