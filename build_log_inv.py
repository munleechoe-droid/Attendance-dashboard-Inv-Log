#!/usr/bin/env python3
"""
Attendance Dashboard Builder
==============================
Usage:  python3 build.py [path_to_xlsx]
        Default xlsx: 0_DailyAttendanceReport_Master.xlsx (same folder as this script)

Output: AttendanceDashboard_LogInv.html  — Logistic + Inventory departments only

Steps:  1. Drop updated xlsx in the same folder
        2. Run:  python3 build_log_inv.py
        3. Share AttendanceDashboard_LogInv.html
"""
import pandas as pd, json, os, sys, warnings, calendar, re
from datetime import datetime, timezone, timedelta
warnings.filterwarnings('ignore')

SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
DATA_FILE    = sys.argv[1] if len(sys.argv) > 1 else os.path.join(SCRIPT_DIR, '0_DailyAttendanceReport_Master.xlsx')
SHEET_NAME   = 'Master_Daily Attendance'
TEMPLATE     = os.path.join(SCRIPT_DIR, 'index_loginv.html')
OUTPUT_FILE  = os.path.join(SCRIPT_DIR, 'AttendanceDashboard_LogInv.html')  # for Google Drive
DATA_JS      = os.path.join(SCRIPT_DIR, 'data_loginv.js')              # for GitHub Pages

MONTH_MAP = {
    '01-26':'Jan 26','02-26':'Feb 26','03-26':'Mar 26','04-26':'Apr 26',
    '05-26':'May 26','06-26':'Jun 26','07-26':'Jul 26','08-26':'Aug 26',
    '09-26':'Sep 26','10-26':'Oct 26','11-26':'Nov 26','12-26':'Dec 26',
    '01-25':'Jan 25','02-25':'Feb 25','03-25':'Mar 25','04-25':'Apr 25',
}

# Malaysia time (UTC+8)
myt = timezone(timedelta(hours=8))
generated = datetime.now(myt).strftime('%d %b %Y %H:%M MYT')

print(f"Reading: {DATA_FILE}")
df = pd.read_excel(DATA_FILE, sheet_name=SHEET_NAME)
df['Date'] = pd.to_datetime(df['Date'])
df['Date Of Joining'] = pd.to_datetime(df['Date Of Joining'], errors='coerce')
df['Date Of Exit']    = pd.to_datetime(df['Date Of Exit'],    errors='coerce')
df['Employee Id'] = df['Employee Id'].astype(str).str.strip()
df['Direct Manager Employee Id'] = df['Direct Manager Employee Id'].astype(str).str.strip()

def clean(s): return str(s).replace('\u2800','').replace('\u200e','').strip() if pd.notna(s) else ''

df['Month_Label'] = df['Month'].map(MONTH_MAP).fillna(df['Month'])
df['Year']    = df['Date'].dt.year
df['Quarter'] = df['Date'].dt.quarter
df['WeekNum'] = df['Date'].dt.isocalendar().week.astype(int)
df['DayStr']  = df['Date'].dt.strftime('%d')

# Shift label: "09:00 AM - 05:30 PM" → "0900-1730"
def parse_shift(s):
    if pd.isna(s): return ''
    m = re.search(r'(\d{1,2}:\d{2}\s*[AP]M)\s*-\s*(\d{1,2}:\d{2}\s*[AP]M)', str(s))
    if not m: return ''
    def t24(t):
        t=t.strip(); h,rest=t.split(':'); mn,ap=rest[:2],rest[2:].strip()
        h,mn=int(h),int(mn)
        if ap=='PM' and h!=12: h+=12
        if ap=='AM' and h==12: h=0
        return f"{h:02d}{mn:02d}"
    return t24(m.group(1))+'-'+t24(m.group(2))

df['ShiftLabel'] = df['Shift'].apply(parse_shift)

# Duration → hours
def dur_hrs(val):
    if pd.isna(val): return 0.0
    try:
        if hasattr(val,'hour'): return round(val.hour+val.minute/60+val.second/3600,2)
        p=str(val).split(':')
        return round(int(p[0])+int(p[1])/60+(int(p[2]) if len(p)>2 else 0)/3600,2)
    except: return 0.0

df['DurHours'] = df['Final Work Duration'].apply(dur_hrs)

# Working base — all rows where employee was expected at work
# ── FILTER: Logistic + Inventory only ────────────────────────────────────────
DEPTS = ['LOGISTIC', 'INVENTORY']
df = df[df['Current Department'].isin(DEPTS)].copy()
print(f"  Filtered to {DEPTS}: {df['Employee Id'].nunique()} employees")

work = df[(df['Working Days']==1)|(df['Absent']==1)|(df['Single Punch']==1)|(df['On Leave']==1)].copy()

def sc(status, is_late, wd, sp, leave, absent):
    s = str(status) if pd.notna(status) else ''
    if 'Present' in s: return 'PL' if is_late=='Yes' else 'P'
    if sp==1:    return 'SP'
    if leave==1: return 'L'
    if absent==1:return 'A'
    return 'N'

# ── SUMMARY (no WeekNum in group — avoids duplicate rows) ────────────────────
print("Building summary...")
def agg(g):
    present = int(g['Status'].str.contains('Present',na=False).sum())
    absent  = int(g['Absent'].sum())
    sp      = int(g['Single Punch'].sum())
    leave   = int(g['On Leave'].sum())
    wd      = present + absent + sp + leave   # all days expected at work
    return pd.Series({
        'emp':    int(g['Employee Id'].nunique()),
        'wd':     wd,
        'present':present,
        'late':   int((g['Is Late ']=='Yes').sum()),
        'sp':     sp,
        'leave':  leave,
        'absent': absent,
    })

summary = work.groupby(['Current Department','Branch','Current Designation','Month_Label','Year','Quarter']).apply(agg).reset_index()
summary['att'] = (summary['present']/summary['wd']*100).where(summary['wd']>0,0).round(1)
summary_rows = summary.rename(columns={'Current Department':'dept','Branch':'branch','Current Designation':'desig'}).to_dict('records')

# Summary with WeekNum — used only when week filter is active
summary_wk = work.groupby(['Current Department','Branch','Current Designation','Month_Label','Year','Quarter','WeekNum']).apply(agg).reset_index()
summary_wk['att'] = (summary_wk['present']/summary_wk['wd']*100).where(summary_wk['wd']>0,0).round(1)
summary_wk_rows = summary_wk.rename(columns={'Current Department':'dept','Branch':'branch','Current Designation':'desig'}).to_dict('records')
print(f"  summary: {len(summary_rows)} rows | summary_wk: {len(summary_wk_rows)} rows")

# ── WEEK-DAY MAP ──────────────────────────────────────────────────────────────
week_day_map = {}
for (month, week), grp in df.groupby(['Month_Label','WeekNum']):
    if month not in week_day_map: week_day_map[month] = {}
    week_day_map[month][str(week)] = sorted(grp['DayStr'].unique().tolist())

# ── HEATMAP ───────────────────────────────────────────────────────────────────
print("Building heatmap...")
detail_heat = {}
for _, row in df.iterrows():
    branch = str(row['Branch']).strip()
    month  = str(row['Month_Label'])
    if not month or month == 'nan': continue
    emp_id = str(row['Employee Id']).strip()
    day    = row['DayStr']
    sl     = row['ShiftLabel'] or ''
    status_c = sc(row['Status'], row['Is Late '], row['Working Days'],
                  row['Single Punch'], row['On Leave'], row['Absent'])
    if branch not in detail_heat: detail_heat[branch] = {}
    if month  not in detail_heat[branch]: detail_heat[branch][month] = {}
    if emp_id not in detail_heat[branch][month]:
        detail_heat[branch][month][emp_id] = {
            'n': clean(row['Employee Name']),
            'd': str(row['Current Department']).strip(),
            'dg':str(row['Current Designation']).strip(),
            'm': str(row['Direct Manager Employee Id']).strip(),
            'days': {}
        }
    if status_c != 'N':
        entry = {'sc': status_c}
        if sl: entry['sh'] = sl
        detail_heat[branch][month][emp_id]['days'][day] = entry

# ── WEEKLY ANALYSIS ───────────────────────────────────────────────────────────
print("Building weekly analysis...")
emp_week = work.groupby(['Employee Id','WeekNum']).agg(
    branch=('Branch','first'),
    dept=('Current Department','first'),
    desig=('Current Designation','first'),
    wd=('Working Days','sum'),
    present=('Status', lambda x: x.str.contains('Present',na=False).sum()),
    sp=('Single Punch','sum'),
    leave=('On Leave','sum'),
    absent=('Absent','sum'),
    rest_wknd=('Rest Day on Weekend','max'),
).reset_index()
hrs_pw = df[df['DurHours']>0].groupby(['Employee Id','WeekNum'])['DurHours'].sum().round(1).reset_index()
hrs_pw.columns = ['Employee Id','WeekNum','hrs']
emp_week = emp_week.merge(hrs_pw, on=['Employee Id','WeekNum'], how='left')
emp_week['hrs']   = emp_week['hrs'].fillna(0).round(1)
emp_week['days6'] = emp_week['present'] + emp_week['leave'] + emp_week['sp']

def bw_agg(g):
    return pd.Series({
        'emp':       int(g['Employee Id'].nunique()),
        'avg_hrs':   round(g['hrs'].mean(), 1),
        'hit45':     int((g['hrs']>=45).sum()),
        'miss45':    int((g['hrs']<45).sum()),
        'hit6days':  int((g['days6']>=6).sum()),
        'miss6days': int((g['days6']<6).sum()),
        'rest_wknd': int(g['rest_wknd'].sum()),
    })

branch_week = emp_week.groupby(['branch','WeekNum']).apply(bw_agg).reset_index()
weekly_rows = branch_week.to_dict('records')

# IDX: id=0,wk=1,br=2,dp=3,dg=4,hrs=5,pres=6,lv=7,sp=8,abs=9,d6=10,rw=11
emp_wk_compact = [
    [str(r['Employee Id']), int(r['WeekNum']), str(r['branch']), str(r['dept']), str(r['desig']),
     float(r['hrs']), int(r['present']), int(r['leave']), int(r['sp']),
     int(r['absent']), int(r['days6']), int(r['rest_wknd'])]
    for _, r in emp_week.iterrows()
]

emp_names = {str(eid): clean(name) for eid, name in df.groupby('Employee Id')['Employee Name'].first().items()}

# ── DOJ / DOE MAPS ────────────────────────────────────────────────────────────
emp_dates = df.groupby('Employee Id').agg(doj=('Date Of Joining','first'), doe=('Date Of Exit','first')).reset_index()
doj_map = {str(r['Employee Id']): r['doj'].strftime('%Y-%m-%d') for _,r in emp_dates.iterrows() if pd.notna(r['doj'])}
doe_map = {str(r['Employee Id']): r['doe'].strftime('%Y-%m-%d') for _,r in emp_dates.iterrows() if pd.notna(r['doe'])}

# ── ORG ───────────────────────────────────────────────────────────────────────
print("Building org...")
ep = df.groupby('Employee Id').agg(
    name=('Employee Name', lambda x: clean(x.iloc[0])),
    dept=('Current Department','first'),
    desig=('Current Designation','first'),
    branch=('Branch','first'),
    mgr_id=('Direct Manager Employee Id','first'),
    mgr_name=('Direct Manager Name', lambda x: clean(x.iloc[0])),
).reset_index()
ws_agg = work.groupby('Employee Id').agg(
    present=('Status', lambda x: x.str.contains('Present',na=False).sum()),
    absent =('Absent','sum'),
    sp     =('Single Punch','sum'),
    leave  =('On Leave','sum'),
    late   =('Is Late ', lambda x: (x=='Yes').sum()),
).reset_index()
ws_agg['wd'] = ws_agg['present'] + ws_agg['absent'] + ws_agg['sp'] + ws_agg['leave']
ep = ep.merge(ws_agg, on='Employee Id', how='left')
ep['att'] = (ep['present']/ep['wd']*100).round(1).fillna(0)
ep['Employee Id'] = ep['Employee Id'].astype(str)
ep['mgr_id']      = ep['mgr_id'].astype(str)
org_list = ep.to_dict('records')

# ── FILTER OPTIONS ────────────────────────────────────────────────────────────
month_order = [m for m in [
    'Jan 25','Feb 25','Mar 25','Apr 25','May 25','Jun 25',
    'Jul 25','Aug 25','Sep 25','Oct 25','Nov 25','Dec 25',
    'Jan 26','Feb 26','Mar 26','Apr 26','May 26','Jun 26',
    'Jul 26','Aug 26','Sep 26','Oct 26','Nov 26','Dec 26',
] if m in df['Month_Label'].unique()]

months_num = {'Jan':1,'Feb':2,'Mar':3,'Apr':4,'May':5,'Jun':6,
              'Jul':7,'Aug':8,'Sep':9,'Oct':10,'Nov':11,'Dec':12}
month_cfg = {}
for ml in month_order:
    p = ml.split(); mn = months_num[p[0]]; yr = 2000+int(p[1])
    month_cfg[ml] = {
        'days':  calendar.monthrange(yr, mn)[1],
        'start': (calendar.weekday(yr, mn, 1)+1) % 7,
    }

weeks = sorted(df['WeekNum'].unique().tolist())
week_labels = {
    str(w): f"W{w} ({df[df['WeekNum']==w]['Date'].min().strftime('%d %b')} – {df[df['WeekNum']==w]['Date'].max().strftime('%d %b')})"
    for w in weeks
}

# ── ROSTER (future dates from WeeklyOff + Shift Variance files) ──────────────
SHIFT_FILE = os.path.join(SCRIPT_DIR, 'ShiftVarianceReport.xlsx')
WO_FILE    = os.path.join(SCRIPT_DIR, 'WeeklyOffVarianceReport.xlsx')

roster_heat = {}

def build_roster():
    if not os.path.exists(WO_FILE):
        print("  No WeeklyOffVarianceReport.xlsx found — skipping roster")
        return

    print("Building roster (future dates)...")
    wo = pd.read_excel(WO_FILE, sheet_name='Data', dtype={'Employee Id': str})
    wo['Date'] = pd.to_datetime(wo['Date'])
    wo['Employee Id'] = wo['Employee Id'].astype(str).str.strip()
    last_att = df['Date'].max()
    wo_future = wo[wo['Date'] > last_att].copy()
    print(f"  WO future rows: {len(wo_future)} | Employees: {wo_future['Employee Id'].nunique()}")

    # Shift file (optional — only covers subset of employees)
    shift_lookup = {}
    if os.path.exists(SHIFT_FILE):
        sv = pd.read_excel(SHIFT_FILE, sheet_name='Data', dtype={'Employee Id': str})
        sv['Shift Date'] = pd.to_datetime(sv['Shift Date'])
        sv['Employee Id'] = sv['Employee Id'].astype(str).str.strip()
        sv_future = sv[sv['Shift Date'] > last_att].copy()
        sv_future['ShiftLabel'] = sv_future['Current Shift'].apply(parse_shift)
        for _, row in sv_future.iterrows():
            eid = str(row['Employee Id'])
            dt  = row['Shift Date'].strftime('%Y-%m-%d')
            sl  = row['ShiftLabel']
            if sl:
                if eid not in shift_lookup: shift_lookup[eid] = {}
                shift_lookup[eid][dt] = sl
        print(f"  Shift future rows: {len(sv_future)} | With shift label: {len(shift_lookup)}")

    def extract_rest_days(wo_str):
        if pd.isna(wo_str): return set()
        s = str(wo_str).lower()
        days_map = {'monday':0,'tuesday':1,'wednesday':2,'thursday':3,
                    'friday':4,'saturday':5,'sunday':6}
        result = set()
        for day, num in days_map.items():
            # Match both "(Rest Day)" and "(Off Day)" — both are non-working days
            if re.search(day + r'\s*\((rest day|off day)\)', s):
                result.add(num)
        return result

    # Employee profile from master attendance
    profile_dict = {}
    for _, r in df.groupby('Employee Id').agg(
        name=('Employee Name', lambda x: clean(x.iloc[0])),
        branch=('Branch','first'),
        dept=('Current Department','first'),
        desig=('Current Designation','first'),
        mgr_id=('Direct Manager Employee Id','first'),
    ).reset_index().iterrows():
        profile_dict[str(r['Employee Id'])] = {
            'name': r['name'], 'branch': str(r['branch']).strip(),
            'dept': str(r['dept']).strip(), 'desig': str(r['desig']).strip(),
            'mgr_id': str(r['mgr_id']).strip()
        }

    for _, row in wo_future.iterrows():
        eid  = str(row['Employee Id'])
        prof = profile_dict.get(eid)
        if not prof: continue  # skip employees not in master attendance

        dt     = row['Date']
        branch = prof['branch']
        ml     = MONTH_MAP.get(dt.strftime('%m-%y'), dt.strftime('%b %y').replace(' ','').capitalize())
        # Normalise to match MONTH_MAP format e.g. 'Apr 26'
        ml = dt.strftime('%b %y').replace(' ',' ').title()
        # Fix: strftime gives 'Apr 26' format directly
        ml = dt.strftime('%b') + ' ' + dt.strftime('%y')
        day    = dt.strftime('%d')
        dt_str = dt.strftime('%Y-%m-%d')

        rest_days = extract_rest_days(row['Current Weekly Off'])
        sc = 'R' if dt.weekday() in rest_days else 'W'
        sl = shift_lookup.get(eid, {}).get(dt_str, '')

        if branch not in roster_heat: roster_heat[branch] = {}
        if ml not in roster_heat[branch]: roster_heat[branch][ml] = {}
        if eid not in roster_heat[branch][ml]:
            roster_heat[branch][ml][eid] = {
                'n': prof['name'], 'd': prof['dept'],
                'dg': prof['desig'], 'm': prof['mgr_id'], 'days': {}
            }
        entry = {'sc': sc}
        if sl: entry['sh'] = sl
        roster_heat[branch][ml][eid]['days'][day] = entry

    total = sum(len(roster_heat[b][m]) for b in roster_heat for m in roster_heat[b])
    print(f"  Roster: {len(roster_heat)} branches, {total} employee-month combos")

build_roster()

# ── ASSEMBLE DATA ─────────────────────────────────────────────────────────────
out = {
    'summary':      summary_rows,
    'summary_wk':   summary_wk_rows,
    'weekly_branch':weekly_rows,
    'weekly_emp':   emp_wk_compact,
    'emp_names':    emp_names,
    'heat':         detail_heat,
    'roster_heat':  roster_heat,
    'org':          org_list,
    'doj_map':      doj_map,
    'doe_map':      doe_map,
    'branches':     sorted(df['Branch'].dropna().unique().tolist()),
    'departments':  sorted(df['Current Department'].dropna().unique().tolist()),
    'designations': sorted(df['Current Designation'].dropna().unique().tolist()),
    'months':       month_order,
    'years':        sorted(df['Year'].unique().tolist()),
    'quarters':     sorted(df['Quarter'].unique().tolist()),
    'weeks':        weeks,
    'week_labels':  week_labels,
    'week_day_map': week_day_map,
    'month_cfg':    month_cfg,
    'generated':    generated,
}

data_js = 'var D = ' + json.dumps(out) + ';'
print(f"  Data size: {len(data_js)//1024} KB")

# ── MERGE INTO SINGLE HTML (Google Drive) ────────────────────────────────────
print(f"Merging into {OUTPUT_FILE}...")
with open(TEMPLATE, 'r', encoding='utf-8') as f:
    html = f.read()

if '<script src="data_loginv.js"></script>' not in html:
    print("ERROR: Could not find <script src=\"data_loginv.js\"></script> in index_loginv.html")
    print("Make sure index.html is in the same folder as build.py")
    sys.exit(1)

merged = html.replace('<script src="data_loginv.js"></script>', '<script>\n' + data_js + '\n</script>')

with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
    f.write(merged)

# ── ALSO WRITE data.js (GitHub Pages — upload index.html + data.js) ───────────
with open(DATA_JS, 'w', encoding='utf-8') as f:
    f.write(data_js)

sz = os.path.getsize(OUTPUT_FILE)
print(f"\n✅ Done!")
print(f"   📁 AttendanceDashboard.html ({sz//1024} KB) → share via Google Drive")
print(f"   📁 data.js ({os.path.getsize(DATA_JS)//1024} KB) + index.html → upload both to GitHub Pages")
print(f"   Generated: {generated}")
print(f"   Employees: {df['Employee Id'].nunique()} | Branches: {df['Branch'].nunique()} | Months: {len(month_order)}")
