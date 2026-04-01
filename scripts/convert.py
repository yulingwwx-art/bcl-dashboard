"""
BCL Dashboard 資料轉換腳本
每次 xlsx 上傳到 GitHub 後自動執行，輸出 data/bcl_data.json
"""
import pandas as pd, numpy as np, json, os, glob, warnings
from datetime import datetime
warnings.filterwarnings('ignore')

# ── 找 Excel 檔案（xlsx 優先）────────────────────────────────────────────────
def find_excel():
    for pattern in ['*.xlsx', '*.xlsm']:
        files = glob.glob(pattern)
        if files:
            print(f"✅ 找到檔案: {files[0]}")
            return files[0]
    raise FileNotFoundError("找不到 Excel 檔案，請確認已上傳到 repo 根目錄")

FILE = find_excel()

def read_sheet(name):
    try:
        return pd.read_excel(FILE, sheet_name=name, header=None, engine='openpyxl')
    except Exception as e:
        print(f"  ⚠️ 無法讀取 [{name}]：{e}")
        return None

def si(v):   # safe int
    try: f = float(v); return 0 if np.isnan(f) else int(f)
    except: return 0

def sf(v, d=2):  # safe float
    try: f = float(v); return 0.0 if np.isnan(f) else round(f, d)
    except: return 0.0

def is_fy25(yr, mo):
    return (yr == 2025 and mo >= 4) or (yr == 2026 and mo <= 3)

# ════════════════════════════════════════════════════
# 1. FY24 月份彙總
# ════════════════════════════════════════════════════
print("\n[1/7] FY24 月份數據...")
fy24_monthly = {'主動': [], '自辦': [], '異業': [], '總計': []}
df = read_sheet('FY24每日')
if df is not None:
    d = df.iloc[5:].copy()
    d.columns = ['年','月','日','主動on','自辦off','自辦on','異業off','總計']
    d = d.apply(pd.to_numeric, errors='coerce')
    d['年'] = d['年'].ffill(); d['月'] = d['月'].ffill()
    d = d.dropna(subset=['日']); d = d[d['日']>0].fillna(0)
    d['主動'] = d['主動on']; d['自辦'] = d['自辦off']+d['自辦on']; d['異業'] = d['異業off']
    m = d.groupby(['年','月']).agg({'主動':'sum','自辦':'sum','異業':'sum','總計':'sum'}).reset_index()
    fy24 = m[((m['年']==2024)&(m['月']>=4))|((m['年']==2025)&(m['月']<=3))]
    fy24_monthly = {k: [si(v) for v in fy24[k]] for k in ['主動','自辦','異業','總計']}

# ════════════════════════════════════════════════════
# 2. FY25 月份彙總
# ════════════════════════════════════════════════════
print("[2/7] FY25 月份數據...")
fy25_monthly = {'主動':[],'自辦':[],'異業':[],'總計':[],'月份':[]}
df = read_sheet('最大塊數據_月齡x管道')
if df is not None:
    raw = df.iloc[5:, 13:19].copy()
    raw.columns = ['年','月','主動','自辦','異業','總計']
    raw = raw.apply(pd.to_numeric, errors='coerce')
    raw['年'] = raw['年'].ffill()
    raw = raw.dropna(subset=['月','總計'])
    fy25 = raw[raw.apply(lambda r: is_fy25(r['年'], r['月']), axis=1)].copy()
    fy25_monthly = {
        '主動':  [si(v) for v in fy25['主動']],
        '自辦':  [si(v) for v in fy25['自辦']],
        '異業':  [si(v) for v in fy25['異業']],
        '總計':  [si(v) for v in fy25['總計']],
        '月份':  [f"{int(r['年'])}/{int(r['月'])}" for _,r in fy25.iterrows()],
    }

# ════════════════════════════════════════════════════
# 3. 每日活動（前一天活動）
# ════════════════════════════════════════════════════
print("[3/7] 每日活動數據...")
daily_all = {'dates':[],'主動':[],'自辦':[],'異業':[],'總計':[]}
last8 = {'dates':[],'主動':[],'自辦':[],'異業':[],'總計':[]}
yesterday = {}; prev7_avg = {}

df = read_sheet('前一天活動')
if df is not None:
    d = df.iloc[5:].copy()
    d.columns = ['年','月','日','主動','自辦','異業','總計']
    d = d.apply(pd.to_numeric, errors='coerce')
    d['年'] = d['年'].ffill(); d['月'] = d['月'].ffill()
    d = d.dropna(subset=['日']); d = d[d['日']>0].fillna(0)
    d25 = d[d.apply(lambda r: is_fy25(r['年'],r['月']), axis=1)].copy()

    daily_all = {
        'dates': [f"{int(r['年'])}/{int(r['月'])}/{int(r['日'])}" for _,r in d25.iterrows()],
        '主動':  [si(r['主動'])  for _,r in d25.iterrows()],
        '自辦':  [si(r['自辦'])  for _,r in d25.iterrows()],
        '異業':  [si(r['異業'])  for _,r in d25.iterrows()],
        '總計':  [si(r['總計'])  for _,r in d25.iterrows()],
    }
    tail8 = d25.tail(8)
    last8 = {
        'dates': [f"{int(r['月'])}/{int(r['日'])}" for _,r in tail8.iterrows()],
        '主動':  [si(r['主動'])  for _,r in tail8.iterrows()],
        '自辦':  [si(r['自辦'])  for _,r in tail8.iterrows()],
        '異業':  [si(r['異業'])  for _,r in tail8.iterrows()],
        '總計':  [si(r['總計'])  for _,r in tail8.iterrows()],
    }
    if len(tail8) >= 2:
        lr = tail8.iloc[-1]; p7 = tail8.iloc[:-1]
        yesterday = {
            'date': f"{int(lr['年'])}/{int(lr['月'])}/{int(lr['日'])}",
            '主動': si(lr['主動']), '自辦': si(lr['自辦']),
            '異業': si(lr['異業']), '總計': si(lr['總計']),
        }
        prev7_avg = {k: sf(p7[k].mean()) for k in ['主動','自辦','異業','總計']}

# ════════════════════════════════════════════════════
# 4. 每日月齡分布（前一天月齡-全）
# ════════════════════════════════════════════════════
print("[4/7] 月齡分布數據...")
fy25_monthly_age = {'孕期':[],'m036':[],'m36p':[],'nobday':[]}
yesterday_age = {}; prev7_age_avg = {}

df = read_sheet('前一天月齡-全')
if df is not None:
    hdr = list(df.iloc[5])
    nineK = next((i for i,c in enumerate(hdr) if str(c).strip()=='9999.0' or (isinstance(c,float) and c==9999)), None)
    neg_cols  = [i for i,c in enumerate(hdr) if isinstance(c,(int,float)) and -99<c<0]
    m036_cols = [i for i,c in enumerate(hdr) if isinstance(c,(int,float)) and 0<=c<=36]
    m36p_cols = [i for i,c in enumerate(hdr) if isinstance(c,(int,float)) and 36<c<9999]

    rows = df.iloc[6:].copy()
    meta = rows.iloc[:,:3].copy(); meta.columns=['年','月','日']
    meta = meta.apply(pd.to_numeric, errors='coerce')
    meta['年'] = meta['年'].ffill(); meta['月'] = meta['月'].ffill()
    meta = meta.dropna(subset=['日']); meta = meta[meta['日']>0]

    results = []
    for idx in meta.index:
        row = rows.loc[idx]
        yr,mo = meta.loc[idx,'年'], meta.loc[idx,'月']
        if not is_fy25(yr, mo): continue
        pre  = sum(sf(row.iloc[i]) for i in neg_cols  if i<len(row))
        m036 = sum(sf(row.iloc[i]) for i in m036_cols if i<len(row))
        m36p = sum(sf(row.iloc[i]) for i in m36p_cols if i<len(row))
        nb   = sf(row.iloc[nineK]) if nineK and nineK<len(row) else 0.0
        results.append({'年':yr,'月':mo,'日':meta.loc[idx,'日'],
                        '孕期':pre,'m036':m036,'m36p':max(0,m36p-nb),'nobday':nb})

    if results:
        dr = pd.DataFrame(results)
        ma = dr.groupby(['年','月']).agg({'孕期':'sum','m036':'sum','m36p':'sum','nobday':'sum'}).reset_index()
        fy25_monthly_age = {k:[si(v) for v in ma[k]] for k in ['孕期','m036','m36p','nobday']}
        tail8a = dr.tail(8); prev7a = dr.iloc[-8:-1] if len(dr)>=8 else dr.iloc[:-1]
        yesterday_age  = {k:si(dr.iloc[-1][k]) for k in ['孕期','m036','m36p','nobday']}
        prev7_age_avg  = {k:sf(prev7a[k].mean()) for k in ['孕期','m036','m36p','nobday']}

# ════════════════════════════════════════════════════
# 5. 系統保有月齡（含 999999 無生日）
# ════════════════════════════════════════════════════
print("[5/7] 系統保有月齡...")
stock_age = {'孕期':{'total':0,'valid':0,'invalid':0},'m036':{'total':0,'valid':0,'invalid':0},
             'm36p':{'total':0,'valid':0,'invalid':0},'nobday':{'total':0,'valid':0,'invalid':0},
             'grand_total':0,'valid_total':0}

df = read_sheet('系統保有&綁定月齡')
if df is not None:
    s = df.iloc[5:,:5].copy(); s.columns=['出生年齡','月齡','有效','無效','總計']
    s = s.apply(pd.to_numeric, errors='coerce').dropna(subset=['總計'])
    s = s[s['總計']>0]
    def agg(rows): return {'total':si(rows['總計'].sum()),'valid':si(rows['有效'].sum()),'invalid':si(rows['無效'].sum())}
    stock_age = {
        '孕期':  agg(s[s['月齡']<0]),
        'm036':  agg(s[(s['月齡']>=0)&(s['月齡']<=36)]),
        'm36p':  agg(s[(s['月齡']>36)&(s['月齡']<999)]),
        'nobday':agg(s[s['月齡']>=999]),
        'grand_total': si(s['總計'].sum()),
        'valid_total':  si(s['有效'].sum()),
    }

# ════════════════════════════════════════════════════
# 6. 管道月齡品質（綁定月齡，正確含 9999）
# ════════════════════════════════════════════════════
print("[6/7] 管道月齡品質...")
channel_age = {
    '主動': {'孕期':414,'m036':3783,'m36p':44,'nobday':1219,'total':5460},
    '自辦': {'孕期':1744,'m036':15779,'m36p':117,'nobday':419,'total':18059},
    '異業': {'孕期':3111,'m036':4623,'m36p':170,'nobday':1081,'total':8985},
}
df = read_sheet('綁定月齡')
if df is not None:
    try:
        row9999 = df.iloc[55, 7:12].values   # [9999, 主動, 自辦, 異業, 總計]
        rowtotal= df.iloc[56, 7:12].values
        bind = df.iloc[6:55, 7:12].copy()
        bind.columns = ['月齡','主動','自辦','異業','總計']
        bind = bind.apply(pd.to_numeric, errors='coerce').dropna(subset=['月齡'])

        for ch, nb, tot in [
            ('主動', si(row9999[1]), si(rowtotal[1])),
            ('自辦', si(row9999[2]), si(rowtotal[2])),
            ('異業', si(row9999[3]), si(rowtotal[3])),
        ]:
            pre  = si(bind[bind['月齡'].between(-99,-1)][ch].sum())
            m036 = si(bind[bind['月齡'].between(0,36)][ch].sum())
            m36p = max(0, si(bind[bind['月齡'].between(37,998)][ch].sum()) - nb)
            channel_age[ch] = {'孕期':pre,'m036':m036,'m36p':m36p,'nobday':nb,'total':tot}
    except Exception as e:
        print(f"  ⚠️ 管道品質解析警告：{e}")

# ════════════════════════════════════════════════════
# 7. 月齡均分
# ════════════════════════════════════════════════════
print("[7/7] 月齡均分...")
quality_scores = {'整體':6.19,'主動':5.48,'自辦':7.22,'異業':4.55}
df = read_sheet('月齡評分')
if df is not None:
    try:
        for _, row in df.iloc[4:8, 8:12].iterrows():
            row.index = ['管道','筆數','均分','總分']
            ch = str(row['管道']).strip()
            sc = sf(row['均分'])
            if ch in ('主動','自辦','異業') and sc > 0:
                quality_scores[ch] = sc
        overall = sf(df.iloc[8, 10])
        if overall > 0: quality_scores['整體'] = overall
    except: pass

# ════════════════════════════════════════════════════
# 輸出 JSON
# ════════════════════════════════════════════════════
os.makedirs('data', exist_ok=True)
output = {
    'meta': {
        'updated_at': datetime.now().strftime('%Y/%m/%d %H:%M'),
        'source_file': os.path.basename(FILE),
    },
    'fy24_monthly':     fy24_monthly,
    'fy25_monthly':     fy25_monthly,
    'daily_all':        daily_all,
    'last8':            last8,
    'yesterday':        yesterday,
    'prev7_avg':        prev7_avg,
    'yesterday_age':    yesterday_age,
    'prev7_age_avg':    prev7_age_avg,
    'fy25_monthly_age': fy25_monthly_age,
    'stock_age':        stock_age,
    'channel_age':      channel_age,
    'quality_scores':   quality_scores,
}

with open('data/bcl_data.json','w',encoding='utf-8') as f:
    json.dump(output, f, ensure_ascii=False, indent=2)

# 摘要
gt = stock_age.get('grand_total', 0)
yd = yesterday.get('date','N/A')
yt = yesterday.get('總計',0)
f25 = len(fy25_monthly.get('總計',[]))
print(f"\n{'='*50}")
print(f"✅ 完成！→ data/bcl_data.json")
print(f"   更新時間：{output['meta']['updated_at']}")
print(f"   前一日：{yd}，新增 {yt} 筆")
print(f"   FY25 月份數：{f25} 個月")
print(f"   保有總計：{gt:,} 筆")
print(f"{'='*50}")
