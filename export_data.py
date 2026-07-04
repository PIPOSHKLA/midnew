#!/usr/bin/env python3
"""Export tourism and stock data from MOTS + Yahoo Finance to CSV files."""
import subprocess, sys, os, warnings, io, csv, urllib.request, datetime
warnings.filterwarnings('ignore')
for pkg in ['pandas', 'numpy', 'yfinance']:
    try:
        __import__(pkg.replace('-', '_'))
    except ImportError:
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', pkg, '-q'])
import numpy as np
import pandas as pd
import yfinance as yf

# 1. Download MOTS data
print("Downloading MOTS data...")
url = ("https://ckan.mots.go.th/dataset/445c66d8-a06a-49d9-adfc-35faca6fc785/"
       "resource/faffc63c-9507-451a-80b7-554cc0787368/download/est_2024_04_01.csv")
req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
raw = urllib.request.urlopen(req, timeout=60).read()
content = raw.decode('utf-8-sig')
reader = csv.DictReader(io.StringIO(content))
rows = list(reader)

ASEAN = {'BRUNEI','CAMBODIA','INDONESIA','LAOS','MALAYSIA','MYANMAR','PHILIPPINES','SINGAPORE','THAILAND','VIETNAM','TIMOR-LESTE','VIET NAM'}
market_map = {}
records = []
for row in rows:
    country = row['Country'].strip().upper()
    continent = row['continent'].strip()
    if 'CHINA' in country or 'HONG KONG' in country or 'TAIWAN' in country or 'MACAO' in country:
        mkt = 'China'
    elif 'INDIA' in country:
        mkt = 'India'
    elif continent == 'Europe':
        mkt = 'Europe'
    elif country in ASEAN:
        mkt = 'ASEAN'
    else:
        market_map.setdefault('Other', []).append(row)
        continue
    try:
        num = int(row[' Number '].replace(',', '').strip())
    except:
        continue
    parts = row['date'].strip().split('/')
    if len(parts) == 3:
        d, m, y = int(parts[0]), int(parts[1]), int(parts[2])
        dt = datetime.date(y, m, d)
        records.append({'date': dt, 'market': mkt, 'arrivals': num, 'country': country, 'continent': continent})

# Aggregate to monthly
df = pd.DataFrame(records)
df['month_key'] = df['date'].apply(lambda d: datetime.date(d.year, d.month, 1))
monthly = df.groupby(['market', 'month_key'])['arrivals'].sum().reset_index()
pivot = monthly.pivot_table(index='month_key', columns='market', values='arrivals', aggfunc='sum').fillna(0).sort_index()
pivot.index = pd.to_datetime(pivot.index)
for m in ['China', 'India', 'Europe', 'ASEAN']:
    if m not in pivot.columns:
        pivot[m] = 0.0

pivot.to_csv('tourism_arrivals_monthly.csv')
print(f"  Saved: tourism_arrivals_monthly.csv ({len(pivot)} rows)")

# 2. Download stock data
print("Downloading stock data from Yahoo Finance...")
stock_symbols = ['AOT.BK', 'AAV.BK', 'MINT.BK', 'CENTEL.BK']
stock_labels = ['AOT', 'AAV', 'MINT', 'CENTEL']
start = pivot.index[0].strftime('%Y-%m-%d')
end = pivot.index[-1].strftime('%Y-%m-%d')
stock_prices = pd.DataFrame(index=pivot.index)
for sym, label in zip(stock_symbols, stock_labels):
    df = yf.download(sym, start=start, end=end, auto_adjust=True, progress=False)
    if df.empty:
        print(f"  WARNING: no data for {sym}")
        continue
    df.columns = [c[0] if isinstance(c, tuple) else c for c in df.columns]
    monthly_close = df['Close'].resample('ME').last()
    monthly_close.index = monthly_close.index.tz_localize(None)
    aligned = monthly_close.reindex(pivot.index, method='ffill')
    stock_prices[label] = aligned.values
    print(f"  {sym} -> {label}: {len(df)} daily rows")

stock_prices.to_csv('stock_prices_monthly.csv')
print(f"  Saved: stock_prices_monthly.csv ({len(stock_prices)} rows)")

# 3. Merged dataset
merged = pivot.join(stock_prices)
merged.to_csv('tourism_stock_merged.csv')
print(f"  Saved: tourism_stock_merged.csv ({len(merged)} rows, {len(merged.columns)} cols)")

# 4. Correlation matrix
from scipy.stats import spearmanr
market_names = ['China', 'India', 'Europe', 'ASEAN']
corr_data = []
for mkt in market_names:
    for stk in stock_labels:
        r, p = spearmanr(pivot[mkt].values, stock_prices[stk].values)
        corr_data.append({'market': mkt, 'stock': stk, 'spearman_r': round(r, 4), 'p_value': round(p, 4)})
corr_df = pd.DataFrame(corr_data)
corr_df.to_csv('correlation_matrix.csv', index=False)
print(f"  Saved: correlation_matrix.csv ({len(corr_df)} rows)")

# 5. Summary stats
summary = pivot.describe().round(0).astype(int)
summary.to_csv('arrivals_summary_stats.csv')
print(f"  Saved: arrivals_summary_stats.csv")

print("\nDone! Files:")
for f in ['tourism_arrivals_monthly.csv', 'stock_prices_monthly.csv', 'tourism_stock_merged.csv', 'correlation_matrix.csv', 'arrivals_summary_stats.csv']:
    size = os.path.getsize(f) if os.path.exists(f) else 0
    print(f"  - {f} ({size:,} bytes)")
