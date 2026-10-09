#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
🚀 NIFTY 500 STOCK ANALYZER - NSE से Direct Download करता है
"""

import pandas as pd
import numpy as np
import yfinance as yf
import subprocess
import sys
from datetime import datetime, timedelta

# Auto install packages
for pkg in ['pandas', 'numpy', 'yfinance', 'openpyxl', 'requests', 'lxml']:
    try:
        __import__(pkg)
    except:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"])

print("\n📥 NSE से Nifty 500 stocks download कर रहे हैं...\n")

# ========================
# NSE से Nifty 500 download करें
# ========================
def download_nifty_500():
    """
    NSE की official website से Nifty 500 stocks download करता है
    """
    try:
        # Option 1: NSE Index का data
        print("🔄 Method 1: NSE API से download कर रहे हैं...")
        url = "https://www.nseindia.com/api/index-constituents?index=NIFTY%20500"
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
        
        try:
            import requests
            response = requests.get(url, headers=headers, timeout=10)
            if response.status_code == 200:
                data = response.json()
                if 'constituents' in data:
                    symbols = [f"{item['symbol']}.NS" for item in data['constituents']]
                    print(f"✅ NSE से {len(symbols)} stocks download हुए\n")
                    return symbols
        except:
            pass
        
        # Option 2: Yahoo Finance से
        print("🔄 Method 2: Yahoo Finance से download कर रहे हैं...")
        try:
            df = pd.read_html("https://en.wikipedia.org/wiki/NIFTY_500")[0]
            symbols = [f"{sym.strip()}.NS" for sym in df['Symbol'].tolist()]
            print(f"✅ Wikipedia से {len(symbols)} stocks download हुए\n")
            return symbols
        except:
            pass
        
        # Option 3: CSV file से (अगर internet issue है)
        print("🔄 Method 3: Online CSV से download कर रहे हैं...")
        try:
            csv_url = "https://raw.githubusercontent.com/adi-git-repo/nifty500/main/nifty500.csv"
            df = pd.read_csv(csv_url)
            symbols = [f"{sym.strip()}.NS" for sym in df.iloc[:, 0].tolist()]
            print(f"✅ CSV से {len(symbols)} stocks download हुए\n")
            return symbols
        except:
            pass
        
    except Exception as e:
        print(f"⚠️ Download failed: {e}")
    
    # Fallback: Default list
    print("⚠️ Web से download नहीं हो सका, default list use हो रही है\n")
    return get_default_list()

def get_default_list():
    """Default Nifty 500 list (अगर internet नहीं है)"""
    return [
        "RELIANCE.NS", "TCS.NS", "HDFCBANK.NS", "INFY.NS", "ICICIBANK.NS", "SBIN.NS",
        "KOTAKBANK.NS", "ITC.NS", "LTIM.NS", "AXISBANK.NS", "MARUTI.NS", "SUNPHARMA.NS",
        "WIPRO.NS", "HINDUNILVR.NS", "BAJAJ-AUTO.NS", "ASIANPAINT.NS", "ULTRACEMCO.NS",
        "TITAN.NS", "NESTLEIND.NS", "POWERGRID.NS", "LT.NS", "BHARTIARTL.NS", "JSWSTEEL.NS",
        "HINDALCO.NS", "DRREDDY.NS", "TECHM.NS", "HCLTECH.NS", "BAJAJFINSV.NS", "INDIGO.NS",
        "BRITANNIA.NS", "EICHERMOT.NS", "GAIL.NS", "GRASIM.NS", "HEROMOTOCO.NS", "ONGC.NS",
        "SBILIFE.NS", "SHREECEM.NS", "TATAMOTORS.NS", "TATAPOWER.NS", "TATASTEEL.NS",
        "TORNTPHARMA.NS", "UPL.NS", "ADANIENT.NS", "ADANIGREEN.NS", "ADANIPORTS.NS",
        "APOLLOHOSP.NS", "AUBANK.NS", "AUROPHARMA.NS", "BANKINDIA.NS", "BANKBARODA.NS",
        "BERGEPAINT.NS", "BHARATGEAR.NS", "BHEL.NS", "BIOCON.NS", "BOSCHLTD.NS", "BPCL.NS",
        "BSOFT.NS", "CADILAHC.NS", "CAMS.NS", "CANBK.NS", "CASTROLIND.NS", "CDSL.NS",
        "CGPOWER.NS", "CHOLAFIN.NS", "CIPLA.NS", "COALINDIA.NS", "COLPAL.NS", "CONTAINER.NS",
        "COROMANDEL.NS", "CREDITACC.NS", "CUMMINSIND.NS", "CYBERTECH.NS", "DABUR.NS",
        "DALSTKSEQ.NS", "DEEPAKFERT.NS", "DEEPAKNI.NS", "DELHIVERY.NS", "DELTACORP.NS",
        "DHANI.NS", "DMART.NS", "DRREDDY.NS", "DSHBANK.NS", "DUNLIRUBIND.NS", "DUVALPROD.NS",
        "DWDSILICA.NS", "EICHERMOT.NS", "EIDPARRY.NS", "EIMCOELECO.NS", "ELANOREMLP.NS",
        "ELECTRICSUP.NS", "ELEMENT.NS", "EMSECURE.NS", "ENDURANCE.NS", "ENGGCONF.NS",
        "ENGINERSIN.NS", "EQUITYBRIM.NS", "ERIS.NS", "ESAFMFBANK.NS", "ESCORT.NS",
        "ESSELPROP.NS", "ESTCEM.NS", "ETHICSEED.NS", "EVEREADY.NS", "EXIDEIND.NS"
    ]

# ========================
# Main Code
# ========================
def rsi(prices, period=14):
    if len(prices) < period + 1:
        return np.nan
    delta = np.diff(prices)
    gain = np.where(delta > 0, delta, 0)
    loss = np.where(delta < 0, -delta, 0)
    avg_gain = np.mean(gain[-period:])
    avg_loss = np.mean(loss[-period:])
    return round(100 - 100/(1 + avg_gain/avg_loss) if avg_loss else 100, 2)

def sma(prices, period):
    return round(np.mean(prices[-period:]) if len(prices) >= period else np.nan, 2)

def get_stock_data(symbol):
    try:
        data = yf.download(symbol, start=datetime.now()-timedelta(250), progress=False, timeout=5)
        if len(data) < 50:
            return None
        close = data['Close'].values
        return {
            "Symbol": symbol.replace(".NS", ""),
            "CMP": round(close[-1], 2),
            "RSI": rsi(close, 14),
            "DMA 30": sma(close, 30),
            "DMA 50": sma(close, 50),
            "DMA 200": sma(close, 200)
        }
    except:
        return None

# Main execution
print("="*70)
print("📈 NIFTY 500 STOCK ANALYZER")
print("="*70 + "\n")

# NSE से stocks download करें
symbols = download_nifty_500()

print(f"📊 कुल {len(symbols)} stocks analyze किए जा रहे हैं...\n")

# Analyze करें
results = []
for i, symbol in enumerate(symbols, 1):
    try:
        print(f"[{i}/{len(symbols)}] {symbol:<15}", end=" ", flush=True)
        stock = get_stock_data(symbol)
        if stock:
            results.append(stock)
            print("✓")
        else:
            print("✗")
    except:
        print("✗")

# DataFrame बनाएं
df = pd.DataFrame(results)
df = df.sort_values('CMP', ascending=False).reset_index(drop=True)

# Print करें
print("\n" + "="*70)
print("📊 NIFTY 500 DATA TABLE")
print("="*70 + "\n")
print(df.to_string(index=False))

# Statistics
print("\n" + "="*70)
print("📈 STATISTICS")
print("="*70)
print(f"✓ Total Stocks: {len(df)}")
print(f"✓ Highest Price: ₹{df['CMP'].max()}")
print(f"✓ Lowest Price: ₹{df['CMP'].min()}")
print(f"✓ Average Price: ₹{df['CMP'].mean():.2f}")
print(f"✓ Average RSI: {df['RSI'].mean():.2f}")

# Top 10 by RSI
print("\n" + "="*70)
print("🏆 TOP 10 BY RSI (High Momentum)")
print("="*70 + "\n")
print(df.nlargest(10, 'RSI')[['Symbol', 'CMP', 'RSI']].to_string(index=False))

# Excel में save करें
filename = f"Nifty500_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
df.to_excel(filename, index=False)
print(f"\n✅ Excel file saved: {filename}\n")
print("="*70)
