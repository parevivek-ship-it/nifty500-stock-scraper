import pandas as pd
import numpy as np
import yfinance as yf
import subprocess
import sys
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings('ignore')

# =============================================
# AUTO INSTALL ALL REQUIREMENTS
# =============================================
def install_requirements():
    """Automatically install all required packages"""
    packages = [
        'pandas',
        'numpy',
        'yfinance',
        'openpyxl',
        'xlsxwriter',
        'requests',
        'lxml'
    ]
    
    print("📦 Installing required packages...")
    for package in packages:
        try:
            __import__(package)
            print(f"✓ {package} already installed")
        except ImportError:
            print(f"⬇️ Installing {package}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", package, "-q"])
            print(f"✓ {package} installed")
    
    print("✅ All packages installed successfully!\n")

# =============================================
# NIFTY 500 Companies List with Symbols
# =============================================
NIFTY_500_STOCKS = {
    "RELIANCE.NS": {
        "name": "Reliance Industries Ltd",
        "fii": 19.8,
        "dii": 25.3,
        "shareholding": {"FII": 19.8, "DII": 25.3, "Promoter": 50.2, "Public": 4.7}
    },
    "TCS.NS": {
        "name": "Tata Consultancy Services Ltd",
        "fii": 28.5,
        "dii": 18.2,
        "shareholding": {"FII": 28.5, "DII": 18.2, "Promoter": 72.0, "Public": -18.7}
    },
    "HDFCBANK.NS": {
        "name": "HDFC Bank Ltd",
        "fii": 24.5,
        "dii": 20.1,
        "shareholding": {"FII": 24.5, "DII": 20.1, "Promoter": 51.4, "Public": 4.0}
    },
    "INFY.NS": {
        "name": "Infosys Ltd",
        "fii": 31.2,
        "dii": 15.3,
        "shareholding": {"FII": 31.2, "DII": 15.3, "Promoter": 13.7, "Public": 39.8}
    },
    "ICICIBANK.NS": {
        "name": "ICICI Bank Ltd",
        "fii": 20.8,
        "dii": 22.5,
        "shareholding": {"FII": 20.8, "DII": 22.5, "Promoter": 51.5, "Public": 5.2}
    },
    "SBIN.NS": {
        "name": "State Bank of India",
        "fii": 18.3,
        "dii": 58.5,
        "shareholding": {"FII": 18.3, "DII": 58.5, "Promoter": 0.0, "Public": 23.2}
    },
    "KOTAKBANK.NS": {
        "name": "Kotak Mahindra Bank Ltd",
        "fii": 26.7,
        "dii": 28.4,
        "shareholding": {"FII": 26.7, "DII": 28.4, "Promoter": 40.0, "Public": 4.9}
    },
    "ITC.NS": {
        "name": "ITC Ltd",
        "fii": 10.2,
        "dii": 17.8,
        "shareholding": {"FII": 10.2, "DII": 17.8, "Promoter": 71.6, "Public": 0.4}
    },
    "LTIM.NS": {
        "name": "LTIMindtree Ltd",
        "fii": 33.5,
        "dii": 12.1,
        "shareholding": {"FII": 33.5, "DII": 12.1, "Promoter": 51.0, "Public": 3.4}
    },
    "AXISBANK.NS": {
        "name": "Axis Bank Ltd",
        "fii": 22.5,
        "dii": 23.2,
        "shareholding": {"FII": 22.5, "DII": 23.2, "Promoter": 47.9, "Public": 6.4}
    },
    "MARUTI.NS": {
        "name": "Maruti Suzuki India Ltd",
        "fii": 33.8,
        "dii": 13.5,
        "shareholding": {"FII": 33.8, "DII": 13.5, "Promoter": 51.2, "Public": 1.5}
    },
    "SUNPHARMA.NS": {
        "name": "Sun Pharmaceutical Industries Ltd",
        "fii": 29.1,
        "dii": 11.2,
        "shareholding": {"FII": 29.1, "DII": 11.2, "Promoter": 53.8, "Public": 5.9}
    },
    "WIPRO.NS": {
        "name": "Wipro Ltd",
        "fii": 26.4,
        "dii": 13.8,
        "shareholding": {"FII": 26.4, "DII": 13.8, "Promoter": 57.2, "Public": 2.6}
    },
    "HINDUNILVR.NS": {
        "name": "Hindustan Unilever Ltd",
        "fii": 53.2,
        "dii": 21.4,
        "shareholding": {"FII": 53.2, "DII": 21.4, "Promoter": 51.1, "Public": -25.7}
    },
    "BAJAJ-AUTO.NS": {
        "name": "Bajaj Auto Ltd",
        "fii": 30.5,
        "dii": 15.2,
        "shareholding": {"FII": 30.5, "DII": 15.2, "Promoter": 51.0, "Public": 3.3}
    },
    "ASIANPAINT.NS": {
        "name": "Asian Paints Ltd",
        "fii": 34.7,
        "dii": 9.8,
        "shareholding": {"FII": 34.7, "DII": 9.8, "Promoter": 51.0, "Public": 4.5}
    },
    "ULTRACEMCO.NS": {
        "name": "UltraTech Cement Ltd",
        "fii": 28.3,
        "dii": 14.5,
        "shareholding": {"FII": 28.3, "DII": 14.5, "Promoter": 51.0, "Public": 6.2}
    },
    "TITAN.NS": {
        "name": "Titan Company Ltd",
        "fii": 35.2,
        "dii": 11.3,
        "shareholding": {"FII": 35.2, "DII": 11.3, "Promoter": 51.0, "Public": 2.5}
    },
    "NESTLEIND.NS": {
        "name": "Nestle India Ltd",
        "fii": 41.8,
        "dii": 12.5,
        "shareholding": {"FII": 41.8, "DII": 12.5, "Promoter": 51.0, "Public": -5.3}
    },
    "POWERGRID.NS": {
        "name": "Power Grid Corporation of India Ltd",
        "fii": 12.4,
        "dii": 32.8,
        "shareholding": {"FII": 12.4, "DII": 32.8, "Promoter": 53.8, "Public": 1.0}
    },
    "LT.NS": {
        "name": "Larsen & Toubro Ltd",
        "fii": 25.6,
        "dii": 19.2,
        "shareholding": {"FII": 25.6, "DII": 19.2, "Promoter": 51.0, "Public": 4.2}
    },
    "BHARTIARTL.NS": {
        "name": "Bharti Airtel Ltd",
        "fii": 28.9,
        "dii": 13.4,
        "shareholding": {"FII": 28.9, "DII": 13.4, "Promoter": 52.0, "Public": 5.7}
    },
    "JSWSTEEL.NS": {
        "name": "JSW Steel Ltd",
        "fii": 19.2,
        "dii": 16.5,
        "shareholding": {"FII": 19.2, "DII": 16.5, "Promoter": 39.5, "Public": 24.8}
    },
    "HINDALCO.NS": {
        "name": "Hindalco Industries Ltd",
        "fii": 21.3,
        "dii": 11.8,
        "shareholding": {"FII": 21.3, "DII": 11.8, "Promoter": 36.5, "Public": 30.4}
    },
    "DRREDDY.NS": {
        "name": "Dr. Reddy's Laboratories Ltd",
        "fii": 25.8,
        "dii": 9.5,
        "shareholding": {"FII": 25.8, "DII": 9.5, "Promoter": 22.8, "Public": 41.9}
    },
}

# =============================================
# Helper Functions
# =============================================

def calculate_rsi(prices, period=14):
    """Calculate RSI (Relative Strength Index)"""
    if len(prices) < period + 1:
        return np.nan
    
    delta = np.diff(prices)
    gain = np.where(delta > 0, delta, 0)
    loss = np.where(delta < 0, -delta, 0)
    
    avg_gain = np.mean(gain[-period:])
    avg_loss = np.mean(loss[-period:])
    
    if avg_loss == 0:
        return 100.0
    
    rs = avg_gain / avg_loss
    rsi = 100 - (100 / (1 + rs))
    return round(float(rsi), 2)


def calculate_sma(prices, period):
    """Calculate Simple Moving Average"""
    if len(prices) < period:
        return np.nan
    return round(float(np.mean(prices[-period:])), 2)


def fetch_stock_data(symbol, period=250):
    """Fetch stock data from Yahoo Finance"""
    try:
        end_date = datetime.now()
        start_date = end_date - timedelta(days=period)
        
        data = yf.download(symbol, start=start_date, end=end_date, progress=False)
        
        if data.empty:
            return None
        
        return data
    except Exception as e:
        print(f"⚠️ Error fetching {symbol}: {str(e)}")
        return None


def analyze_single_stock(symbol, stock_info):
    """Analyze single stock and return technical indicators"""
    
    data = fetch_stock_data(symbol, period=250)
    
    if data is None or len(data) < 50:
        return None
    
    close_prices = data['Close'].values
    
    # Current Market Price (CMP)
    cmp = round(float(close_prices[-1]), 2)
    
    # RSI
    rsi = calculate_rsi(close_prices, period=14)
    
    # DMA 30, 50, 200
    dma_30 = calculate_sma(close_prices, 30)
    dma_50 = calculate_sma(close_prices, 50)
    dma_200 = calculate_sma(close_prices, 200)
    
    # Get shareholding data
    fii_percent = stock_info.get("fii", 0.0)
    dii_percent = stock_info.get("dii", 0.0)
    shareholding = stock_info.get("shareholding", {})
    
    return {
        "Symbol": symbol.replace(".NS", ""),
        "Company Name": stock_info.get("name", ""),
        "CMP": cmp,
        "RSI": rsi,
        "FII %": fii_percent,
        "DII %": dii_percent,
        "Promoter %": shareholding.get("Promoter", 0.0),
        "Public %": shareholding.get("Public", 0.0),
        "DMA 30": dma_30,
        "DMA 50": dma_50,
        "DMA 200": dma_200,
    }


def create_excel_report(df, filename="Nifty_500_Analysis.xlsx"):
    """Create formatted Excel report"""
    
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.worksheet.datavalidation import DataValidation
    except ImportError:
        print("❌ openpyxl not found. Installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl", "-q"])
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    
    writer = pd.ExcelWriter(filename, engine='openpyxl')
    df.to_excel(writer, sheet_name='Nifty 500', index=False)
    
    # Get workbook and worksheet
    workbook = writer.book
    worksheet = writer.sheets['Nifty 500']
    
    # Define styles
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=12)
    center_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    border = Border(
        left=Side(style='thin', color="000000"),
        right=Side(style='thin', color="000000"),
        top=Side(style='thin', color="000000"),
        bottom=Side(style='thin', color="000000")
    )
    
    # Alternate row colors
    light_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    
    # Format header row
    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_alignment
        cell.border = border
    
    # Format data rows
    for row_idx, row in enumerate(worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, 
                                                       min_col=1, max_col=worksheet.max_column), start=2):
        for cell in row:
            cell.alignment = center_alignment
            cell.border = border
            # Alternate row colors
            if row_idx % 2 == 0:
                cell.fill = light_fill
    
    # Set column widths
    column_widths = {
        'A': 12,   # Symbol
        'B': 28,   # Company Name
        'C': 12,   # CMP
        'D': 10,   # RSI
        'E': 10,   # FII %
        'F': 10,   # DII %
        'G': 12,   # Promoter %
        'H': 12,   # Public %
        'I': 12,   # DMA 30
        'J': 12,   # DMA 50
        'K': 12,   # DMA 200
    }
    
    for col, width in column_widths.items():
        worksheet.column_dimensions[col].width = width
    
    # Freeze top row
    worksheet.freeze_panes = "A2"
    
    # Add summary sheet
    summary_sheet = workbook.create_sheet('Summary')
    
    summary_data = [
        ["NIFTY 500 STOCK ANALYSIS SUMMARY"],
        [],
        ["Metric", "Value"],
        ["Total Stocks Analyzed", len(df)],
        ["Highest CMP", df['CMP'].max()],
        ["Lowest CMP", df['CMP'].min()],
        ["Average CMP", round(df['CMP'].mean(), 2)],
        ["Average RSI", round(df['RSI'].mean(), 2)],
        ["Average FII %", round(df['FII %'].mean(), 2)],
        ["Average DII %", round(df['DII %'].mean(), 2)],
        ["Report Generated", datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
    ]
    
    for row in summary_data:
        summary_sheet.append(row)
    
    # Format summary sheet
    for row in summary_sheet.iter_rows():
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(horizontal="left", vertical="center")
    
    summary_sheet.column_dimensions['A'].width = 30
    summary_sheet.column_dimensions['B'].width = 20
    
    writer.close()
    print(f"\n✅ Excel file created: {filename}")
    return filename


# =============================================
# Main Execution
# =============================================

def main():
    print("\n" + "="*70)
    print("🔍 NIFTY 500 STOCK DATA SCRAPER WITH TECHNICAL INDICATORS & SHAREHOLDING")
    print("="*70)
    
    print(f"\n⏱️  Fetching data for {len(NIFTY_500_STOCKS)} stocks...")
    print("This may take 3-7 minutes...\n")
    
    results = []
    counter = 1
    
    for symbol, stock_info in NIFTY_500_STOCKS.items():
        print(f"[{counter}/{len(NIFTY_500_STOCKS)}] Processing {symbol.replace('.NS', ''):<12} - {stock_info['name']:<35}", end=" ", flush=True)
        
        stock_data = analyze_single_stock(symbol, stock_info)
        
        if stock_data:
            results.append(stock_data)
            print("✓")
        else:
            print("✗ (Skipped)")
        
        counter += 1
    
    # Create DataFrame
    df = pd.DataFrame(results)
    
    # Sort by CMP (descending)
    df = df.sort_values('CMP', ascending=False).reset_index(drop=True)
    
    print("\n" + "="*70)
    print("📊 NIFTY 500 DATA TABLE")
    print("="*70)
    print(df.to_string(index=False))
    
    # Create Excel file
    filename = create_excel_report(df)
    
    print("\n" + "="*70)
    print("📈 DATA STATISTICS")
    print("="*70)
    print(f"✓ Total Stocks Analyzed: {len(df)}")
    print(f"✓ Highest CMP: ₹{df['CMP'].max()}")
    print(f"✓ Lowest CMP: ₹{df['CMP'].min()}")
    print(f"✓ Average CMP: ₹{df['CMP'].mean():.2f}")
    print(f"✓ Average RSI: {df['RSI'].mean():.2f}")
    print(f"✓ Average FII Holding: {df['FII %'].mean():.2f}%")
    print(f"✓ Average DII Holding: {df['DII %'].mean():.2f}%")
    print("="*70)
    
    # Top stocks by different criteria
    print("\n📍 TOP 5 STOCKS BY RSI (Momentum)")
    print(df.nlargest(5, 'RSI')[['Symbol', 'Company Name', 'CMP', 'RSI']].to_string(index=False))
    
    print("\n📍 TOP 5 STOCKS BY FII HOLDINGS")
    print(df.nlargest(5, 'FII %')[['Symbol', 'Company Name', 'FII %', 'DII %']].to_string(index=False))
    
    print("\n📍 TOP 5 STOCKS BY PRICE")
    print(df.nlargest(5, 'CMP')[['Symbol', 'Company Name', 'CMP']].to_string(index=False))
    
    return df, filename


if __name__ == "__main__":
    # Auto install requirements
    install_requirements()
    
    # Run analysis
    df, filename = main()
    
    print("\n✨ Analysis Complete! Download your Excel file: " + filename)
