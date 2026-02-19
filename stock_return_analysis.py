import pandas as pd
import yfinance as yf
import numpy as np

# ==========================================
# STEP 1: DOWNLOAD DATA
# ==========================================
stock_ticker = "NVDA"
benchmark_ticker = "^GSPC"

print(f"⏳ Downloading 5 years of data for {stock_ticker} and {benchmark_ticker}...")
data = yf.download([stock_ticker, benchmark_ticker], period="5y", auto_adjust=True)

# Create the raw price dataset. 
# We DO NOT use dropna() here so we don't accidentally delete Day 0!
prices = pd.DataFrame({
    'NVDA_Price': data['Close'][stock_ticker],
    'Market_Price': data['Close'][benchmark_ticker]
})

# Remove timezone data from the dates so Excel doesn't crash
prices.index = prices.index.tz_localize(None)

# ==========================================
# STEP 2: CALCULATE BASIC METRICS (FIXED MATH)
# ==========================================
# 1. Daily Returns (We fill the first blank day with 0 instead of deleting it)
prices['NVDA_Return'] = prices['NVDA_Price'].pct_change().fillna(0)
prices['Market_Return'] = prices['Market_Price'].pct_change().fillna(0)

# 2. Growth of $10,000
prices['NVDA_10k_Growth'] = (1 + prices['NVDA_Return']).cumprod() * 10000
prices['Market_10k_Growth'] = (1 + prices['Market_Return']).cumprod() * 10000

# ==========================================
# STEP 3: PRINT THE LINKEDIN SUMMARY
# ==========================================
# Calculate Total Return 
stock_total_return = (prices['NVDA_Price'].iloc[-1] / prices['NVDA_Price'].iloc[0] - 1) * 100
market_total_return = (prices['Market_Price'].iloc[-1] / prices['Market_Price'].iloc[0] - 1) * 100

# Calculate Volatility (ignoring the Day 0 zero for accurate standard deviation)
stock_vol = prices['NVDA_Return'].iloc[1:].std() * np.sqrt(252) * 100
market_vol = prices['Market_Return'].iloc[1:].std() * np.sqrt(252) * 100

print("\n--- 5-YEAR PERFORMANCE SUMMARY ---")
print(f"📈 NVDA Return: {stock_total_return:,.2f}% | Volatility: {stock_vol:.2f}%")
print(f"📈 Benchmark Return: {market_total_return:,.2f}% | Volatility: {market_vol:.2f}%")

# ==========================================
# STEP 4: EXPORT TO EXCEL AND DRAW THE CHART
# ==========================================
excel_filename = "Project1_NVDA_Analysis.xlsx"
print(f"\n💾 Generating {excel_filename} with native Excel Chart...")

# Open the Excel Writer using the 'xlsxwriter' engine
writer = pd.ExcelWriter(excel_filename, engine='xlsxwriter')
prices.to_excel(writer, sheet_name='Data')

# Access the workbook and worksheet objects
workbook = writer.book
worksheet = writer.sheets['Data']

# Create a Line Chart object inside Excel
chart = workbook.add_chart({'type': 'line'})

# Get the number of rows so the chart knows exactly how much data to read
max_row = len(prices)

# Add the NVDA line to the chart (Column F is index 5)
chart.add_series({
    'name':       ['Data', 0, 5],
    'categories': ['Data', 1, 0, max_row, 0],
    'values':     ['Data', 1, 5, max_row, 5],
    'line':       {'color': 'green', 'width': 2}
})

# Add the Market line to the chart (Column G is index 6)
chart.add_series({
    'name':       ['Data', 0, 6],
    'categories': ['Data', 1, 0, max_row, 0],
    'values':     ['Data', 1, 6, max_row, 6],
    'line':       {'color': 'gray', 'dash_type': 'dash', 'width': 2}
})

# Format the Chart's title and axes
chart.set_title({'name': 'Growth of $10,000: NVDA vs S&P 500'})
chart.set_x_axis({'name': 'Date'})
chart.set_y_axis({'name': 'Portfolio Value ($)'})
chart.set_size({'width': 800, 'height': 400}) 

# Insert the chart into the Excel sheet at cell I2
worksheet.insert_chart('I2', chart)

# Save and close the Excel file
writer.close()

print("✅ Complete! Open the Excel file to see your chart and matching math.")
