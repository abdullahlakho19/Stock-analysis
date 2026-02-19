# Stock-analysis
this is only for educational purpose

# Quantitative Finance : Financial Data & Excel Automation

## 📌 Overview
This project bridges the gap between Python-based data extraction and traditional financial reporting. It is a fully automated script that hits the Yahoo Finance API to pull historical market data, calculates core institutional risk/return metrics, and automatically generates a formatted Excel report complete with native, interactive charts.

This eliminates the need for manual data entry and demonstrates how Python can streamline daily reporting tasks for a portfolio management team.

## 🚀 Key Features
* **Automated Data Pipeline:** Extracts 5 years of clean, adjusted daily closing prices for an asset (e.g., NVDA) and a benchmark (e.g., S&P 500) using the `yfinance` API.
* **Financial Math Engine:** Calculates essential performance metrics using `pandas` and `numpy`:
  * **Holding Period Return (Total Return)**
  * **Annualized Volatility ($\sigma$)**
  * **Cumulative Returns (Growth of a $10,000 initial investment)**
* **Excel Reporting Automation:** Bypasses basic CSV exports by using the `xlsxwriter` engine to construct a fully formatted `.xlsx` file, injecting a native Excel Line Chart directly into the worksheet.

## 🛠️ Technology Stack
* **Python 3.x**
* **Pandas:** Time-series data manipulation and cleaning.
* **NumPy:** Vectorized financial math (standard deviation, square roots).
* **yFinance:** Market data API integration.
* **XlsxWriter:** Programmatic generation of Excel files and charts.

## 📦 Installation & Setup

1. Clone this repository to your local machine.
2. Ensure you have the required financial and data science libraries installed. You can install them via terminal:
   ```bash
   pip install pandas numpy yfinance xlsxwriter

   
