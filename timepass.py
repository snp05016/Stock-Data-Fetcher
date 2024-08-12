import pandas as pd
import yfinance as yf
from fpdf import FPDF

# Load the Excel file
df_xlsx = pd.read_excel("Book2.xlsx")

# Filter buy and sell data
data_for_buy = df_xlsx[df_xlsx['TYPE'] == 'Buy']
data_for_sell = df_xlsx[df_xlsx['TYPE'] == 'Sell']

# Summarize the quantities by ticker
buy_summary = data_for_buy.groupby('EXCHANGE CODE: TICKER')['UNITS'].sum().to_dict()
sell_summary = data_for_sell.groupby('EXCHANGE CODE: TICKER')['UNITS'].sum().to_dict()
buy_price_dict = data_for_buy.groupby('EXCHANGE CODE: TICKER')['PRICE'].sum().to_dict()

# Function to fetch the current stock price using yfinance for Indian stocks
def get_stock_price(ticker):
    try:
        stock = yf.Ticker(f"{ticker[4:]}.NS")  # Use '.BO' for BSE
        price = stock.history(period="1d")['Close'].iloc[-1]  # Get the latest closing price
        return round(float(price), 2)
    except Exception as e:
        print(f"Error fetching price for {ticker[4:]}: {e}")
        return None

# Calculate the remaining stock and get the current price
result = {}
for key in sell_summary:
    buy_quantity = buy_summary.get(key, 0)
    sell_quantity = sell_summary[key]
    remaining_quantity = buy_quantity - sell_quantity
    buy_price = buy_price_dict.get(key, 0)
    current_price = get_stock_price(key)
    
    if current_price is not None:
        unrealized_profit = (current_price - buy_price) * remaining_quantity
        current_value = round(remaining_quantity * current_price, 2)
        total_cost_of_remaining_shares = buy_price * remaining_quantity
        profit_percentage = (unrealized_profit / total_cost_of_remaining_shares) * 100 if total_cost_of_remaining_shares > 0 else 0
        
        result[key] = [
            round(remaining_quantity, 2),  # Remaining quantity
            current_price,  # Current price of the stock
            round(unrealized_profit, 2),  # Unrealized profit
            current_value,  # Current value of the portfolio/investment
            round(profit_percentage, 2)  # Profit percentage
        ]

# Create a PDF report
pdf = FPDF()
pdf.set_auto_page_break(auto=True, margin=15)
pdf.add_page()

# Set title
pdf.set_font("Arial", "B", 16)
pdf.cell(200, 10, txt="Stock Portfolio Summary", ln=True, align="C")

# Set column headers
pdf.set_font("Arial", "B", 12)
pdf.cell(40, 10, txt="Stock Name", border=1)
pdf.cell(40, 10, txt="Total Quantities", border=1)
pdf.cell(30, 10, txt="Current Price", border=1)
pdf.cell(30, 10, txt="Profit", border=1)
pdf.cell(30, 10, txt="Profit %", border=1)
pdf.cell(40, 10, txt="Current Value", border=1)
pdf.ln()

# Add data rows
pdf.set_font("Arial", "", 12)
for key, value in result.items():
    pdf.cell(40, 10, txt=key[4:], border=1)  # Stock name without the exchange prefix
    pdf.cell(40, 10, txt=str(value[0]), border=1)  # Total quantities
    pdf.cell(30, 10, txt=str(value[1]), border=1)  # Current price
    pdf.cell(30, 10, txt=str(value[2]), border=1)  # Profit
    pdf.cell(30, 10, txt=str(value[4]) + '%', border=1)  # Profit percentage
    pdf.cell(40, 10, txt=str(value[3]), border=1)  # Current value
    pdf.ln()

# Save the PDF to a file
pdf.output("stock_portfolio_summary.pdf")

print("PDF created successfully.")
