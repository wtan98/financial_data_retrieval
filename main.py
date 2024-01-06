import yfinance as yf
import pandas as pd


ticker_symbol = input("Give me ticker symbol of company: ")
company = yf.Ticker(ticker_symbol)


balance_sheet_data = [
    "Total Assets",
    "Total Liabilities Net Minority Interest",
    "Total Equity Gross Minority Interest",
    "Common Stock Equity",
    "Total Debt",
    "Cash And Cash Equivalents",
    "Working Capital",
    "Net Tangible Assets",
    "Total Capitalization",
]


income_statement_data = [
    "Total Revenue",
    "Net Income",
    "Gross Profit",
    "EBITDA",
    "Operating Income",
    "Interest Expense",
    "Tax Provision",
    "Diluted EPS",
]


cash_flow_statement_data = [
    "Operating Cash Flow",
    "Free Cash Flow",
    "Investing Cash Flow",
    "Financing Cash Flow",
    "Changes In Cash",
    "Repayment Of Debt",
    "Issuance Of Debt",
]


balance_sheet_df = company.balance_sheet.loc[balance_sheet_data]
income_statement_df = company.income_stmt.loc[income_statement_data]
cash_flow_statement_df = company.cashflow.loc[cash_flow_statement_data]
dividend_df = company.dividends


with pd.ExcelWriter(f"{ticker_symbol}_financial_statements.xlsx", engine="xlsxwriter") as writer:
    balance_sheet_df.to_excel(writer, sheet_name="Balance Sheet")
    income_statement_df.to_excel(writer, sheet_name="Income Statement")
    cash_flow_statement_df.to_excel(writer, sheet_name="Cash Flow Statement")
    dividend_df.to_excel(writer, sheet_name="Dividend History", index=False)


for news_item in (company.news):
    news_link = news_item.get("link", "")
    related_tickers = news_item.get("relatedTickers", [])


    print(f"Link: {news_link}")
    if related_tickers:
        print(f"   Related Tickers: {', '.join(related_tickers)}")