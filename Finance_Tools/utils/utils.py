import math
import json
import requests
import datetime
import pandas as pd
import xlsxwriter

def get_stock_tickers(file_path):
    """
    Retrieves stock tickers from a CSV file.

    Parameters:
    file_path (str): Path to the CSV file containing stock tickers.

    Returns:
    list: List of stock tickers.
    """
    
    return pd.read_csv(file_path)['Ticker'].tolist()

def get_stock_data(stock_tickers, base_api_url, api_key, function):
    """
    Retrieves stock data using the Alpha Vantage API.

    Parameters:
    stock_tickers (list): List of stock tickers.
    base_api_url (str): Base URL for the Alpha Vantage API.
    api_key (str): Alpha Vantage API key.

    Returns:
    list: List of dictionaries containing stock data.
    """

    all_stocks_data = []
    for stock_ticker in stock_tickers:
        params = {
            'function': function,
            'symbol': stock_ticker,
            'apikey': api_key
        }
        response = requests.get(base_api_url, params=params)
        data = json.loads(response.text)
        all_stocks_data.append(data)

    return all_stocks_data

def extract_last_closing_price(stock_data):
    """
    Extracts price data from the stock data.

    Parameters:
    stock_data (list): List of dictionaries containing stock data.

    Returns:
    DataFrame: DataFrame containing price data.
    """

    stock_prices = {}
    for data in stock_data:
        symbol = data.get("Meta Data")["2. Symbol"]
        last_refreshed = data.get("Meta Data")["3. Last Refreshed"]
        closing_price = data.get("Monthly Time Series")[last_refreshed]["4. close"]
        stock_prices[symbol] = closing_price

    return stock_prices
    
def calculate_monthly_return_percentage(stock_data, time_periods):
    """
    Calculates monthly returns for each stock.

    Parameters:
    stock_price_data (DataFrame): DataFrame containing stock price data.
    time_periods (list): List of time periods.

    Returns:
    DataFrame: DataFrame with monthly returns calculated.
    """

    # print(stock_data)
    stock_price_returns = {}
    for data in stock_data:
        symbol = data.get("Meta Data")["2. Symbol"]
        last_refreshed = data.get("Meta Data")["3. Last Refreshed"]
        last_closing_price = data.get("Monthly Time Series")[last_refreshed]["4. close"]
        monthly_series_keys = list(data.get("Monthly Time Series").keys())

        stock_price_returns[symbol] = {
            "Price": last_closing_price,
            time_periods[0]: (float(last_closing_price) - float(data.get("Monthly Time Series")[monthly_series_keys[11]]["4. close"])) / float(data.get("Monthly Time Series")[monthly_series_keys[11]]["4. close"]) * 100,
            time_periods[1]: (float(last_closing_price) - float(data.get("Monthly Time Series")[monthly_series_keys[5]]["4. close"])) / float(data.get("Monthly Time Series")[monthly_series_keys[5]]["4. close"]) * 100,
            time_periods[2]: (float(last_closing_price) - float(data.get("Monthly Time Series")[monthly_series_keys[2]]["4. close"])) / float(data.get("Monthly Time Series")[monthly_series_keys[2]]["4. close"]) * 100,
            time_periods[3]: (float(last_closing_price) - float(data.get("Monthly Time Series")[monthly_series_keys[0]]["4. close"])) / float(data.get("Monthly Time Series")[monthly_series_keys[0]]["4. close"]) * 100
        }

    stock_price_returns_dataframe = pd.DataFrame(columns = ["Ticker", "Price", "One-Year Price Return", "Six-Month Price Return", "Three-Month Price Return", "One-Month Price Return"])
    for symbol, stock in stock_price_returns.items():
        stock_price_returns_dataframe = stock_price_returns_dataframe.append(
            pd.Series(
                [
                    symbol,
                    stock['Price'],
                    stock[time_periods[0]],
                    stock[time_periods[1]],
                    stock[time_periods[2]],
                    stock[time_periods[3]]
                ],
                index=["Ticker", "Price", "One-Year Price Return", "Six-Month Price Return", "Three-Month Price Return", "One-Month Price Return"]
            ),
            ignore_index=True
        )

    return stock_price_returns_dataframe

def get_ratios(stock_data, price_data):
    """
    Extracts ratios from the stock data.

    Parameters:
    stock_data (list): List of dictionaries containing stock data.
    ratios (list): List of ratios to extract.

    Returns:
    DataFrame: DataFrame containing ratios.
    """

    stock_ratios = {}
    for data in stock_data:
        symbol = data.get("Symbol")
        price = price_data.get(symbol)
        pe_ratio = data.get("PERatio")
        pb_ratio = data.get("PriceToBookRatio")
        ps_ratio = data.get("PriceToSalesRatioTTM")
        ev_ebitda = data.get("EVToEBITDA")
        ev_revenue = data.get("EVToRevenue")

        stock_ratios[symbol] = {
            "Price": price,
            "Price-to-Earnings Ratio": pe_ratio,
            "Price-to-Book Ratio": pb_ratio,
            "Price-to-Sales Ratio": ps_ratio,
            "EV/EBITDA": ev_ebitda,
            "EV/RE": ev_revenue
        }

    stock_ratios_dataframe = pd.DataFrame(columns = ["Ticker", "Price", "Price-to-Earnings Ratio", "Price-to-Book Ratio", "Price-to-Sales Ratio", "EV/EBITDA", "EV/RE"])
    for symbol, stock in stock_ratios.items():
        stock_ratios_dataframe = stock_ratios_dataframe.append(
            pd.Series(
                [
                    symbol,
                    float(stock["Price"]),
                    float(stock["Price-to-Earnings Ratio"]),
                    float(stock["Price-to-Book Ratio"]),
                    float(stock["Price-to-Sales Ratio"]),
                    float(stock["EV/EBITDA"]),
                    float(stock["EV/RE"])
                ],
                index=["Ticker", "Price", "Price-to-Earnings Ratio", "Price-to-Book Ratio", "Price-to-Sales Ratio", "EV/EBITDA", "EV/RE"]
            ),
            ignore_index=True
        )

    return stock_ratios_dataframe

def calculate_number_of_shares_to_buy(portfolio_amount, stock_data):
    """
    Calculates the number of shares to buy based on the portfolio amount.

    Parameters:
    portfolio_amount (float): Amount to invest in the portfolio.
    stock_data (DataFrame): DataFrame containing stock data.

    Returns:
    DataFrame: DataFrame with number of shares to buy calculated.
    """

    for row in stock_data.index:
        stock_data.loc[row, 'Number of Shares to Buy'] = math.floor(portfolio_amount / float(stock_data.loc[row, 'Price']))

    return stock_data

# def save_to_excel(data, sheet_name, file_path, column_formats):
#     """
#     Saves the data to an Excel file.

#     Parameters:
#     data (DataFrame): Data to be saved.
#     file_path (str): Path to the Excel file.
#     """ 
#     writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
#     data.to_excel(writer, sheet_name=sheet_name, index=False)

#     background_color = '#0a0a23'
#     font_color = '#ffffff'

#     string_format = writer.book.add_format(
#         {
#             'font_color': font_color,
#             'bg_color': background_color,
#             'border': 1
#         }
#     )

#     dollar_format = writer.book.add_format(
#         {
#             'num_format':'$0.00',
#             'font_color': font_color,
#             'bg_color': background_color,
#             'border': 1
#         }
#     )

#     integer_format = writer.book.add_format(
#         {
#             'num_format':'0',
#             'font_color': font_color,
#             'bg_color': background_color,
#             'border': 1
#         }
#     )

#     float_format = writer.book.add_format(
#         {
#             'num_format':'0.00',
#             'font_color': font_color,
#             'bg_color': background_color,
#             'border': 1
#         }
#     )

#     percent_format = writer.book.add_format(
#         {
#             'num_format':'0.0%',
#             'font_color': font_color,
#             'bg_color': background_color,
#             'border': 1
#         }
#     )

#     column_formats = {
#         'A': ['Ticker', string_format],
#         'B': ['Price', dollar_format],
#         'C': ['Price-to-Earnings Ratio', float_format],
#         'D': ['Price-to-Book Ratio', float_format],
#         'E': ['Price-to-Sales Ratio', float_format],
#         'F': ['EV/EBITDA', float_format],
#         'G': ['EV/RE', float_format],
#         'H': ['Number of Shares to Buy', integer_format]
#     }

#     for column, format in column_formats.items():
#         writer.sheets[sheet_name].set_column(f'{column}:{column}', 20, format[1])
#         writer.sheets[sheet_name].write(f'{column}1', format[0], format[1])

#     writer.save()