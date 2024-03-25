# Standard Imports
import xlsxwriter
import pandas as pd
from scipy import stats
from statistics import mean

# Constants
metrics = {
    'Price-to-Earnings Ratio': 'PE Percentile',
    'Price-to-Book Ratio': 'PB Percentile',
    'Price-to-Sales Ratio': 'PS Percentile',
    'EV/EBITDA': 'EV/EBITDA Percentile',
    'EV/RE': 'EV/RE Percentile'
}

# Functions

def extract_attributes(stock_data):
    """
    Extracts attributes from stock data.

    Parameters:
    stock_data (DataFrame): DataFrame containing stock data.

    Returns:
    DataFrame: DataFrame containing extracted attributes.
    """

    stock_data_columns = ['Ticker', 'Price', 'Number of Shares to Buy', 'Price-to-Earnings Ratio', 'PE Percentile', 'Price-to-Book Ratio', 'PB Percentile', 'Price-to-Sales Ratio', 'PS Percentile', 'EV/EBITDA', 'EV/EBITDA Percentile', 'EV/RE', 'EV/RE Percentile', 'RV Score']
    stock_attributes = pd.DataFrame(columns=stock_data_columns)
    for row in stock_data.index:
        stock_attributes = stock_attributes.append(
            pd.Series(
                [
                    stock_data.loc[row, 'Ticker'],
                    stock_data.loc[row, 'Price'],
                    'N/A',
                    stock_data.loc[row, 'Price-to-Earnings Ratio'],
                    'N/A',
                    stock_data.loc[row, 'Price-to-Book Ratio'],
                    'N/A',
                    stock_data.loc[row, 'Price-to-Sales Ratio'],
                    'N/A',
                    stock_data.loc[row, 'EV/EBITDA'],
                    'N/A',
                    stock_data.loc[row, 'EV/RE'],
                    'N/A',
                    'N/A'
                ],
                index=stock_data_columns
            ),
            ignore_index=True
        )

    return stock_attributes

def calculate_ratios_percentile(stock_ratios):
    """
    Calculates the percentile of ratios for each stock.

    Parameters:
    stock_ratios (DataFrame): DataFrame containing stock ratios.

    Returns:
    DataFrame: DataFrame containing stock ratios and their percentiles.
    """

    for metric, percentile_column in metrics.items():
        stock_ratios[percentile_column] = stock_ratios[metric].apply(lambda x: stats.percentileofscore(stock_ratios[metric], x) / 100)
    
    return stock_ratios

def calculate_rv_score(stock_data):
    """
    Calculates the Robust Value (RV) score for each stock.

    Parameters:
    stock_data (DataFrame): DataFrame containing stock data.

    Returns:
    DataFrame: DataFrame containing stock data with RV scores.
    """

    for row in stock_data.index:
        rv_percentiles = []
        for metric in metrics.keys():
            rv_percentiles.append(stock_data.loc[row, metrics[metric]])
        stock_data.loc[row, 'RV Score'] = mean(rv_percentiles)
    
    return stock_data

def top_rv_stocks(stock_data, number_of_stocks):
    """
    Returns the top n stocks based on the RV score.

    Parameters:
    stock_data (DataFrame): DataFrame containing stock data.
    number_of_stocks (int): Number of top stocks to return.

    Returns:
    DataFrame: DataFrame containing top n stocks based on the RV score.
    """

    stock_data.sort_values('RV Score', ascending=False, inplace=True)
    stock_data = stock_data[:number_of_stocks]
    return stock_data

def calculate_percentile_RV(stock_ratios):
    """
    Calculates the percentile of ratios for each stock and the Robust Value (RV) score.

    Parameters:
    stock_ratios (DataFrame): DataFrame containing stock ratios.

    Returns:
    DataFrame: DataFrame containing stock ratios and their percentiles and RV scores.
    """

    stock_ratios = extract_attributes(stock_ratios)
    stock_ratios = calculate_ratios_percentile(stock_ratios)
    stock_ratios = calculate_rv_score(stock_ratios)
    stock_ratios = top_rv_stocks(stock_ratios, 50)
    return stock_ratios

def save_recommended_trades(ratio_data):
    """
    Saves the recommended trades to a file.

    Parameters:
    ratio_data (DataFrame): DataFrame containing recommended trades.
    """
    file_path = 'data/output_data/recommended_trades.xlsx'
    writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
    ratio_data.to_excel(writer, sheet_name='ratio_analysis', index=False)

    background_color = '#0a0a23'
    font_color = '#ffffff'

    string_format = writer.book.add_format(
        {
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

    dollar_format = writer.book.add_format(
        {
            'num_format':'$0.00',
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

    integer_format = writer.book.add_format(
        {
            'num_format':'0',
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

    float_format = writer.book.add_format(
        {
            'num_format':'0.00',
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

    percent_format = writer.book.add_format(
        {
            'num_format':'0.0%',
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

    column_formats = {
        'A': ['Ticker', string_format],
        'B': ['Price', dollar_format],
        'C': ['Number of Shares to Buy', integer_format],
        'D': ['Price-to-Earnings Ratio', float_format],
        'E': ['PE Percentile', percent_format],
        'F': ['Price-to-Book Ratio', float_format],
        'G': ['PB Percentile', percent_format],
        'H': ['Price-to-Sales Ratio', float_format],
        'I': ['PS Percentile', percent_format],
        'J': ['EV/EBITDA', float_format],
        'K': ['EV/EBITDA Percentile', percent_format],
        'L': ['EV/RE', float_format],
        'M': ['EV/RE Percentile', percent_format],
        'N': ['RV Score', float_format]
    }

    for column, format in column_formats.items():
        writer.sheets['ratio_analysis'].set_column(f'{column}:{column}', 25, format[1])
        writer.sheets['ratio_analysis'].write(f'{column}1', format[0], format[1])

    writer.save()

    return None

