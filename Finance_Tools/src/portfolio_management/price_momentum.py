# Standard Imports
import xlsxwriter
import pandas as pd
from scipy import stats
from statistics import mean

# Constants
time_periods = ['One-Year', 'Six-Month', 'Three-Month', 'One-Month']

# Functions
def extract_attributes(Stock_Dataframe):
    """
    Extracts attributes from stock data.

    Parameters:
    Dataframe (DataFrame): DataFrame containing stock data.

    Returns:
    DataFrame: DataFrame containing extracted attributes.
    """
    Dataframe_columns = ['Ticker', 'Price', 'Number of Shares to Buy', 'One-Year Price Return', 'One-Year Return Percentile', 'Six-Month Price Return', 'Six-Month Return Percentile', 'Three-Month Price Return', 'Three-Month Return Percentile', 'One-Month Price Return', 'One-Month Return Percentile', 'HQM Score']
    Dataframe = pd.DataFrame(columns=Dataframe_columns)
    for row in Stock_Dataframe.index:
        Dataframe = Dataframe.append(
            pd.Series(
                [
                    Stock_Dataframe.loc[row, 'Ticker'],
                    Stock_Dataframe.loc[row, 'Price'],
                    'N/A',
                    Stock_Dataframe.loc[row, 'One-Year Price Return'],
                    'N/A',
                    Stock_Dataframe.loc[row, 'Six-Month Price Return'],
                    'N/A',
                    Stock_Dataframe.loc[row, 'Three-Month Price Return'],
                    'N/A',
                    Stock_Dataframe.loc[row, 'One-Month Price Return'],
                    'N/A',
                    'N/A'
                ],
                index=Dataframe_columns
            ),
            ignore_index=True
        )
    return Dataframe

def calculate_return_percentile(Dataframe):
    """
    Calculates return percentiles for each time period.

    Parameters:
    Dataframe (DataFrame): DataFrame containing stock data.

    Returns:
    DataFrame: DataFrame with return percentiles calculated.
    """
    for row in Dataframe.index:
        for time_period in time_periods:
            Dataframe.loc[row, f'{time_period} Return Percentile'] = stats.percentileofscore(Dataframe[f'{time_period} Price Return'], Dataframe.loc[row, f'{time_period} Price Return']) / 100
    return Dataframe

def calculate_hqm_score(Dataframe):
    """
    Calculates the High-Quality Momentum (HQM) score for each stock.

    Parameters:
    Dataframe (DataFrame): DataFrame containing stock data.

    Returns:
    DataFrame: DataFrame with HQM score calculated.
    """
    for row in Dataframe.index:
        momentum_percentiles = []
        for time_period in time_periods:
            momentum_percentiles.append(Dataframe.loc[row, f'{time_period} Return Percentile'])
        Dataframe.loc[row, 'HQM Score'] = mean(momentum_percentiles)
    return Dataframe

def get_top_momentum_stocks(Dataframe):
    """
    Retrieves the top momentum stocks based on their HQM score.

    Parameters:
    Dataframe (DataFrame): DataFrame containing stock data.

    Returns:
    DataFrame: DataFrame containing top momentum stocks.
    """
    Dataframe.sort_values('HQM Score', ascending=False, inplace=True)
    Dataframe = Dataframe[:50]
    Dataframe.reset_index(drop=True, inplace=True)
    return Dataframe

def get_high_quality_momentum_stocks(stock_data):
    """
    Retrieves high-quality momentum stocks.

    Parameters:
    stock_data (list): List of dictionaries containing stock data.

    Returns:
    DataFrame: DataFrame containing high-quality momentum stocks.
    """
    # Extract attributes from stock data
    extracted_data = extract_attributes(stock_data)
    
    # Calculate return percentiles
    with_percentiles = calculate_return_percentile(extracted_data)
    
    # Calculate HQM scores
    with_scores = calculate_hqm_score(with_percentiles)
    
    # Filter top momentum stocks
    top_momentum_stocks = get_top_momentum_stocks(with_scores)
    
    return top_momentum_stocks

def save_recommended_trades(Dataframe):
    """
    Saves the recommended trades to a file.

    Parameters:
    Dataframe (DataFrame): DataFrame containing recommended trades.
    """
    file_path = 'data/output_data/recommended_trades.xlsx'
    
    writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
    Dataframe.to_excel(writer, sheet_name='price_momentum', index=False)

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
        'D': ['One-Year Price Return', percent_format],
        'E': ['One-Year Return Percentile', percent_format],
        'F': ['Six-Month Price Return', percent_format],
        'G': ['Six-Month Return Percentile', percent_format],
        'H': ['Three-Month Price Return', percent_format],
        'I': ['Three-Month Return Percentile', percent_format],
        'J': ['One-Month Price Return', percent_format],
        'K': ['One-Month Return Percentile', percent_format],
        'L': ['HQM Score', percent_format]
    }

    for column, format in column_formats.items():
        writer.sheets['price_momentum'].set_column(f'{column}:{column}', 25, format[1])
        writer.sheets['price_momentum'].write(f'{column}1', format[0], format[1])

    writer.save()