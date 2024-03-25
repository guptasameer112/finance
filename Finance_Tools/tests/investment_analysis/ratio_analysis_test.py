
# # Fetch the stock data
# FUNCTION = 'TIME_SERIES_MONTHLY'
# stock_data_monthly = utils.get_stock_data(stock_tickers, BASE_API_URL, ALPHAVANTAGE_API_KEY, FUNCTION)

# # Obtain Price and monthly returns dataframe
# stock_price_data = utils.extract_last_closing_price(stock_data_monthly)

# # Fetch overview ratios of the company
# FUNCTION = 'OVERVIEW'
# stock_overview_data = utils.get_stock_data(stock_tickers, BASE_API_URL, ALPHAVANTAGE_API_KEY, FUNCTION)

# # Extract the ratios
# stock_ratios = utils.get_ratios(stock_overview_data, stock_price_data)

# # Calculate the Percentile of ratio for each stock
# stock_ratios = ratio_analysis.calculate_percentile_RV(stock_ratios)

# # Calculate the number of shares to buy
# ratio_data = utils.calculate_number_of_shares_to_buy(portfolio_amount, stock_ratios)

# # Save the recommended trades to excel
# ratio_analysis.save_recommended_trades(ratio_data)

# print("Recommended trades have been saved to 'recommended_trades.xlsx'.")