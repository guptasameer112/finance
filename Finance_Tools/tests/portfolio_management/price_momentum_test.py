# # Fetch the stock data
# FUNCTION = 'TIME_SERIES_MONTHLY'
# stock_data_monthly = utils.get_stock_data(stock_tickers, BASE_API_URL, ALPHAVANTAGE_API_KEY, FUNCTION)

# # Obtain Price and monthly returns dataframe
# stock_price_data = utils.extract_last_closing_price(stock_data_monthly)

# # Obtain Price Returns
# time_periods = ['One-Year', 'Six-Month', 'Three-Month', 'One-Month']
# stock_price_return_data = utils.calculate_monthly_return_percentage(stock_data_monthly, time_periods)

# # Run the price momentum strategy
# price_momentum_data = price_momentum.get_high_quality_momentum_stocks(stock_price_return_data)

# # Calculate the number of shares to buy
# price_momentum_data = utils.calculate_number_of_shares_to_buy(portfolio_amount, price_momentum_data)

# # Save the recommended trades to excel
# price_momentum.save_recommended_trades(price_momentum_data)

# print("Recommended trades have been saved to 'recommended_trades.xlsx'.")
