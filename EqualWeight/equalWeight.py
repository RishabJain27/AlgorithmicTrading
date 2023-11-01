import numpy as np 
import pandas as pd 
import requests 
import xlsxwriter
import math

#Get API token
from sec import IEX_CLOUD_API_TOKEN

#Get Input from user on the amount of money user can invest
portfolio_size = input('Enter the value of your portfolio:')
try:
    val = float(portfolio_size)
except ValueError:
    print("That's not a number. \nPlease try again:")
    portfolio_size = print('Enter the value of your portfolio:')
    val = float(portfolio_size)

#Get S&P500 stock list
stocks = pd.read_csv('sp_500_stocks.csv')

#Split list by n
def batches(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i:i + n]
        

#Add Price and MarketCap to Table
my_columns = ['Ticker', 'Stock Price', 'Market Cap', 'Number of Shares to Buy']
symbol_groups = list(batches(stocks['Ticker'], 100))
symbol_strings = []
for i in range(0, len(symbol_groups)):
    symbol_strings.append(','.join(symbol_groups[i]))
final_dataframe = pd.DataFrame(columns=my_columns)

for symbol_string in symbol_strings:
    batch_api_call_url = f'https://api.iex.cloud/v1/data/core/quote/{symbol_string}?token={IEX_CLOUD_API_TOKEN}'
    data = requests.get(batch_api_call_url).json()
    for symbol in symbol_string.split(','):
        for elem in data:
            if symbol in elem['symbol']:
                latestPrice = elem['latestPrice']
                market_cap = elem['marketCap']
                break  
        final_dataframe = final_dataframe._append(
            pd.Series(
                [
                    symbol,
                    latestPrice,
                    market_cap,
                    'N/A'
                ], index=my_columns),
            ignore_index=True
        )

# Calculate number of shares
position_size = val/len(final_dataframe.index)
for i in range(0, len(final_dataframe.index)):
    final_dataframe.loc[i, 'Number of Shares to Buy'] = math.floor(position_size/final_dataframe.loc[i, 'Stock Price'])

# Write Trades to Excel File
writer = pd.ExcelWriter('equal_weight.xlsx', engine = 'xlsxwriter')
final_dataframe.to_excel(writer, 'Equal Weight Trades', index = False)

background_color = '#0066ff'
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

column_formats = { 
                    'A': ['Ticker', string_format],
                    'B': ['Price', dollar_format],
                    'C': ['Market Cap', dollar_format],
                    'D': ['Number of Shares to Buy', integer_format]
                    }

for column in column_formats.keys():
    writer.sheets['Equal Weight Trades'].set_column(f'{column}:{column}', 20, column_formats[column][1])
    writer.sheets['Equal Weight Trades'].write(f'{column}1', column_formats[column][0], string_format)
    
writer.close()
