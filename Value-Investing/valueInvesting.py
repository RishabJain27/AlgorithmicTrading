import numpy as np 
import pandas as pd 
import requests 
import math 
import xlsxwriter 
from scipy.stats import percentileofscore as score 
from statistics import mean
from sec import IEX_CLOUD_API_TOKEN

# Get S&P500 stocks
stocks = pd.read_csv('sp_500_stocks.csv')

# Get user input for number of shares
def portfolio_input():
    global portfolio_size
    portfolio_size = input("Enter the value of your portfolio:")

    try:
        val = float(portfolio_size)
    except ValueError:
        print("That's not a number! \n Try again:")
        portfolio_size = input("Enter the value of your portfolio:")
portfolio_input()

# Split list of stocks to list of 100 for API Call
def chunks(lst, n):
    """Yield successive n-sized chunks from lst."""
    for i in range(0, len(lst), n):
        yield lst[i:i + n]   
        
symbol_groups = list(chunks(stocks['Ticker'], 100))
symbol_strings = []
for i in range(0, len(symbol_groups)):
    symbol_strings.append(','.join(symbol_groups[i]))


# Build Value Investing Strategy Table
def extractQuote(list, sym):
    for elem in list:
        try:
            if sym in elem['symbol']:
                return elem
        except KeyError:
            continue

def extractStats(list, name):
    for elem in list:
        try:
            if name in elem['companyName'] and elem['peRatio'] and elem['priceToBook']:
                return elem
        except KeyError:
            continue

rv_columns = [
    'Ticker',
    'Price',
    'Number of Shares to Buy', 
    'Price-to-Earnings Ratio',
    'PE Percentile',
    'Price-to-Book Ratio',
    'PB Percentile',
    'Price-to-Sales Ratio',
    'PS Percentile',
    'EV/EBITDA',
    'EV/EBITDA Percentile',
    'EV/GP',
    'EV/GP Percentile',
    'RV Score'
]

rv_dataframe = pd.DataFrame(columns = rv_columns)

for symbol_string in symbol_strings:
    batch_api_url = f'https://api.iex.cloud/v1/data/core/ADVANCED_STATS,QUOTE/{symbol_string}?token={IEX_CLOUD_API_TOKEN}'
    data = requests.get(batch_api_url).json()
    for symbol in symbol_string.split(','):
        try:
            quote = extractQuote(data, symbol)
            companyName = quote['companyName']
            stats = extractStats(data, companyName)
            price = quote['latestPrice']
            peRatio = quote['peRatio']
            pb_ratio = stats['priceToBook']
            ps_ratio = stats['priceToSales']
            enterprise_value = stats['enterpriseValue']
            ebitda = stats['EBITDA']
            gross_profit = stats['grossProfit']
            
            try:
                ev_to_ebitda = enterprise_value/ebitda
            except TypeError:
                ev_to_ebitda = np.NaN
        
            try:
                ev_to_gross_profit = enterprise_value/gross_profit
            except TypeError:
                ev_to_gross_profit = np.NaN
            
            rv_dataframe = rv_dataframe._append(
            pd.Series(
                [
                    symbol,
                    price,
                    'N/A',
                    peRatio,
                    'N/A',
                    pb_ratio,
                    'N/A',
                    ps_ratio,
                    'N/A',
                    ev_to_ebitda,
                    'N/A',
                    ev_to_gross_profit,
                    'N/A',
                    'N/A'
                    ],
                    index = rv_columns),
                ignore_index=True
            )
        except TypeError:
            continue
# Handle Missing Data      
rv_dataframe[rv_dataframe.isnull().any(axis=1)]

# Calculate Percentiles 
metrics = {
            'Price-to-Earnings Ratio': 'PE Percentile',
            'Price-to-Book Ratio':'PB Percentile',
            'Price-to-Sales Ratio': 'PS Percentile',
            'EV/EBITDA':'EV/EBITDA Percentile',
            'EV/GP':'EV/GP Percentile'
}

for row in rv_dataframe.index:
    for metric in metrics.keys():
        rv_dataframe.loc[row, metrics[metric]] = score(rv_dataframe[metric], rv_dataframe.loc[row, metric])/100
    
# Calculate Mean of Values used in Table
for row in rv_dataframe.index:
    value_percentiles = []
    for metric in metrics.keys():
        value_percentiles.append(rv_dataframe.loc[row, metrics[metric]])
    rv_dataframe.loc[row, 'RV Score'] = mean(value_percentiles)
    
# Calculate Number of Shares to Buy
position_size = float(portfolio_size) / len(rv_dataframe.index)
for i in range(0, len(rv_dataframe['Ticker'])-1):
    rv_dataframe.loc[i, 'Number of Shares to Buy'] = math.floor(position_size / rv_dataframe['Price'][i])  

# Write to Excel
writer = pd.ExcelWriter('value_strategy.xlsx', engine='xlsxwriter')
rv_dataframe.to_excel(writer, sheet_name='Value Strategy', index = False)

background_color = '#0a0a23'
font_color = '#ffffff'

string_template = writer.book.add_format(
        {
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

dollar_template = writer.book.add_format(
        {
            'num_format':'$0.00',
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

integer_template = writer.book.add_format(
        {
            'num_format':'0',
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

float_template = writer.book.add_format(
        {
            'num_format':'0',
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

percent_template = writer.book.add_format(
        {
            'num_format':'0.0%',
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

column_formats = {
                    'A': ['Ticker', string_template],
                    'B': ['Price', dollar_template],
                    'C': ['Number of Shares to Buy', integer_template],
                    'D': ['Price-to-Earnings Ratio', float_template],
                    'E': ['PE Percentile', percent_template],
                    'F': ['Price-to-Book Ratio', float_template],
                    'G': ['PB Percentile',percent_template],
                    'H': ['Price-to-Sales Ratio', float_template],
                    'I': ['PS Percentile', percent_template],
                    'J': ['EV/EBITDA', float_template],
                    'K': ['EV/EBITDA Percentile', percent_template],
                    'L': ['EV/GP', float_template],
                    'M': ['EV/GP Percentile', percent_template],
                    'N': ['RV Score', percent_template]
                }

for column in column_formats.keys():
    writer.sheets['Value Strategy'].set_column(f'{column}:{column}', 25, column_formats[column][1])
    writer.sheets['Value Strategy'].write(f'{column}1', column_formats[column][0], column_formats[column][1])
    
writer.close()