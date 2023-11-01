import numpy as np 
import pandas as pd 
import requests 
import math 
from scipy.stats import percentileofscore as score
from statistics import mean
import xlsxwriter

#Get API token
from sec import IEX_CLOUD_API_TOKEN

#Get Size of Portfolio from user
def portfolio_Input():
    global portfolio_size
    portfolio_size = input('Enter the size of your portfolio: ')
    try:
        float(portfolio_size)
    except ValueError:
        print('This is not a number. Please try again: ')
        portfolio_size = input('Enter the size of your portfolio: ')

portfolio_Input();


#Get s&p500 stocks
stocks = pd.read_csv('sp_500_stocks.csv')

# Break list of stocks to list of 100 for API call
def chunks(lst, n):
    """Yield successive n-sized chunks from lst."""
    for i in range(0, len(lst), n):
        yield lst[i:i + n]   
        
symbol_groups = list(chunks(stocks['Ticker'], 100))
symbol_strings = []
for i in range(0, len(symbol_groups)):
    symbol_strings.append(','.join(symbol_groups[i]))
    
#Create columns for Table
hqm_columns = [
    'Ticker',
    'Price',
    'Number of Shares to Buy',
    'One-Year Price Return',
    'One-Year Return Percentile',
    'Six-Month Price Return',
    'Six-Month Return Percentile',
    'Three-Month Price Return',
    'Three-Month Return Percentile',
    'One-Month Price Return',
    'One-Month Return Percentile',
    'HQM Score'
]
hqm_dataframe = pd.DataFrame(columns = hqm_columns)

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
            if name in elem['companyName'] and elem['year1ChangePercent']:
                return elem
        except KeyError:
            continue
        
#Populate Data from API call to Table 
for symbol_string in symbol_strings:
    batch_api_url = f'https://api.iex.cloud/v1/data/core/ADVANCED_STATS,QUOTE/{symbol_string}?token={IEX_CLOUD_API_TOKEN}'
    data = requests.get(batch_api_url).json()
    #print(data)
    for symbol in symbol_string.split(','):
        try:
            quote = extractQuote(data, symbol)
            companyName = quote['companyName']
            stats = extractStats(data, companyName)
            price = quote['latestPrice']
            yearChange = stats['year1ChangePercent']
            sixthMonthChange = stats['month6ChangePercent']
            threeMonthChange = stats['month3ChangePercent']
            oneMonthChange = stats['month1ChangePercent']
            hqm_dataframe = hqm_dataframe._append(
            pd.Series(
                [
                    symbol,
                    price,
                    'N/A',
                    yearChange,
                    'N/A',
                    sixthMonthChange,
                    'N/A',
                    threeMonthChange,
                    'N/A',
                    oneMonthChange,
                    'N/A',
                    'N/A'
                    ],
                    index = hqm_columns),
                ignore_index=True
            )
        except TypeError:
            continue

#Calculate Momentum Percentiles 
time_periods = [
                'One-Year',
                'Six-Month',
                'Three-Month',
                'One-Month'
                ]

for row in hqm_dataframe.index:
    for time_period in time_periods:
        change_col = f'{time_period} Price Return'
        percentile_col = f'{time_period} Return Percentile'
        hqm_dataframe.loc[row, percentile_col] = score(hqm_dataframe[change_col], hqm_dataframe.loc[row, change_col])/100
        

#Calculate HQM mean score 
for row in hqm_dataframe.index:
    momentum_percentiles = []
    for time_period in time_periods:
        momentum_percentiles.append(hqm_dataframe.loc[row, f'{time_period} Return Percentile'])
    hqm_dataframe.loc[row, 'HQM Score'] = mean(momentum_percentiles)
    

#Filter by best 50
hqm_dataframe.sort_values('HQM Score', ascending = False, inplace = True)
hqm_dataframe = hqm_dataframe[:50]
hqm_dataframe.reset_index(inplace = True, drop = True)

#Calculate number of shares
position_size = float(portfolio_size)/ len(hqm_dataframe.index)
for i in hqm_dataframe.index:
    hqm_dataframe.loc[i, 'Number of Shares to Buy'] = math.floor(position_size/hqm_dataframe.loc[i, 'Price'])

# Formate Excel Output
writer = pd.ExcelWriter('momentum_trades.xlsx', engine = 'xlsxwriter')
hqm_dataframe.to_excel(writer, 'Momentum Strategy', index = False)

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
                    'D': ['One-Year Price Return', percent_template],
                    'E': ['One-Year Return Percentile', percent_template],
                    'F': ['Six-Month Price Return', percent_template],
                    'G': ['Six-Month Return Percentile', percent_template],
                    'H': ['Three-Month Price Return', percent_template],
                    'I': ['Three-Month Return Percentile', percent_template],
                    'J': ['One-Month Price Return', percent_template],
                    'K': ['One-Month Return Percentile', percent_template],
                    'L': ['HQM Score', integer_template]
                    }

for column in column_formats.keys():
    writer.sheets['Momentum Strategy'].set_column(f'{column}:{column}', 20, column_formats[column][1])
    writer.sheets['Momentum Strategy'].write(f'{column}1', column_formats[column][0], string_template)

writer.close()