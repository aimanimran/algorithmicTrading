#!/usr/bin/env python
# coding: utf-8

# In[1]:


import numpy as np #computing
import pandas as pd #data science
import requests #HTTP requests
import xlsxwriter
import math


# In[2]:


stocks = pd.read_csv('sp_500_stocks.csv')
from secrets import ALPHA_VANTAGE_API


# In[5]:


my_columns = ['Ticker', 'Stock Price', 'Market Capitalization', 'Number of Shares to Buy']
final_dataframe = pd.DataFrame(columns = my_columns)


# In[7]:


final_dataframe = pd.DataFrame(columns = my_columns)
count = 0

for symbol in stocks['Ticker']:
    #FOR PRICE
    p_api_url = f'https://www.alphavantage.co/query?function=GLOBAL_QUOTE&symbol={symbol}&apikey={ALPHA_VANTAGE_API}'
    pData = requests.get(p_api_url).json()

    #FOR MARKET CAP
    mc_api_url = f'https://www.alphavantage.co/query?function=OVERVIEW&symbol={symbol}&apikey={ALPHA_VANTAGE_API}'
    mcData = requests.get(mc_api_url).json()
    
    #get values
    if 'Global Quote' in pData: 
        price = float(pData['Global Quote']['05. price'])
    else:
        price = 'Not Num'
        
    if 'MarketCapitalization' in mcData: 
        market_cap = mcData['MarketCapitalization']
    else:
        market_cap = 'Not Num'

    final_dataframe = final_dataframe.append(
        pd.Series(
        [
            symbol,
            price,
            market_cap,
            'N/A'
        ],
            index = my_columns
        ),
        ignore_index = True
    )
    count+=1
    
    if symbol=='ZTS':
        break
        
final_dataframe


# In[ ]:


final_dataframe


# In[11]:


portfolio_size = input("Enter the value of your portfolio:")

try:
    val = float(portfolio_size)
except ValueError:
    print("That's not a number! \n Try again:")
    portfolio_size = input("Enter the value of your portfolio:")


# In[12]:


position_size = float(portfolio_size) / len(final_dataframe.index)
for i in range(0, len(final_dataframe['Ticker'])-1):
    final_dataframe.loc[i, 'Number Of Shares to Buy'] = math.floor(position_size / final_dataframe['Price'][i])
final_dataframe


# In[13]:


writer = pd.ExcelWriter('recommended_trades.xlsx', engine='xlsxwriter')
final_dataframe.to_excel(writer, sheet_name='Recommended Trades', index = False)


# In[ ]:


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


# In[ ]:


column_formats = { 
                    'A': ['Ticker', string_format],
                    'B': ['Price', dollar_format],
                    'C': ['Market Capitalization', dollar_format],
                    'D': ['Number of Shares to Buy', integer_format]
                    }

for column in column_formats.keys():
    writer.sheets['Recommended Trades'].set_column(f'{column}:{column}', 20, column_formats[column][1])
    writer.sheets['Recommended Trades'].write(f'{column}1', column_formats[column][0], string_format)


# In[ ]:


writer.save()

