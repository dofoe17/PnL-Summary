#I want a better way to analyse/track my own personal holdings account performance
#Output should be a daily email sent to my GMail account with all details laid out 

import pandas as pd       
import yfinance as yf
import matplotlib.pyplot as plt 
import numpy as np
import time as t 
from datetime import datetime, timedelta

#def pnl():

#Create data frame (static of ticker, date of first deal, price, quantity/cost)
tickers_list = ['GOOGL','AMZN','MSFT','SONY','AXP','AAPL','KO','EQNR','NKE']

#today's date 
dt = datetime.today().date() 

day_num = dt.strftime('%d')

#Subtract day_num from today's date 
som = dt - timedelta(days=int(day_num) - 1)

dt1 = dt - timedelta(days = 2)

last_night_prices = yf.download(tickers=tickers_list, start=dt1, end=dt)['Close']
Day_before_last_prices = last_night_prices.head(1)
last_night_prices = last_night_prices.tail(1)
#print(Day_before_last_prices) 

SOM_prices = yf.download(tickers=tickers_list, start=som, end=dt)['Close']
#Take the first row of the data to get the start of month data only
SOM_prices = SOM_prices.head(1)

#YTD row
SOY_prices = yf.download(tickers=tickers_list, start='2025-01-01', end=dt)['Close']
SOY_prices = SOY_prices.head(1)
#print(SOY_prices)

Inception_prices = yf.download(tickers=tickers_list, start='2023-10-13', end='2023-10-14')['Close']
#print(Inception_prices)

#Combine to one df 
prices_df = pd.concat([last_night_prices, Day_before_last_prices, SOM_prices, SOY_prices, Inception_prices]).sort_values(by='Date')
#print(pnl_df)


DTD_latest = pd.concat([Day_before_last_prices, last_night_prices])
DTD_pnl = DTD_latest.pct_change()*100

MTD_latest = pd.concat([SOM_prices, last_night_prices])
MTD_pnl = MTD_latest.pct_change()*100

YTD_latest = pd.concat([SOY_prices, last_night_prices])
YTD_pnl = YTD_latest.pct_change()*100

ITD_latest = pd.concat([Inception_prices, last_night_prices])
ITD_pnl = ITD_latest.pct_change()*100 


#Combine to one df 
pnl_df = pd.concat([ITD_pnl, YTD_pnl, MTD_pnl, DTD_pnl]).sort_values(by='Date').dropna()
pnl_df = pnl_df.reset_index(drop=True) 
#print(pnl_df)


#New index list data
new_index = {'Timescale': ['ITD', 'YTD', 'MTD', 'DTD']} 
new_index_df = pd.DataFrame(new_index)



final_pnl_df = pd.merge(new_index_df, pnl_df, left_index=True, right_index=True)
#final_pnl_df = new_index_df.join(pnl_df) #Same thing as above
final_pnl_df = final_pnl_df.set_index('Timescale')
print(final_pnl_df)


#Format table for Email 
headers = {'selector':'th.col_headings',
            'props':'background-color:#000066; color:white;'}

table_style = final_pnl_df.style.set_table_styles([headers])


import win32com.client as win32 
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'darren.ofoe@gmail.com'
mail.Subject = 'PNL Summary' + str(dt1)
body = '<html><body><h4> Darren Ofoe - PNL History <h4>' + final_pnl_df.to_html() + '</body></html>'
mail.HTMLbody = body
mail.Display() 



#while(True):
#    pnl()
#    t.sleep(60)     