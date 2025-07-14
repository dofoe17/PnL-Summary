# PnL-Summary
## Description
I wanted a better way to monitor the performance of my own portfolio holdings. This project provides a solution to this with a daily summary. 


First create a list of the tickers in portfolio and create date variables 

```ruby
tickers_list = ['GOOGL','AMZN','MSFT','SONY','AXP','AAPL','KO','EQNR','NKE']


dt = datetime.today().date() 

day_num = dt.strftime('%d')


#Monday
if dt.today().weekday() == 0:
    dt1 = (dt - timedelta(days=4)).strftime('%Y-%m-%d')
else: 
    dt1 = (dt - timedelta(days=2)).strftime('%Y-%m-%d')

#Subtract day_num from today's date 
som = dt - timedelta(days=int(day_num) - 1)
```


Next, pull in the relevant data in from yfinance and convert the data frame of prices into 1 dataframe after using pct_change() 

```ruby 

last_night_prices = yf.download(tickers=tickers_list, start=dt1, end=dt)['Close']
Day_before_last_prices = last_night_prices.head(1)
last_night_prices = last_night_prices.tail(1)

SOM_prices = yf.download(tickers=tickers_list, start=som, end=dt)['Close']
SOM_prices = SOM_prices.head(1)

#YTD row
SOY_prices = yf.download(tickers=tickers_list, start='2025-01-01', end=dt)['Close']
SOY_prices = SOY_prices.head(1)

Inception_prices = yf.download(tickers=tickers_list, start='2023-10-13', end='2023-10-14')['Close']

#Combine to one df 
prices_df = pd.concat([last_night_prices, Day_before_last_prices, SOM_prices, SOY_prices, Inception_prices]).sort_values(by='Date')

DTD_latest = pd.concat([Day_before_last_prices, last_night_prices])
DTD_pnl = DTD_latest.pct_change()

MTD_latest = pd.concat([SOM_prices, last_night_prices])
MTD_pnl = MTD_latest.pct_change()

YTD_latest = pd.concat([SOY_prices, last_night_prices])
YTD_pnl = YTD_latest.pct_change()

ITD_latest = pd.concat([Inception_prices, last_night_prices])
ITD_pnl = ITD_latest.pct_change()

#Combine to one df 
pnl_df = pd.concat([ITD_pnl, YTD_pnl, MTD_pnl, DTD_pnl]).sort_values(by='Date').dropna()
pnl_df = pnl_df.reset_index(drop=True) 
```

Create a new index column and add it to existing dataframe

```ruby
new_index = {'': ['ITD', 'YTD', 'MTD', 'DTD']} 
new_index_df = pd.DataFrame(new_index)

final_pnl_df = pd.merge(new_index_df, pnl_df, left_index=True, right_index=True)
```

Finally, import email libraries and foramt HTML table and send out final email

```ruby
import win32com.client as win32

#Format table
headers = {'selector': 'th.col_heading',
           'props': [('background-color', 'darkblue'), 
                     ('color', 'white')]}

borders = {'selector':'',
           'props': [('border', '1px solid black')]}



table_style = final_pnl_df.style.set_table_styles([headers, borders]).format({'AAPL': '{:,.2%}',
                                                                              'AMZN': '{:,.2%}',
                                                                              'AXP': '{:,.2%}',
                                                                              'EQNR': '{:,.2%}', 
                                                                              'GOOGL': '{:,.2%}',
                                                                              'KO': '{:,.2%}',
                                                                              'MSFT': '{:,.2%}',
                                                                              'NKE': '{:,.2%}',
                                                                              'SONY': '{:,.2%}'}).background_gradient(axis=None).hide()

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'darren.ofoe@gmail.com'
mail.Subject = 'PNL Summary ' + str(dt)
body = '<html><body><h4> Darren Ofoe - PNL History <h4>' + table_style.to_html() + '</body></html>'
mail.HTMLbody = body
mail.Display() 
```

Here's the final output which is sent to my gmail daily.

<img width="1105" height="619" alt="image" src="https://github.com/user-attachments/assets/1aed7c8b-9064-4ed6-b456-603d235c13f3" />


So we can see my top performing stock performance to date is American Express. Let's visualise this: 

'''ruby
#Create Visualisation of top performing stock AXP
ticker = 'AXP'
vis_df = yf.download(ticker, start='2023-10-13', end=dt1, multi_level_index=False)
fig = go.Figure(data=[go.Candlestick(x=vis_df.index,
                                     open=vis_df['Open'],
                                     high=vis_df['High'],
                                     low=vis_df['Low'],
                                     close=vis_df['Close'])])

fig.update_layout(
    title = f'{ticker} Stock Price',
    xaxis_title = 'Date',
    yaxis_title = 'Price (USD)' 
    #xaxis_rangeslider_visible = False  
)
fig.show() 
'''

<img width="1860" height="847" alt="image" src="https://github.com/user-attachments/assets/7607ff89-9232-4faf-8b10-0033e48cce83" />



