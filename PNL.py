import pandas as pd       
import yfinance as yf
from datetime import datetime, timedelta


import logging
log_path = r'C:\Users\Computer\OneDrive\Python\Pnl Summary\pnl_email.log'
logging.basicConfig(
    filename=log_path, 
    level=logging.INFO, 
    format='%(asctime)s - %(levelname)s - %(message)s',
    filemode='w'    #This clears the log at the start
)
logging.info('PNL Script started')

def main():
    try:
        logging.info('Main job started')
        #Create data frame (static of ticker, date of first deal, price, quantity/cost)
        tickers_list = ['GOOGL','AMZN','MSFT','SONY',
                        'AXP','AAPL','KO','EQNR','NKE']

        #today's date 
        dt = datetime.today().date() 
        day_num = dt.strftime('%d')


        #Monday
        if dt.weekday() == 0:
            dt1 = (dt - timedelta(days=4)).strftime('%Y-%m-%d')
        else: 
            dt1 = (dt - timedelta(days=2)).strftime('%Y-%m-%d')

        #Subtract day_num from today's date 
        som = dt - timedelta(days=int(day_num) - 1)


        last_night_prices = yf.download(tickers=tickers_list, start=dt1, end=dt)['Close']
        Day_before_last_prices = last_night_prices.head(1)
        last_night_prices = last_night_prices.tail(1)

        SOM_prices = yf.download(tickers=tickers_list, start=som, end=dt)['Close']
        #Take the first row of the data to get the start of month data only
        SOM_prices = SOM_prices.head(1)

        #YTD row
        SOY_prices = yf.download(tickers=tickers_list, start='2025-01-01', end=dt)['Close']
        SOY_prices = SOY_prices.head(1)

        Inception_prices = yf.download(tickers=tickers_list, start='2023-10-13', end='2023-10-14')['Close']

        DTD_latest = pd.concat([Day_before_last_prices, last_night_prices])
        DTD_pnl = DTD_latest.pct_change()

        MTD_latest = pd.concat([SOM_prices, last_night_prices])
        MTD_pnl = MTD_latest.pct_change()

        YTD_latest = pd.concat([SOY_prices, last_night_prices])
        YTD_pnl = YTD_latest.pct_change()

        ITD_latest = pd.concat([Inception_prices, last_night_prices])
        ITD_pnl = ITD_latest.pct_change()


        #Combine to one df 
        pnl_df = pd.concat([ITD_pnl, 
                            YTD_pnl, 
                            MTD_pnl, 
                            DTD_pnl]).sort_values(by='Date').dropna()
        pnl_df = pnl_df.reset_index(drop=True) 


        #New index list data
        new_index = {'': ['ITD', 'YTD', 'MTD', 'DTD']} 
        new_index_df = pd.DataFrame(new_index)

        final_pnl_df = pd.merge(new_index_df, pnl_df, left_index=True, right_index=True)
        
        #Make a copy to not overwrite original
        final_pnl_df_formatted = final_pnl_df.copy()

        for col in final_pnl_df_formatted.columns[1:]:  #Skip first column
            final_pnl_df_formatted[col] = final_pnl_df_formatted[col].map("{:,.2%}".format)

        table_html = final_pnl_df_formatted.to_html(index=False)


        logging.info("Sending Email...")

        import win32com.client as win32
        #Format table
        table_html = final_pnl_df_formatted.to_html(index=False)

        html_template = f"""
        <html> 
        <head> 
        <style> 
            table{{
                border-collapse: collapse;
                font-family: Arial, sans-serif;
                width:100%; 
            }}
            th, td {{
                border: 1px solid black; 
                padding: 6px; 
                text-align: center; 
            }}
            th {{
                background-color: #003366;
                color: white;
            }}
        </style> 
        </head> 
        <body> 
            <h4>Darren Ofoe - PNL History</h4> 
            {table_html}
        </body> 
        </html> 
        """ 

        try:
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = 'darren.ofoe@gmail.com'
            mail.Subject = 'PNL Summary ' + str(dt)
            mail.HTMLbody = html_template
            mail.Send() 
            logging.info('Email sent successfully!')
        except Exception as e: 
            logging.exception('Failed to send Email')

    except Exception as e:
        logging.exception('PNL Script failed')
        raise 

if __name__ == "__main__": 
    main() 