# fundamental 
## Dataframe for fundamental financials of multiple stocks
It is a Python 3 library for generating list of stocks' fundamental data through YahooFinance.

The library is primarily based on yfinance package. 

fundamental requires list of tickers input from the user and generate dataframe of fundamentals.

It can be used for fundamental ratio analysis, relative value ratio analysis and industry relative analysis.


## Data inclusion
- financials for recent 3 years. (income statement, balance sheet, cash flow statement)
- calendar year end financials prorated by months
- Trailing twelve months financial by aggregate recent 4 quarters
- Basic share and price for each fiscal year end and recent fiscal quarter end
- It takes about 10 minutes to create a dataframe for 100 stocks. 


## Usage
```
# pip install fundamental

from fundamental import fundamental

symlist = ['MSFT','GOOG','FB'] 
df = fundamental.get_df_list(symlist)        

```

## Limitation
- slow internet connection would lead scraping error and the program will auto try 3 times. 

