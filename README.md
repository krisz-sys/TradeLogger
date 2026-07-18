This program uses Python to log trading activity on Binance and MEXC.
Crypto trading platforms usually only retain trade history for three months, and logging transactions manually can be exhausting. 
This program automates that logging process.


You need to fill out the settings.xlsx file using your data by following the instructions inside it.

The second input file is CryptoPortfolio1.xlsx, which stores the previously logged trades. 
The program automatically appends the new trades to these existing logs.

The output .xlsx file will contain both the existing trades and the new ones fetched from the platforms.
