# Trade Analyzer for T-Bank Investments

## Description

The application analyzes how successfully you trade through the T-Bank Investments service. 

The application uses [T-Invest API](https://developer.tbank.ru/invest/intro/intro) and [Invest Python gRPC client](https://github.com/RussianInvestments/invest-python).

The application requires an API token to read information about your brokerage account and completed transactions. The token is used only for reading. Add your token to environment file.

The application is not able to perform operations on the exchange.