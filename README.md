# Trade Analyzer for T-Bank Investments

![Python](https://img.shields.io/badge/python-3.11-blue.svg)
[![Checked with mypy](http://www.mypy-lang.org/static/mypy_badge.svg)](https://mypy-lang.org/)
[![license](https://img.shields.io/badge/licence-MIT-green.svg)](https://opensource.org/licenses/MIT)

## Description

The application analyzes how successfully you trade through the T-Bank Investments service. 

The application uses [T-Invest API](https://developer.tbank.ru/invest/intro/intro) and [Invest Python gRPC client](https://github.com/RussianInvestments/invest-python).

The application requires an API token to read information about your brokerage account and completed transactions. The token is used only for reading. Add your token to environment file.

The application is not able to perform operations on the exchange.

## 📚 References 

- [T-Invest API](https://developer.tbank.ru/invest/intro/intro)
- [Invest Python gRPC client](https://github.com/RussianInvestments/invest-python)
- [tinkoff-investments on PyPI](https://pypi.org/project/tinkoff-investments/)
