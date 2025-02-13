import os
import time
import pytz
import requests
import pandas as pd
from pathlib import Path
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from collections import OrderedDict
from datetime import datetime, timedelta, timezone


from tinkoff.invest import Client, GetOperationsByCursorRequest
from tinkoff.invest.services import Services
from tinkoff.invest.constants import INVEST_GRPC_API
from tinkoff.invest.schemas import OperationType, OperationState, TradeDirection, CandleInterval
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import mplfinance as mpf


from utils import clear_or_create_dir, format_xlsx, crop_image_white_margins

pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

load_dotenv()
TIMEZONE = os.getenv("TIMEZONE")
os.environ["TZ"] = TIMEZONE
time.tzset()
INPUT_PATH = Path(os.getenv("INPUT_PATH"))
OUTPUT_PATH = Path(os.getenv("OUTPUT_PATH"))
IMG_WIDTH, IMG_HEIGHT, IMG_DPI = 3600, 2000, 150
TOKEN = os.getenv("TOKEN")
CB_CURRENCIES_URL = os.getenv("CB_CURRENCIES_URL")

operations_types = {
    OperationType.OPERATION_TYPE_UNSPECIFIED: 'Неизвестно',
    OperationType.OPERATION_TYPE_INPUT: 'Пополнение брокерского счета',
    OperationType.OPERATION_TYPE_OUTPUT: 'Вывод денег',
    OperationType.OPERATION_TYPE_BUY_CARD: 'Покупка с карты',
    OperationType.OPERATION_TYPE_BUY: 'Покупка',
    OperationType.OPERATION_TYPE_SELL: 'Продажа',
    OperationType.OPERATION_TYPE_BROKER_FEE: 'Комиссия брокера',
    OperationType.OPERATION_TYPE_DIVIDEND: 'Выплата дивидендов',
    OperationType.OPERATION_TYPE_TAX: 'Налоги',
    OperationType.OPERATION_TYPE_OVERNIGHT: 'Овернайт',
    OperationType.OPERATION_TYPE_DIVIDEND_TAX: 'Налоги c дивидендов',
    OperationType.OPERATION_TYPE_TAX_CORRECTION: 'Возврат налога',
    OperationType.OPERATION_TYPE_SERVICE_FEE: 'Комиссия за обслуживание'
}

operations_states = {
    OperationState.OPERATION_STATE_UNSPECIFIED: 'Неизвестно',
    OperationState.OPERATION_STATE_EXECUTED: 'Выполнена',
    OperationState.OPERATION_STATE_CANCELED: 'Отменена',
    OperationState.OPERATION_STATE_PROGRESS: 'В процессе'
}

instrument_types = {
    'bond': 'облигация',
    'share': 'акция',
    'currency': 'валюта',
    'etf': 'фонд',
    'futures': 'фьючерс'
}

trade_directions = {
    TradeDirection.TRADE_DIRECTION_UNSPECIFIED: 'Направление сделки не определено',
    TradeDirection.TRADE_DIRECTION_BUY: 'Покупка',
    TradeDirection.TRADE_DIRECTION_SELL: 'Продажа'
}


def get_cb_currencies_rate():
    """Retrieves currency exchange rates from the Central Bank's website.

        This function fetches the latest currency exchange rates from a specified URL
        (CB_CURRENCIES_URL), parses the HTML content, and extracts the relevant data.
        It returns a dictionary where keys are currency names and values are lists
        containing the currency's full name, nominal value, and exchange rate.

        Returns: An ordered dictionary containing currency exchange rates
    """
    currency_rates = OrderedDict(
        [
            ('RUB', {'code': 643, 'nominal': 1, 'name': 'Российский рубль', 'rate': 1.0})
        ]
    )
    try:
        resp = requests.get(CB_CURRENCIES_URL)
        resp.raise_for_status()  # Raise HTTPError for bad responses (4xx or 5xx)
        resp.encoding = 'utf-8'
        soup = BeautifulSoup(resp.text, 'lxml')
        table = soup.find('table')  # Assumes data is in a table.  Adjust if needed.
        if table is None:
            raise ValueError("Could not find currency table on the page.")

        rows = table.find_all('tr')
        # Skip header row
        for row in rows[1:]:  # Assumes first row is header. Adjust if needed.
            cells = row.find_all('td')
            if len(cells) < 4:
                raise ValueError(f"Unexpected number of cells in row: {row}")
            short_name = cells[1].text.strip()
            currency_rates[short_name] = {
                'code': cells[0].text.strip(),
                'nominal': cells[2].text.strip(),
                'name': cells[3].text.strip(),
                'rate': cells[4].text.strip()
            }
    except requests.exceptions.RequestException as e:
        print(f"Error fetching data: {e}")
    except ValueError as e:
        print(f"Error parsing data: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
    return currency_rates


currencies_rates = get_cb_currencies_rate()


def colorize_xlsx(writer: pd.ExcelWriter, df: pd.DataFrame, sheet_name: str = 'Sheet1'):
    """
    Function for coloring cells in an XlsxWriter object based on their values

    Args:
        param writer (pandas.io.excel._xlsxwriter._XlsxWriter): object of type XlsxWriter
        param df (pandas.core.frame.DataFrame): pandas dataframe with data
        param sheet_name (str): sheet name
    """
    if df.shape[0] > 0:
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        condition_format_1 = workbook.add_format({'bg_color': '#fa6464'})  # red
        condition_format_2 = workbook.add_format({'bg_color': '#faed64'})  # yellow
        condition_format_3 = workbook.add_format({'bg_color': '#64fa6e'})  # green
        condition_format_4 = workbook.add_format({'bg_color': '#64c0fa'})  # blue
        header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
                                             'border': 1})
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        worksheet.conditional_format('B2:B1000', {'type': 'cell', 'criteria': 'equal to',
                                                  'value': f'"{operations_types[OperationType.OPERATION_TYPE_INPUT]}"',
                                                  'format': condition_format_4})
        worksheet.conditional_format('B2:B1000', {'type': 'cell', 'criteria': 'equal to',
                                                  'value': f'"{operations_types[OperationType.OPERATION_TYPE_BUY]}"',
                                                  'format': condition_format_1})
        worksheet.conditional_format('B2:B1000', {'type': 'cell', 'criteria': 'equal to',
                                                  'value': f'"{operations_types[OperationType.OPERATION_TYPE_SELL]}"',
                                                  'format': condition_format_3})
        worksheet.conditional_format('B2:B1000', {'type': 'cell', 'criteria': 'equal to',
                                                  'value': f'"{operations_types[OperationType.OPERATION_TYPE_BROKER_FEE]}"',
                                                  'format': condition_format_1})
        worksheet.conditional_format('B2:B1000', {'type': 'cell', 'criteria': 'equal to',
                                                  'value': f'"{operations_types[OperationType.OPERATION_TYPE_OUTPUT]}"',
                                                  'format': condition_format_2})
    return writer

def get_stock_candles(client: Services, figi: str, start_date: datetime):
    stocks = client.instruments.shares().instruments
    st = [x for x in stocks if x.figi == figi]
    if len(st) > 0:
        stock = st[0]
        # ws = ['blueskies', 'brasil', 'charles', 'checkers', 'classic', 'default', 'mike',
        #       'nightclouds', 'sas', 'starsandstripes', 'yahoo']
        mc = mpf.make_marketcolors(
            up='#00c93c',
            down='#e30505',
            volume='#1649c9',
            edge='none')
        style = mpf.make_mpf_style(
            base_mpf_style='yahoo',
            rc={
                'axes.titlesize': 14,
                'axes.labelsize': 8,
                'xtick.labelsize': 6,
                'ytick.labelsize': 6
            },
            marketcolors=mc,
            gridaxis='both',
            gridcolor='#DCDCDC',
            mavcolors=['#4f8a8b', '#fbd46d', '#87556f']
        )
        cd2 = start_date.astimezone(tz=timezone.utc)
        cd1 = cd2 - timedelta(days=2)
        candles_list = []
        for candle in client.get_all_candles(
                figi=figi,
                from_=cd1, to=cd2,
                interval=CandleInterval.CANDLE_INTERVAL_30_MIN,
        ):
                candles_list.append({
                    "Date": candle.time,
                    "Open": float(candle.open.units + candle.open.nano * 10**-9),
                    "High": float(candle.high.units + candle.high.nano * 10**-9),
                    "Low": float(candle.low.units + candle.low.nano * 10**-9),
                    "Close": float(candle.close.units + candle.close.nano * 10**-9),
                    "Volume": float(candle.volume)}
                )
        candles_df = pd.DataFrame(candles_list, columns=["Date", "Open", "High", "Low", "Close", "Volume"])
        candles_df['Date'] = candles_df['Date'].dt.tz_convert(TIMEZONE)
        candles_df.set_index('Date', inplace=True)
        title = f"\n\n\n\n{stock.ticker} {stock.name} {cd1.strftime('%d.%m.%Y')} - {cd2.strftime('%d.%m.%Y')}"
        fig, axes = mpf.plot(candles_df, type='candle', datetime_format='%d.%m.%y %H:%M', figratio=(16, 9),
                             volume=False, returnfig=True, xlabel= 'Time', ylabel='Price', xrotation=45,
                             show_nontrading=True, style=style, axtitle=title, tight_layout=True,)
        ax = axes[0]
        ax.xaxis.set_major_locator(mdates.MinuteLocator(interval=60))
        filename = f"{OUTPUT_PATH}/{stock.ticker}.png"
        fig.savefig(filename, dpi=600, bbox_inches='tight' )
        plt.close()
        crop_image_white_margins(old_filename=filename, new_filename= filename)
    else:
        print(f'Stock with FIGI={figi} not found in shared stocks data')


def get_operations_df(client: Services, start_date: datetime, end_date: datetime) -> pd.DataFrame:
    stocks = client.instruments.shares().instruments
    account_id = client.users.get_accounts().accounts[0].id

    def get_request(cursor=""):
        return GetOperationsByCursorRequest(
            account_id=account_id,
            instrument_id=None,
            state=OperationState.OPERATION_STATE_EXECUTED,
            cursor=cursor,
            from_=start_date,
            to=end_date,
            limit=5,
        )

    operations = client.operations.get_operations_by_cursor(get_request())
    operations_list = []
    while operations.has_next:
        for operation in operations.items:
            d = operation.date.astimezone(tz=pytz.timezone(TIMEZONE))
            r = [x for x in stocks if x.figi == operation.figi]
            operations_list.append({
                "Дата": d,
                "Тип операции": operations_types[operation.type.value],
                "FIGI": operation.figi,
                "Тикер": r[0].ticker if len(r) > 0 else '',
                "Название": r[0].name if len(r) > 0 else '',
                "Тип": r[0].share_type if len(r) > 0 else '',
                "Сумма": round(operation.payment.units + operation.payment.nano / 1e9, 2),
                "Цена лота": round(operation.price.units + operation.price.nano / 1e9, 2),
                "Количество лотов": operation.quantity,
                "Статус": operations_states[operation.state]
            })
        request = get_request(cursor=operations.next_cursor)
        operations = client.operations.get_operations_by_cursor(request)
    operations_df = pd.DataFrame(operations_list, columns=["Дата", "Тип операции", "Название", "Тикер", "FIGI",
                                                           "Цена лота", "Количество лотов", "Сумма", "Статус"])
    operations_df.sort_values(by='Дата', ascending=True, inplace=True)
    operations_df['Дата'] = operations_df['Дата'].apply(lambda x: x.strftime('%d.%m.%Y   %H.%M.%S'))
    return operations_df


def get_money_df(client: Services) -> pd.DataFrame:
    account_id = client.users.get_accounts().accounts[0].id
    positions = client.operations.get_positions(account_id=account_id)
    money_list = []
    for money in positions.money:
        currency_name = str(money.currency).upper()
        currency_balance = money.units + money.nano / 10**9
        currency = currencies_rates[currency_name]
        currency_balance = round(float(currency_balance) / currency['nominal'] * currency['rate'], 3)
        money_list.append({
                "Валюта": currency['name'],
                "Сумма": currency_balance,
                "Сумма в рублях": currency_balance
            })
    money_df = pd.DataFrame(money_list, columns=["Валюта", "Сумма", "Сумма в рублях"])
    return money_df


def get_stocks_df(client: Services) -> pd.DataFrame:
    account_id = client.users.get_accounts().accounts[0].id
    stocks = client.instruments.shares().instruments
    portfolio = client.operations.get_portfolio(account_id=account_id)
    stocks_list = []
    for stock in portfolio.positions:
        if stock.instrument_type == 'share':
            r = [x for x in stocks if x.figi == stock.figi]
            current_price = stock.current_price.units + stock.current_price.nano * 10**-9
            quantity = stock.quantity.units + stock.quantity.nano * 10**-9
            quantity_lots = stock.quantity_lots.units + stock.quantity_lots.nano * 10**-9
            average_position_price = stock.average_position_price.units + stock.average_position_price.nano * 10**-9
            expected_yield = stock.expected_yield.units + stock.expected_yield.nano * 10**-9
            total_cost = quantity * current_price
            stocks_list.append({
                "Название": r[0].name if len(r) > 0 else '',
                "FIGI": stock.figi,
                "Тикер": r[0].ticker if len(r) > 0 else '',
                "Текущая цена акции": current_price,
                "Количество акций": quantity,
                "Количество лотов": quantity_lots,
                "Общая стоимость акций": total_cost,
                "Средневзвешенная цена акции": average_position_price,
                "Доход за все время": expected_yield}
            )
        stocks_df = pd.DataFrame(stocks_list, columns=["Название", "FIGI", "Тикер", "Количество акций",
                                                      "Количество лотов", "Текущая цена акции",
                                                      "Средневзвешенная цена акции", "Общая стоимость акций",
                                                      "Доход за все время"])
        return stocks_df


def get_last_trades(client: Services, figi: str, start_date: datetime, end_date: datetime):
    trades = client.market_data.get_last_trades(figi=figi, from_=start_date, to=end_date)
    for trade in trades.trades:
        print(f'{trade_directions[trade.direction]}, количество: {trade.quantity}, цена {trade.price.units + trade.price.nano * 10**-9}')


def get_orders(client: Services, figi: str, depth: int):
    orders = client.market_data.get_order_book(figi=figi, depth=depth)
    # Множество пар значений на покупку
    for bid in orders.bids:
        print(f'Покупка, количество: {bid.quantity}, цена {bid.price.units + bid.price.nano * 10**-9}')
    # Множество пар значений на продажу
    for ask in orders.asks:
        print(f'Продажа, количество: {ask.quantity}, цена {ask.price.units + ask.price.nano * 10**-9}')


def make_report(client: Services, start_date: datetime, end_date: datetime):
    writer = pd.ExcelWriter(f"report.xlsx", engine='xlsxwriter')
    operations_df = get_operations_df(client, start_date, end_date)
    operations_df.to_excel(excel_writer=writer, sheet_name='Операции', header=True, index=False)
    writer = format_xlsx(writer, operations_df, alignments='cllcccccc', sheet_name='Операции')
    writer = colorize_xlsx(writer, operations_df, sheet_name='Операции')

    unique_figi = operations_df['FIGI'].dropna().unique()
    unique_figi = unique_figi[unique_figi != '']

    for figi in unique_figi:
        get_stock_candles(client, figi, end_date)

    total_operations_df = operations_df.groupby(by="Тип операции", as_index=False)['Сумма'].sum()
    # total_operations_df['Сумма'] = total_operations_df['Сумма'].apply(lambda x: abs(float(x)))
    total_operations_df.to_excel(excel_writer=writer, sheet_name='Итог', header=True, index=False)
    writer = format_xlsx(writer, total_operations_df, 'lc', sheet_name='Итог')

    money_df = get_money_df(client)
    money_df.to_excel(excel_writer=writer, sheet_name='Валюта', header=True, index=False)
    writer = format_xlsx(writer, money_df, alignments='c' * money_df.shape[1], sheet_name='Валюта')

    stocks_df = get_stocks_df(client)
    stocks_df.to_excel(excel_writer=writer, sheet_name='Акции', header=True, index=False)
    writer = format_xlsx(writer, stocks_df, alignments='c' * stocks_df.shape[1], sheet_name='Акции')
    writer.close()


def main():
    with Client(TOKEN, target=INVEST_GRPC_API) as client:
        clear_or_create_dir(OUTPUT_PATH)
        start_date = datetime(2024, 10, 1, 0, 0, 0).replace(tzinfo=timezone.utc)
        # end_date = datetime(2024, 11, 1, 0, 0, 0).replace(tzinfo=timezone.utc)
        end_date = datetime.now().replace(tzinfo=timezone.utc)
        make_report(client, start_date, end_date)


if __name__ == '__main__':
    start_time = time.time()
    print("Started...")
    main()
    print(f"Done. Elapsed time {round((time.time() - start_time), 1)} seconds")