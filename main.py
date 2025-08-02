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

import mplfinance as mpf
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

from tinkoff.invest.services import Services
from tinkoff.invest.constants import INVEST_GRPC_API
from tinkoff.invest import Client, GetOperationsByCursorRequest
from tinkoff.invest.schemas import OperationType, OperationState, TradeDirection, CandleInterval

from libs.utils import clear_or_create_dir, format_xlsx, crop_image_white_margins


pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)


load_dotenv('env/.env')
TIMEZONE = os.getenv("TIMEZONE")
os.environ["TZ"] = TIMEZONE
time.tzset()
INPUT_PATH = Path(os.getenv("INPUT_PATH"))
OUTPUT_PATH = Path(os.getenv("OUTPUT_PATH"))
IMG_WIDTH, IMG_HEIGHT, IMG_DPI = 3600, 2000, 150
TOKEN = os.getenv("TOKEN", '')
CB_CURRENCIES_URL = os.getenv("CB_CURRENCIES_URL")
REQUEST_DELAY_SECONDS = float(os.getenv("REQUEST_DELAY_SECONDS"))

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
    OperationType.OPERATION_TYPE_SERVICE_FEE: 'Комиссия за обслуживание',
    OperationType.OPERATION_TYPE_BENEFIT_TAX: 'Налог за материальную выгоду',
    OperationType.OPERATION_TYPE_BENEFIT_TAX_PROGRESSIVE: 'Налог за материальную выгоду по ставке 15%'
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


def get_cb_currencies_rate() -> dict:
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


def colorize_operations_report(writer: pd.ExcelWriter, df: pd.DataFrame, sheet_name: str = 'Sheet1') -> pd.ExcelWriter:
    """
    Function for coloring cells in an XlsxWriter object based on their values (for total operations report)

    Args:
        param writer (pandas.io.excel._xlsxwriter._XlsxWriter): object of type XlsxWriter
        param df (pandas.core.frame.DataFrame): pandas dataframe with data
        param sheet_name (str): sheet name
    """
    if df.shape[0] > 0:
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'vcenter',
            'align': 'center',
            'border': 1
        })
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        cells = f'B2:B{df.shape[0] + 1}'
        condition_format_red = workbook.add_format({'bg_color': '#fa6464'})
        condition_format_yellow = workbook.add_format({'bg_color': '#faed64'})
        condition_format_green = workbook.add_format({'bg_color': '#64fa6e'})
        condition_format_blue = workbook.add_format({'bg_color': '#64c0fa'})
        condition_format_by_operations_types = {
            OperationType.OPERATION_TYPE_INPUT: condition_format_blue,
            OperationType.OPERATION_TYPE_BUY: condition_format_red,
            OperationType.OPERATION_TYPE_SELL: condition_format_green,
            OperationType.OPERATION_TYPE_BROKER_FEE: condition_format_red,
            OperationType.OPERATION_TYPE_OUTPUT: condition_format_yellow,
        }
        for operation_type, condition_format in condition_format_by_operations_types.items():
            worksheet.conditional_format(cells, {
                'type': 'cell',
                'criteria': 'equal to',
                'value': f'"{operations_types[operation_type]}"',
                'format': condition_format
            })
    return writer

def colorize_companies_report(writer: pd.ExcelWriter, df: pd.DataFrame, sheet_name: str = 'Sheet1') -> pd.ExcelWriter:
    """
    Function for coloring cells in an XlsxWriter object based on their values (for total operations by companies report)

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
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'vcenter',
            'align': 'center',
            'border': 1
        })
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        cells = f'B2:B{df.shape[0] + 1}'
        worksheet.conditional_format(cells, {
            'type': 'cell',
            'criteria': 'less than',
            'value': 0,
            'format': condition_format_1
        })
        worksheet.conditional_format(cells, {
            'type': 'cell',
            'criteria': 'equal to',
            'value': 0,
            'format': condition_format_2
        })
        worksheet.conditional_format(cells, {
            'type': 'cell',
            'criteria': 'greater than',
            'value': 0,
            'format': condition_format_3
        })
    return writer


def get_stock_candles(client: Services, figi: str, start_date: datetime, candles_path: str | Path) -> None:
    candles_path = Path(candles_path)
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
        cd1 = (cd2 - timedelta(days=2))
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
        if len(candles_list) > 0:
            candles_df = pd.DataFrame(candles_list, columns=["Date", "Open", "High", "Low", "Close", "Volume"])
            candles_df['Date'] = candles_df['Date'].dt.tz_convert(TIMEZONE)
            candles_df.set_index('Date', inplace=True)
            title = f"\n\n\n\n{stock.ticker} {stock.name} {cd1.strftime('%d.%m.%Y')} - {cd2.strftime('%d.%m.%Y')}"
            fig, axes = mpf.plot(candles_df, type='candle', datetime_format='%d.%m.%y %H:%M', figratio=(16, 9),
                                 volume=False, returnfig=True, xlabel= 'Time', ylabel='Price', xrotation=45,
                                 show_nontrading=True, style=style, axtitle=title, tight_layout=True,)
            ax = axes[0]
            ax.xaxis.set_major_locator(mdates.MinuteLocator(interval=60))
            filename = candles_path / f'{stock.ticker}.png'
            fig.savefig(filename, dpi=600, bbox_inches='tight' )
            plt.close()
            crop_image_white_margins(old_filename=filename, new_filename= filename)
        else:
            print(f'Candles for stock with FIGI={figi} not found')
    else:
        print(f'Stock with FIGI={figi} not found in shared stocks data')


def get_operations_df(client: Services, start_date: datetime, end_date: datetime) -> pd.DataFrame:
    stocks = client.instruments.shares().instruments
    account_id = client.users.get_accounts().accounts[0].id

    def get_request(cursor: str="") -> GetOperationsByCursorRequest:
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
                "Тип операции": operations_types[operation.type],
                "FIGI": operation.figi,
                "Тикер": r[0].ticker if len(r) > 0 else '',
                "Название": r[0].name if len(r) > 0 else '',
                "Тип": r[0].share_type if len(r) > 0 else '',
                "Сумма": round(operation.payment.units + operation.payment.nano / 1e9, 2),
                "Цена лота": round(operation.price.units + operation.price.nano / 1e9, 2),
                "Количество лотов": operation.quantity,
                "Статус": operations_states[operation.state]
            })
        time.sleep(REQUEST_DELAY_SECONDS)
        request = get_request(cursor=operations.next_cursor)
        operations = client.operations.get_operations_by_cursor(request)
    operations_df = pd.DataFrame(operations_list, columns=["Дата", "Тип операции", "Название", "Тикер", "FIGI",
                                                           "Цена лота", "Количество лотов", "Сумма", "Статус"])
    operations_df.sort_values(by='Дата', ascending=True, inplace=True)
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


def get_last_trades(client: Services, figi: str, start_date: datetime, end_date: datetime) -> None:
    trades = client.market_data.get_last_trades(figi=figi, from_=start_date, to=end_date)
    for trade in trades.trades:
        print(f'{trade_directions[trade.direction]}, количество: {trade.quantity}, цена {trade.price.units + trade.price.nano * 10**-9}')


def get_orders(client: Services, figi: str, depth: int) -> None:
    orders = client.market_data.get_order_book(figi=figi, depth=depth)
    # Множество пар значений на покупку
    for bid in orders.bids:
        print(f'Покупка, количество: {bid.quantity}, цена {bid.price.units + bid.price.nano * 10**-9}')
    # Множество пар значений на продажу
    for ask in orders.asks:
        print(f'Продажа, количество: {ask.quantity}, цена {ask.price.units + ask.price.nano * 10**-9}')


def get_favourite_instruments(client: Services) -> list[dict[str, str]]:
    favourite = client.instruments.get_favorites()
    return [
        {
            'figi': i.figi,
            'ticker': i.ticker,
            'name': i.name,
            'type': i.instrument_type
        } for i in favourite.favorite_instruments
    ]


def get_stocks_by_date(operations_df: pd.DataFrame, stocks_df: pd.DataFrame) -> pd.DataFrame:
    cur_stocks_amount = {}
    for index, row in stocks_df.iterrows():
        cur_stocks_amount['Название'] = row["Количество лотов"]
    stocks_by_date_list = []
    d = datetime.now().date()
    for k, v in cur_stocks_amount.items():
        stocks_by_date_list.append({
            'Дата': d,
            'Название': k,
            'Количество лотов': v
        })
    operations_df = operations_df.loc[operations_df['Тип операции'].isin(['Продажа', 'Покупка'])]
    operations_df['Дата'] = operations_df['Дата'].dt.date
    operations_df.sort_values(by='Дата', ascending=False, inplace=True)
    d = d - timedelta(days=1)
    for index, row in operations_df.iterrows():
        if row['Дата'] < d:
            for k, v in cur_stocks_amount.items():
                stocks_by_date_list.append({
                    'Дата': d,
                    'Название': k,
                    'Количество лотов': v
                })
            d = d - timedelta(days=1)
        if row['Тип операции'] == 'Продажа':
            cur_stocks_amount[row['Название']] = cur_stocks_amount.get(row['Название'], 0) - row["Количество лотов"]
        elif row['Тип операции'] == 'Покупка':
            cur_stocks_amount[row['Название']] = cur_stocks_amount.get(row['Название'], 0) + row["Количество лотов"]
    stocks_by_date_df = pd.DataFrame(stocks_by_date_list, columns=["Дата", "Название", "Количество лотов"])
    return stocks_by_date_df

def make_candle_charts(client: Services, figi_lst: list[str], start_date: datetime, end_date: datetime) -> None:
    candles_path = OUTPUT_PATH / 'candles'
    clear_or_create_dir(candles_path)
    for figi in figi_lst:
        get_stock_candles(client, figi, end_date, candles_path)
        time.sleep(REQUEST_DELAY_SECONDS)


def make_report(client: Services, start_date: datetime, end_date: datetime, draw_candles: bool = False) -> None:
    filename = f"{OUTPUT_PATH}/report.xlsx"
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')

    sheet_name = 'Операции'
    operations_df = get_operations_df(client, start_date, end_date)
    operations_df['Дата'] = operations_df['Дата'].apply(lambda x: x.strftime('%d.%m.%Y   %H.%M.%S'))
    operations_df.to_excel(excel_writer=writer, sheet_name=sheet_name, header=True, index=False)
    writer = format_xlsx(writer, operations_df, alignments='cllcccccc', sheet_name=sheet_name)
    writer = colorize_operations_report(writer, operations_df, sheet_name=sheet_name)

    sheet_name = 'Итог по типам операций'
    total_operations_df = operations_df.groupby(by="Тип операции", as_index=False)['Сумма'].sum()
    total_operations_df.to_excel(excel_writer=writer, sheet_name=sheet_name, header=True, index=False)
    writer = format_xlsx(writer, total_operations_df, 'lc', sheet_name=sheet_name)

    sheet_name = 'Валюта'
    money_df = get_money_df(client)
    money_df.to_excel(excel_writer=writer, sheet_name=sheet_name, header=True, index=False)
    writer = format_xlsx(writer, money_df, alignments='c' * money_df.shape[1], sheet_name=sheet_name)

    sheet_name = 'Акции'
    stocks_df = get_stocks_df(client)
    stocks_df.to_excel(excel_writer=writer, sheet_name=sheet_name, header=True, index=False)
    writer = format_xlsx(writer, stocks_df, alignments='c' * stocks_df.shape[1], sheet_name=sheet_name)

    sheet_name = 'Итог по компаниям'
    operations_by_companies_df = operations_df[operations_df['Название'].str.strip().str.len() > 0].\
                            groupby(by="Название", as_index=False)['Сумма'].sum()
    operations_by_companies_df['Сумма'] = operations_by_companies_df['Сумма'].apply(lambda x: round(float(x), 2))
    # when calculating the results, we take into account the shares in the portfolio
    stock_values_lookup = stocks_df.set_index('Название')['Общая стоимость акций'].to_dict()
    for index, row in operations_by_companies_df.iterrows():
        company_name = row['Название']
        if company_name in stock_values_lookup:
            stock_value_to_add = stock_values_lookup[company_name]
            operations_by_companies_df.loc[index, 'Сумма'] += stock_value_to_add
    operations_by_companies_df.sort_values(by='Сумма', ascending=False, inplace=True)
    operations_by_companies_df.to_excel(excel_writer=writer, sheet_name=sheet_name, header=True, index=False)
    writer = format_xlsx(writer, operations_by_companies_df, 'lc', sheet_name=sheet_name)
    writer = colorize_companies_report(writer, operations_by_companies_df, sheet_name=sheet_name)

    # stocks_by_date_df = get_stocks_by_date(operations_df, stocks_df)
    # stocks_by_date_df.to_excel(excel_writer=writer, sheet_name='Портфель на дату', header=True, index=False)
    # writer = format_xlsx(writer, stocks_df, alignments='c' * stocks_by_date_df.shape[1], sheet_name='Портфель на дату')
    writer.close()

    if draw_candles:
        # collection of tickers as the sum of tickers from the favorites and tickers for which there were operations
        figi_lst = [instrument['figi'] for instrument in get_favourite_instruments(client)]
        for figi in operations_df['FIGI'].dropna().unique():
            if figi != '' and figi not in figi:
                figi_lst.append(figi)
        make_candle_charts(client, figi_lst, start_date, end_date)


def main() -> None:
    if TOKEN:
        with Client(TOKEN, target=INVEST_GRPC_API) as client:
            clear_or_create_dir(OUTPUT_PATH)
            start_date = datetime(2024, 10, 1, 0, 0, 0).replace(tzinfo=timezone.utc)
            # start_date = datetime(2025, 1, 1, 0, 0, 0).replace(tzinfo=timezone.utc)
            # end_date = datetime(2024, 11, 1, 0, 0, 0).replace(tzinfo=timezone.utc)
            end_date = datetime.now().replace(tzinfo=timezone.utc)
            make_report(client, start_date, end_date, True)
    else:
        print('Token not found')


if __name__ == '__main__':
    start_time = time.perf_counter()
    print("Started...")
    main()
    elapsed_time = time.perf_counter() - start_time
    print(f"Done. Elapsed time {elapsed_time:.1f} seconds")