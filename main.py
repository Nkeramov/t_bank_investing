import os
import time
import pytz
import requests
import pandas as pd
from pathlib import Path
from decimal import Decimal, ROUND_HALF_UP
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from collections import OrderedDict
from datetime import datetime, timedelta, timezone

import mplfinance as mpf
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

from tinkoff.invest.services import Services
from tinkoff.invest.utils import now, money_to_decimal, quotation_to_decimal
from tinkoff.invest.constants import INVEST_GRPC_API
from tinkoff.invest.clients import Client
from tinkoff.invest import GetOperationsByCursorRequest, RequestError
from tinkoff.invest.schemas import OperationState, CandleInterval, TradingDay

from libs.utils import clear_or_create_dir, format_xlsx, crop_image_white_margins
from libs.report_colorize import colorize_operations_report, colorize_companies_report
from libs.grpc_schemas_descriptions import operations_types, operations_states, trade_directions
from libs.log_utils import LoggerSingleton

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

logger = LoggerSingleton(
    log_dir=Path('logs'),
    log_file="app.log",
    level="INFO",
    colored=True
).get_logger()


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
        logger.error(f"Error fetching data: {e}")
    except ValueError as e:
        logger.error(f"Error parsing data: {e}")
    except Exception as e:
        logger.error(f"An unexpected error occurred: {e}")
    return currency_rates


currencies_rates = get_cb_currencies_rate()


def get_moex_trading_schedule_by_date(client: Services, date_: datetime) -> TradingDay:
    """
        date	google.protobuf.Timestamp	Дата.
        is_trading_day	bool	Признак торгового дня на бирже.
        start_time	google.protobuf.Timestamp	Время начала торгов по UTC.
        end_time	google.protobuf.Timestamp	Время окончания торгов по UTC.
        opening_auction_start_time	google.protobuf.Timestamp	Время начала аукциона открытия по UTC.
        closing_auction_end_time	google.protobuf.Timestamp	Время окончания аукциона закрытия по UTC.
        evening_opening_auction_start_time	google.protobuf.Timestamp	Время начала аукциона открытия вечерней сессии по UTC.
        evening_start_time	google.protobuf.Timestamp	Время начала вечерней сессии по UTC.
        evening_end_time	google.protobuf.Timestamp	Время окончания вечерней сессии по UTC.
        clearing_start_time	google.protobuf.Timestamp	Время начала основного клиринга по UTC.
        clearing_end_time	google.protobuf.Timestamp	Время окончания основного клиринга по UTC.
        premarket_start_time	google.protobuf.Timestamp	Время начала премаркета по UTC.
        premarket_end_time	google.protobuf.Timestamp	Время окончания премаркета по UTC.
        closing_auction_start_time	google.protobuf.Timestamp	Время начала аукциона закрытия по UTC.
        opening_auction_end_time	google.protobuf.Timestamp	Время окончания аукциона открытия по UTC.
        intervals	Массив объектов TradingInterval	Торговые интервалы.
    """
    td = client.instruments.trading_schedules(exchange='moex', from_=date_, to=date_)
    ts = td.exchanges[0].days[0]
    return ts


def get_stock_candles(client: Services, figi: str, start_date: datetime, candles_path: str | Path) -> None:
    candles_path = Path(candles_path)
    stocks = client.instruments.shares().instruments
    stock = find_item_by_class_attr(stocks, 'figi', figi)
    if stock:
        logger.debug(f'Stock with FIGI={figi} found in shared stocks data')
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
                'axes.titlesize': 12,
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
        cd1 = (cd2 - timedelta(hours=8))
        candles_list = []
        for candle in client.get_all_candles(
                figi=figi,
                from_=cd1, to=cd2,
                interval=CandleInterval.CANDLE_INTERVAL_5_MIN,
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
            title = f"\n\n\n\n{stock.ticker} {stock.name} {datetime_formatter(cd1)} - {datetime_formatter(cd2)}"
            fig, axes = mpf.plot(candles_df, type='candle', datetime_format='%d.%m  %H:%M', figratio=(16, 9),
                                 volume=False, returnfig=True, xlabel= 'Time', ylabel='Price', xrotation=45,
                                 show_nontrading=True, style=style, axtitle=title, tight_layout=True,)
            ax = axes[0]
            ax.xaxis.set_major_locator(mdates.MinuteLocator(interval=10))
            filename = candles_path / f'{stock.ticker}.png'
            fig.savefig(filename, dpi=600, bbox_inches='tight' )
            plt.close()
            crop_image_white_margins(old_filename=filename, new_filename= filename)
            logger.debug(f"Successfully saved candles for stock with FIGI={stock.name}, Name={stock.name}")
        else:
            logger.warning(f'Candles for stock with FIGI={stock.name}, Name={stock.name} not found')
    else:
        logger.warning(f'Stock with FIGI={figi} not found in shared stocks data')


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
        try:
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
                    "Сумма": money_to_decimal(operation.payment),
                    "Цена лота": quotation_to_decimal(operation.price) if operations_types[operation.type].startswith(('Продажа', 'Покупка')) else None,
                    "Количество лотов": operation.quantity if operations_types[operation.type].startswith(('Продажа', 'Покупка')) else None,
                    "Статус": operations_states[operation.state]
                })
            time.sleep(REQUEST_DELAY_SECONDS)
            request = get_request(cursor=operations.next_cursor)
            operations = client.operations.get_operations_by_cursor(request)
        except RequestError as e:
            logger.error('Waiting for rate limit reset for ', str(e.metadata.ratelimit_reset + 3), ' sec.')
            time.sleep(e.metadata.ratelimit_reset + 3)
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
        currency_balance = money_to_decimal(money)
        currency = currencies_rates[currency_name]
        currency_balance = (currency_balance / Decimal(currency['nominal']) * Decimal(currency['rate']))
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
            current_price = money_to_decimal(stock.current_price)
            quantity = quotation_to_decimal(stock.quantity)
            quantity_lots = quotation_to_decimal(stock.quantity_lots)
            average_position_price = money_to_decimal(stock.average_position_price)
            expected_yield = quotation_to_decimal(stock.expected_yield)
            total_cost = quantity * current_price
            stocks_list.append({
                "Название": r[0].name if len(r) > 0 else '',
                "FIGI": stock.figi,
                "Тикер": r[0].ticker if len(r) > 0 else '',
                "UID": stock.instrument_uid,
                "Текущая цена акции": current_price,
                "Количество акций": quantity,
                "Количество лотов": quantity_lots,
                "Общая стоимость акций": total_cost,
                "Средневзвешенная цена акции": average_position_price,
                "Доход за все время": expected_yield
            })
    stocks_df = pd.DataFrame(stocks_list, columns=["Название", "FIGI", "Тикер", "UID", "Количество акций",
                                                  "Количество лотов", "Текущая цена акции",
                                                  "Средневзвешенная цена акции", "Общая стоимость акций",
                                                  "Доход за все время"])
    return stocks_df


def get_last_trades(client: Services, figi: str, start_date: datetime, end_date: datetime) -> None:
    trades = client.market_data.get_last_trades(figi=figi, from_=start_date, to=end_date)
    for trade in trades.trades:
        logger.info(f'{trade_directions[trade.direction]}, количество: {trade.quantity}, цена {quotation_to_decimal(trade.price)}')


def get_orders(client: Services, figi: str, depth: int) -> None:
    orders = client.market_data.get_order_book(figi=figi, depth=depth)
    # Множество пар значений на покупку
    for bid in orders.bids:
        logger.info(f'Покупка, количество: {bid.quantity}, цена {quotation_to_decimal(bid.price)}')
    # Множество пар значений на продажу
    for ask in orders.asks:
        logger.info(f'Продажа, количество: {ask.quantity}, цена {quotation_to_decimal(ask.price)}')


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
    d = now().date()
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
    def decimal_to_float_6_decimals(value):
        if isinstance(value, Decimal):
            return float(value.quantize(Decimal('0.000001'), rounding=ROUND_HALF_UP))
        return value

    def round_dataframe_with_decimals(df: pd.DataFrame) -> pd.DataFrame:
        df_to_save = df.copy()

        for column in df_to_save.columns:
            if df_to_save[column].apply(lambda x: isinstance(x, Decimal)).any():
                df_to_save[column] = df_to_save[column].apply(decimal_to_float_6_decimals)
        return df_to_save

    filename = f"{OUTPUT_PATH}/report.xlsx"
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')

    sheet_name = 'Операции'
    operations_df = get_operations_df(client, start_date, end_date)
    operations_df['Дата'] = operations_df['Дата'].apply(lambda x: x.strftime('%d.%m.%Y   %H.%M.%S'))
    operations_float_df = round_dataframe_with_decimals(operations_df)
    operations_float_df.to_excel(excel_writer=writer, sheet_name=sheet_name, header=True, index=False, float_format='%.6f')
    writer = format_xlsx(writer, operations_float_df, alignments='cllcccccc', sheet_name=sheet_name)
    writer = colorize_operations_report(writer, operations_float_df, sheet_name=sheet_name)

    sheet_name = 'Акции'
    stocks_df = get_stocks_df(client)
    stocks_float_df = round_dataframe_with_decimals(stocks_df)
    stocks_float_df.to_excel(excel_writer=writer, sheet_name=sheet_name, header=True, index=False, float_format='%.6f')
    writer = format_xlsx(writer, stocks_float_df, alignments='c' * stocks_float_df.shape[1], sheet_name=sheet_name)

    sheet_name = 'Валюта'
    money_df = get_money_df(client)
    money_float_df = round_dataframe_with_decimals(money_df)
    money_float_df.to_excel(excel_writer=writer, sheet_name=sheet_name, header=True, index=False, float_format='%.6f')
    writer = format_xlsx(writer, money_float_df, alignments='c' * money_float_df.shape[1], sheet_name=sheet_name)

    sheet_name = 'Итог по типам операций'
    total_operations_df = pd.DataFrame(operations_df.groupby(by="Тип операции", as_index=False)['Сумма'].sum())
    total_operations_float_df = round_dataframe_with_decimals(total_operations_df)
    total_operations_float_df.to_excel(excel_writer=writer, sheet_name=sheet_name, header=True, index=False, float_format='%.6f')
    writer = format_xlsx(writer, total_operations_float_df, 'lc', sheet_name=sheet_name)

    sheet_name = 'Итог по компаниям'

    companies = operations_df['Название'].unique()
    total_by_companies = []
    for company in companies:
        total_by_companies.append({
            'Название': company,
            'Сумма': operations_df[operations_df['Название'] == company]['Сумма'].sum()
        })
    total_by_companies_df = pd.DataFrame(total_by_companies, columns=['Название', 'Сумма'])
    total_by_companies_df.sort_values(by='Сумма', ascending=False, inplace=True)
    total_by_companies_float_df = round_dataframe_with_decimals(total_by_companies_df)
    total_by_companies_float_df.to_excel(excel_writer=writer, sheet_name=sheet_name, header=True, index=False, float_format='%.6f')
    writer = format_xlsx(writer, total_by_companies_float_df, 'lc', sheet_name=sheet_name)
    writer = colorize_companies_report(writer, total_by_companies_float_df, sheet_name=sheet_name)

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
            td = get_moex_trading_schedule_by_date(client, now())
            tz_changer = lambda x: x.astimezone(tz=pytz.timezone(TIMEZONE))
            dt_formatter = lambda x:  tz_changer(x).strftime("%H:%M:%S") if tz_changer(x).timestamp() > 0 else 'Неизвестно'
            logger.info(f'Дата: {tz_changer(td.date).date()}')
            logger.info(f'Признак торгового дня на бирже: {"Да" if td.is_trading_day else "Нет"}')

            logger.info(f'Время начала торгов: {dt_formatter(td.start_time)}')
            logger.info(f'Время окончания торгов: {dt_formatter(td.end_time)}')

            logger.info(f'Время начала аукциона открытия: {dt_formatter(td.opening_auction_start_time)}')
            logger.info(f'Время окончания аукциона открытия: {dt_formatter(td.opening_auction_end_time)}')

            logger.info(f'Время начала вечерней сессии: {dt_formatter(td.evening_start_time)}')
            logger.info(f'Время окончания вечерней сессии: {dt_formatter(td.evening_end_time)}')

            logger.info(f'Время начала основного клиринга: {dt_formatter(td.clearing_start_time)}')
            logger.info(f'Время окончания основного клиринга: {dt_formatter(td.clearing_end_time)}')

            logger.info(f'Время начала премаркета: {dt_formatter(td.premarket_start_time)}')
            logger.info(f'Время окончания премаркета: {dt_formatter(td.premarket_end_time)}')

            logger.info(f'Время начала аукциона закрытия: {dt_formatter(td.closing_auction_start_time)}')
            logger.info(f'Время окончания аукциона закрытия: {dt_formatter(td.closing_auction_end_time)}')

            clear_or_create_dir(OUTPUT_PATH)
            start_date = datetime(2024, 10, 1, 0, 0, 0).replace(tzinfo=timezone.utc)
            # start_date = datetime(2025, 1, 1, 0, 0, 0).replace(tzinfo=timezone.utc)
            # end_date = datetime(2024, 11, 1, 0, 0, 0).replace(tzinfo=timezone.utc)
            end_date = now()
            make_report(client, start_date, end_date, False)
    else:
        logger.error('Token not found')


if __name__ == '__main__':
    start_time = time.perf_counter()
    logger.info("Started...")
    main()
    elapsed_time = time.perf_counter() - start_time
    logger.info(f"Done. Elapsed time {elapsed_time:.1f} seconds")