import os
import pandas as pd
from pytz import timezone
from datetime import datetime
from dotenv import load_dotenv

from tinkoff.invest import Client
from tinkoff.invest.constants import INVEST_GRPC_API
from tinkoff.invest.schemas import OperationType, OperationState

import utils

pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

load_dotenv()
INPUT_PATH = './input'
OUTPUT_PATH = './output'
IMG_WIDTH, IMG_HEIGHT, IMG_DPI = 3600, 2000, 150
TOKEN = os.getenv("TOKEN")

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
    'Stock': 'Акции',
    'Bond': 'Облигации',
    'Etf': 'ETF'
}

currencies_rates = {
    'RUB': ['Российский рубль', 0, 0],
    'USD': ['Доллар США', 0, 0],
    'EUR': ['Евро', 0, 0],
    'GBP': ['Фунт стерлингов', 0, 0],
    'HKD': ['Гонконгский доллар', 0, 0],
    'CHF': ['Швейцарский франк', 0, 0],
    'JPY': ['Японская иена', 0, 0],
    'CNY': ['Китайский юань', 0, 0],
    'TRY': ['Турецкая лира', 0, 0]
}


def colorize_xlsx(writer: pd.ExcelWriter, df: pd.DataFrame, sheet_name: str = 'Sheet1'):
    """
    Function for coloring cells in an XlsxWriter object based on their values

    Args:
        param writer (pandas.io.excel._xlsxwriter._XlsxWriter): object of type XlsxWriter
        param df (pandas.core.frame.DataFrame): pandas dataframe with data
        param sheet_name (str): sheet name
    """
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


if __name__ == '__main__':
    utils.create_or_clean_dir(OUTPUT_PATH)
    writer = pd.ExcelWriter(f"report.xlsx", engine='xlsxwriter')
    with Client(TOKEN, target=INVEST_GRPC_API) as client:
        account_id = client.users.get_accounts().accounts[0].id
        # Получение списка облигаций
        # bonds = client.get_market_bonds()
        # Получение списка ETF
        # etfs = client.get_market_etfs()
        # Получение списка акций
        stocks = client.instruments.shares().instruments
        start_date = datetime(2023, 1, 1, 0, 0, 0, tzinfo=timezone('Europe/Moscow'))
        end_date = datetime.now(tz=timezone('Europe/Moscow'))  # По настоящее время
        operations = client.operations.get_operations(account_id=account_id, from_=start_date, to=end_date)
        operations_list = []
        for k in operations.operations:
            dt_tz = k.date
            d = dt_tz.replace(tzinfo=None)
            r = [x for x in stocks if x.figi == k.figi]
            if k.state == OperationState.OPERATION_STATE_EXECUTED:
                operations_list.append({
                    "Дата": d,
                    "Тип операции": operations_types[k.operation_type.value],
                    "FIGI": k.figi,
                    "Тикер": r[0].ticker if len(r) > 0 else k.figi,
                    "Название": r[0].name if len(r) > 0 else k.figi,
                    "Тип": r[0].share_type if len(r) > 0 else k.figi,
                    "Сумма": round(k.payment.units + k.payment.nano / 1e9, 2),
                    "Цена лота": round(k.price.units + k.price.nano / 1e9, 2),
                    "Количество лотов": k.quantity
                })
        operations_df = pd.DataFrame(operations_list, columns=["Дата", "Тип операции", "Название", "Тикер", "FIGI",
                                                               "Цена лота", "Количество лотов", "Сумма", "Статус"])
        operations_df.sort_values(by='Дата', ascending=True, inplace=True)
        operations_df['Дата'] = operations_df['Дата'].apply(lambda x: x.strftime('%d.%m.%Y   %H.%M.%S'))
        operations_df.to_excel(excel_writer=writer, sheet_name='Операции', header=True, index=False)

        writer = utils.format_xlsx(writer, operations_df, 'cllcccccc', sheet_name='Операции')
        writer = format_xlsx(writer, operations_df, sheet_name='Операции')
        writer.close()
        # total_operations_df = operations_df.groupby(by="Тип операции", as_index=False)['Сумма'].sum()
        # total_operations_df['Сумма'] = total_operations_df['Сумма'].apply(lambda x: abs(float(x)))
        # total_operations_df.to_excel(excel_writer=writer, sheet_name='Итог', header=True, index=False)
        # writer = utils.format_xlsx(writer, total_operations_df, 'lc', sheet_name='Итог')
        # writer.save()
