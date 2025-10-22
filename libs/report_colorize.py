import pandas as pd
from .grpc_schemas import OperationType, operations_types

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