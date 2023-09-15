import os
import shutil
import pandas as pd


def clear_dir(path):
    """
    Function for recursively clearing a directory (removing all nested files)

    Args:
        param path (str): directory path
    """
    for entry in os.scandir(path):
        if not entry.name.startswith('.') and entry.is_file():
            os.unlink(entry.path)
        elif not entry.name.startswith('.') and entry.is_dir():
            shutil.rmtree(entry.path)


def create_or_clean_dir(path):
    """
    Function to create an empty directory, if the directory exists it is cleared

    Args:
        param path (str): directory path
    """
    if os.path.exists(path):
        clear_dir(path)
    else:
        os.mkdir(path)


def format_xlsx(writer: pd.ExcelWriter, df: pd.DataFrame, alignments: str, sheet_name: str = 'Sheet1',
                line_height: int = 20) -> pd.ExcelWriter:
    """
    Function for formatting an object of type XlsxWriter

    Args:
        param writer (pandas.io.excel._xlsxwriter._XlsxWriter): object of type XlsxWriter
        param df (pandas.core.frame.DataFrame): pandas dataframe with data
        param alignments (str): string indicating column alignment (r, l, c, j)
        param sheet_name (str): sheet name
        param line_height (int): cell height
    """
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    header_list = df.columns.values.tolist()
    # set column width and alignment
    a = {'l': 'left', 'r': 'right', 'c': 'center', 'j': 'justify'}
    for i in range(len(header_list)):
        cw = max([len(str(r)) for r in df[header_list[i]]])
        hw = max(len(header_list[i]), cw) + 5
        cell_format = workbook.add_format()
        cell_format.set_align(a[alignments[i]])
        worksheet.set_column(i, i, hw, cell_format)
    # set cell height
    for i in range(len(df) + 1):
        worksheet.set_row(i, line_height)
    return writer
