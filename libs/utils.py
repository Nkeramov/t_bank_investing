import cv2
import numpy as np
import pandas as pd
from typing import Any
from pathlib import Path
from numpy import ndarray, dtype
from decimal import Decimal, ROUND_HALF_UP


def recursive_rmdir(directory: str | Path, remove_root: bool = True) -> None:
    """
    Function for recursively clearing a directory (removing all nested files and dirs)

    Args:
        directory: path to the directory that will be cleared
        remove_root: True if needed to delete the root directory too, else False
    """
    try:
        path = Path(directory)
        if path.is_dir():
            for entry in path.iterdir():
                if entry.is_file():
                    entry.unlink()
                else:
                    recursive_rmdir(entry)
        if remove_root:
            path.rmdir()
    except PermissionError as e:
        print(f"Insufficient rights to delete. Error message: {e}")
    except FileNotFoundError:
        print(f"File not found: {directory}")


def clear_dir(directory: str | Path) -> None:
    """
    Function for recursively clearing a directory (removing all nested files and dirs)

    Args:
        directory: path to the directory that will be cleared
    """
    recursive_rmdir(directory, False)


def clear_or_create_dir(directory: str | Path) -> None:
    """
    Function to clear a directory or create a new one if the specified directory does not exist

    Args:
        directory: path to the directory that will be cleared or created
    """
    try:
        path = Path(directory)
        if path.is_dir():
            clear_dir(path)
        else:
            path.mkdir(parents=False, exist_ok=True)
    except PermissionError as e:
        print(f"Insufficient rights to delete. Error message: {e}")
    except FileNotFoundError:
        print(f"File not found: {directory}")


def crop_image_white_margins(old_filename: str | Path, xpadding: int = 15, ypadding: int = 15,
                             new_filename: str | Path | None = None) -> None:
    """
    Function for cropping images with graphs (white margins at the edges are cut off).
    If a new filename is not passed, the original file will be overwritten

    Args:
        old_filename: number of bytes
        xpadding: horizontal padding
        ypadding: vertical padding
        new_filename: path to the new (cropped) image
    """
    try:
        img = cv2.imread(str(old_filename))
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        gray = 255 * (gray < 128).astype(np.uint8)
        cords = cv2.findNonZero(gray)
        x, y, w, h = cv2.boundingRect(cords)
        rect = img[y - ypadding: y + h + 2 * ypadding, x - xpadding: x + w + 2 * xpadding]
        is_success, im_buf_arr = cv2.imencode(".png", rect)
        if new_filename is None:
            Path(old_filename).unlink()
            im_buf_arr.tofile(old_filename)
        else:
            im_buf_arr.tofile(new_filename)
    except FileNotFoundError as e:
        print(f"File not found error: {e}")
    except ValueError as e:
        print(f"Image reading error: {e}")
    except IOError as e:
        print(f"IO error: {e}")


def format_xlsx(writer: pd.ExcelWriter, df: pd.DataFrame, sheet_name: str = 'Sheet1',
                alignments: str | None = None, font_size: int | None = None, border_width: int | None = None,
                border_color: str | None = None, cell_height: int = 20) -> pd.ExcelWriter:
    """
    Function for formatting an object of XlsxWriter type.
    Allows to set alignment for each column, adjust cells height, border color, border width and font size.

    Args:
        writer: object of XlsxWriter type
        df: pandas dataframe with data
        sheet_name: name of the sheet to be formatted
        alignments: string indicating columns alignments (r, l, c, j), default is left alignment for all columns
        font_size: font size for all cells
        border_width: border width for all cells
        border_color: border color for all cells
        cell_height: cell height for all cells
    """
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    if alignments is None:
        alignments = 'l' * df.shape[1]
    # set column width and alignment
    a = {'l': 'left', 'r': 'right', 'c': 'center', 'j': 'justify'}
    for col_index, col_name in enumerate(df.columns):
        col_width = len(col_name)
        if df.shape[0] > 0:
            col_width = max(col_width, max(len(str(r)) for r in df[col_name]))
        col_width = round(col_width * 1.2)
        cell_format = workbook.add_format()
        cell_format.set_align(a[alignments[col_index]])
        if font_size:
            cell_format.set_font_size(font_size)
        if border_width:
            cell_format.set_border(border_width)
        if border_color:
            cell_format.set_border_color(border_color)
        worksheet.set_column(col_index, col_index, col_width, cell_format)
    # set cells height
    for i in range(len(df) + 1):
        worksheet.set_row(i, cell_height)
    return writer


def get_entity_with_case(x, word_forms: tuple[str, str, str] = ('день', 'дня', 'дней')) -> str:
    """
    Function for determining the declension of a noun following a numeral

    Args:
        x: numeric value
        word_forms: tuple of string declension values (for 1, for 2-4, for 5-0)
    """
    n = abs(x) % 100
    if 10 < n < 20:
        return word_forms[2]
    n1 = n % 10
    if n1 == 1:
        return word_forms[0]
    elif 1 < n1 < 5:
        return word_forms[1]
    else:
        return word_forms[2]


def decimal_to_float_n_decimals(value: Decimal, decimals: int = 6) -> float:
    """
    Convert a Decimal value to float with specified decimal precision.
    Rounds the Decimal value to the specified number of decimals using
    ROUND_HALF_UP rounding mode before converting to float.

    Args:
        value: decimal value to convert
        decimals: number of decimal places to keep. Defaults to 6.

    Returns:
        float: rounded float value, or original value if not a Decimal
    """
    if isinstance(value, Decimal):
        value = value.quantize(Decimal(f'10e-{decimals - 1}'), rounding=ROUND_HALF_UP)
    return float(value)


def round_dataframe_with_decimals(df: pd.DataFrame, decimals: int = 6) -> pd.DataFrame:
    """
    Round Decimal columns in a DataFrame to specified decimal precision.
    Creates a copy of the DataFrame and converts all Decimal columns to float
    with the specified number of decimal places using half-up rounding.

    Args:
        df: dataFrame to process
        decimals: number of decimal places for rounding. Defaults to 6.

    Returns:
        pd.DataFrame: new DataFrame with Decimal columns converted to rounded floats
    """
    df_to_save = df.copy()
    for column in df_to_save.columns:
        if df_to_save[column].apply(lambda x: isinstance(x, Decimal)).any():
            df_to_save[column] = df_to_save[column].apply(lambda x: decimal_to_float_n_decimals(x, decimals))
    return df_to_save


def get_unique_non_empty(series: pd.Series) -> ndarray[tuple[Any, ...], dtype[Any]]:
    """
    Get unique non-empty values from a pandas Series.
    Filters out NaN/None values and empty strings (after stripping whitespace),
    then returns the unique values from the remaining elements.

    Args:
        series: series to extract unique values from

    Returns:
        ndarray: array of unique non-empty values from the Series
    """
    return series[series.notna() & (series.str.strip().str.len() > 0)].unique()


def find_item_by_dict_key(lst: list, key: str, value: Any) -> Any:
    """
    Find an item in a list of dictionaries by key value.
    Searches through a list of dictionaries and returns the first dictionary
    that contains the specified key with the matching value.

    Args:
        lst: list of dictionaries to search through
        key: dictionary key to check
        value: value to match against the specified key

    Returns:
        Any: matching dictionary item, or None if no match found
    """
    return next((item for item in lst if item.get(key) == value), None)


def find_item_by_class_attr(lst: list, attr_name: str, value: Any) -> Any:
    """
    Find an item in a list of class instances by attribute value.
    Searches through a list of objects and returns the first object
    that has the specified attribute with the matching value.

    Args:
        lst: list of class instances to search through
        attr_name: attribute name to check
        value: value to match against the specified attribute

    Returns:
        Any: matching class instance, or None if no match found
    """
    return next((item for item in lst if getattr(item, attr_name, None) == value), None)