import cv2
import pandas as pd
import numpy as np
from pathlib import Path
from typing import Optional


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
    Allows to set alignment for each column and adjust cells height.

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