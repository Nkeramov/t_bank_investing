import os
import cv2
import pandas as pd
import numpy as np
from pathlib import Path

def recursive_rmdir(directory: str | Path) -> None:
    """
    Function for recursively clearing a directory (removing all nested files and dirs)

    Args:
        directory: path to the directory that will be cleared
    """
    try:
        path = Path(directory)
        if path.is_dir():
            for entry in path.iterdir():
                if entry.is_file():
                    entry.unlink()
                else:
                    recursive_rmdir(entry)
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
    try:
        path = Path(directory)
        if path.is_dir():
            for entry in path.iterdir():
                if entry.is_file():
                    entry.unlink()
                else:
                    recursive_rmdir(entry)
    except PermissionError as e:
        print(f"Insufficient rights to delete. Error message: {e}")
    except FileNotFoundError:
        print(f"File not found: {directory}")


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


def format_xlsx(writer: pd.ExcelWriter, df: pd.DataFrame, alignments: str | None = None,
                sheet_name: str = 'Sheet1', cell_height: int = 20) -> pd.ExcelWriter:
    """
    Function for formatting an object of XlsxWriter type.
    Allows to set alignment for each column and adjust cells height.

    Args:
        writer: object of XlsxWriter type
        df: pandas dataframe with data
        alignments: string indicating columns alignments (r, l, c, j), default is left alignment for all columns
        sheet_name: name of the sheet to be formatted
        cell_height: cell height
    """
    if df.shape[0] > 0:
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        if alignments is None:
            alignments = 'l' * df.shape[1]
        # set column width and alignment
        a = {'l': 'left', 'r': 'right', 'c': 'center', 'j': 'justify'}
        for col_index, col_name in enumerate(df.columns):
            col_width = max(len(col_name), max(len(str(r)) for r in df[col_name])) + 5
            cell_format = workbook.add_format()
            # cell_format = workbook.add_format({'font_size': 12})
            cell_format.set_align(a[alignments[col_index]])
            # cell_format.set_border_color('#000000')
            # cell_format.set_border(1)
            worksheet.set_column(col_index, col_index, col_width, cell_format)
        # set cells height
        for i in range(len(df) + 1):
            worksheet.set_row(i, cell_height)
    return writer


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
    img = cv2.imread(old_filename)
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