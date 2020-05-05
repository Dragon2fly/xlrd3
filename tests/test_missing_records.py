from xlrd import open_workbook
from xlrd.biffh import XL_CELL_TEXT

from .base import from_this_dir

path = from_this_dir('biff4_no_format_no_window2.xls')
book = open_workbook(path)
sheet = book.sheet_by_index(0)


def test_default_format():
    cell = sheet.cell(0, 0)
    assert cell.ctype == XL_CELL_TEXT


def test_default_window2_options():
    assert sheet.cached_page_break_preview_mag_factor == 0
    assert sheet.cached_normal_view_mag_factor == 0
