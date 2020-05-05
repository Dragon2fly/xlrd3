# Portions Copyright (C) 2010, Manfred Moitzi under a BSD licence

import pytest
import xlrd
from xlrd.timemachine import UNICODE_LITERAL

from .base import from_this_dir


@pytest.fixture()
def book():
    bk = xlrd.open_workbook(from_this_dir('profiles.xls'), formatting_info=True)
    return bk


@pytest.fixture()
def sheet():
    bk = xlrd.open_workbook(from_this_dir('profiles.xls'), formatting_info=True)
    sh = bk.sheet_by_name('PROFILEDEF')
    return sh


def test_empty_cell(book):
    sheet = book.sheet_by_name('TRAVERSALCHAINAGE')
    cell = sheet.cell(0, 0)
    assert cell.ctype == xlrd.book.XL_CELL_EMPTY
    assert cell.value == ''
    assert type(cell.value) == type(UNICODE_LITERAL(''))
    assert cell.xf_index > 0


def test_string_cell(sheet):
    cell = sheet.cell(0, 0)
    assert cell.ctype == xlrd.book.XL_CELL_TEXT
    assert cell.value == 'PROFIL'
    assert type(cell.value) == type(UNICODE_LITERAL(''))
    assert cell.xf_index > 0


def test_number_cell(sheet):
    cell = sheet.cell(1, 1)
    assert cell.ctype == xlrd.book.XL_CELL_NUMBER
    assert cell.value == 100
    assert cell.xf_index > 0


def test_calculated_cell(book):
    sheet2 = book.sheet_by_name('PROFILELEVELS')
    cell = sheet2.cell(1, 3)
    assert cell.ctype == xlrd.book.XL_CELL_NUMBER
    assert cell.value == pytest.approx(265.131, 1e-3)
    assert cell.xf_index > 0


def test_merged_cells():
    book = xlrd.open_workbook(from_this_dir('xf_class.xls'), formatting_info=True)
    sheet3 = book.sheet_by_name('table2')
    row_lo, row_hi, col_lo, col_hi = sheet3.merged_cells[0]
    assert sheet3.cell(row_lo, col_lo).value == 'MERGED'
    assert (row_lo, row_hi, col_lo, col_hi), (3, 7, 2 == 5)


def test_merged_cells_xlsx():
    book = xlrd.open_workbook(from_this_dir('merged_cells.xlsx'))

    sheet1 = book.sheet_by_name('Sheet1')
    expected = []
    got = sheet1.merged_cells
    assert expected == got

    sheet2 = book.sheet_by_name('Sheet2')
    expected = [(0, 1, 0, 2)]
    got = sheet2.merged_cells
    assert expected == got

    sheet3 = book.sheet_by_name('Sheet3')
    expected = [(0, 1, 0, 2), (0, 1, 2, 4), (1, 4, 0, 2), (1, 9, 2, 4)]
    got = sheet3.merged_cells
    assert expected == got

    sheet4 = book.sheet_by_name('Sheet4')
    expected = [(0, 1, 0, 2), (2, 20, 0, 1), (1, 6, 2, 5)]
    got = sheet4.merged_cells
    assert expected == got
