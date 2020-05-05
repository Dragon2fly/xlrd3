# Portions Copyright (C) 2010, Manfred Moitzi under a BSD licence

import pytest
import xlrd

from .base import from_this_dir


@pytest.fixture()
def book():
    bk = xlrd.open_workbook(from_this_dir('Formate.xls'), formatting_info=True)
    return bk


@pytest.fixture()
def sheet():
    bk = xlrd.open_workbook(from_this_dir('Formate.xls'), formatting_info=True)
    sh = bk.sheet_by_name('Blätt1')
    return sh


def test_text_cells(sheet):
    for row, name in enumerate(['Huber', 'Äcker', 'Öcker']):
        cell = sheet.cell(row, 0)
        assert cell.ctype == xlrd.book.XL_CELL_TEXT
        assert cell.value == name
        assert cell.xf_index > 0


def test_date_cells(sheet):
    # see also 'Dates in Excel spreadsheets' in the documentation
    # convert: xldate_as_tuple(float, book.datemode) -> (year, month,
    # day, hour, minutes, seconds)
    for row, date in [(0, 2741.), (1, 38406.), (2, 32266.)]:
        cell = sheet.cell(row, 1)
        assert cell.ctype == xlrd.book.XL_CELL_DATE
        assert cell.value == date
        assert cell.xf_index > 0


def test_time_cells(sheet):
    # see also 'Dates in Excel spreadsheets' in the documentation
    # convert: xldate_as_tuple(float, book.datemode) -> (year, month,
    # day, hour, minutes, seconds)
    for row, time in [(3, .273611), (4, .538889), (5, .741123)]:
        cell = sheet.cell(row, 1)
        assert cell.ctype == xlrd.book.XL_CELL_DATE
        assert cell.value == pytest.approx(time, 1e-6)
        assert cell.xf_index > 0


def test_percent_cells(sheet):
    for row, time in [(6, .974), (7, .124)]:
        cell = sheet.cell(row, 1)
        assert cell.ctype == xlrd.book.XL_CELL_NUMBER
        assert cell.value == pytest.approx(time, 1e-3)
        assert cell.xf_index > 0


def test_currency_cells(sheet):
    for row, time in [(8, 1000.30), (9, 1.20)]:
        cell = sheet.cell(row, 1)
        assert cell.ctype == xlrd.book.XL_CELL_NUMBER
        assert cell.value == pytest.approx(time, 1e-2)
        assert cell.xf_index > 0


def test_get_from_merged_cell(book):
    sheet = book.sheet_by_name('ÖÄÜ')
    cell = sheet.cell(2, 2)
    assert cell.ctype == xlrd.book.XL_CELL_TEXT
    assert cell.value == 'MERGED CELLS'
    assert cell.xf_index > 0


def test_ignore_diagram(book):
    sheet = book.sheet_by_name('Blätt3')
    cell = sheet.cell(0, 0)
    assert cell.ctype == xlrd.book.XL_CELL_NUMBER
    assert cell.value == 100
    assert cell.xf_index > 0
