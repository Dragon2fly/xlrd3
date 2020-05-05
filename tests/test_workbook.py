# Portions Copyright (C) 2010, Manfred Moitzi under a BSD licence

import xlrd
from xlrd import open_workbook
from xlrd.book import Book
from xlrd.sheet import Sheet

from .base import from_this_dir

SHEETINDEX = 0
NROWS = 15
NCOLS = 13

sheetnames = ['PROFILEDEF', 'AXISDEF', 'TRAVERSALCHAINAGE', 'AXISDATUMLEVELS', 'PROFILELEVELS']
book = open_workbook(from_this_dir('profiles.xls'))


def test_open_workbook():
    assert isinstance(book, Book)


def test_nsheets():
    assert book.nsheets == 5


def test_sheet_by_name():
    for name in sheetnames:
        sheet = book.sheet_by_name(name)
        assert isinstance(sheet, Sheet)
        assert name == sheet.name


def test_sheet_by_index():
    for index in range(5):
        sheet = book.sheet_by_index(index)
        assert isinstance(sheet, Sheet)
        assert sheet.name == sheetnames[index]


def test_sheets():
    sheets = book.sheets()
    for index, sheet in enumerate(sheets):
        assert isinstance(sheet, Sheet)
        assert sheet.name == sheetnames[index]


def test_sheet_names():
    assert sheetnames == book.sheet_names()


def test_getitem_ix():
    sheet = book[SHEETINDEX]
    assert xlrd.empty_cell != sheet.cell(0, 0)
    assert xlrd.empty_cell != sheet.cell(NROWS - 1, NCOLS - 1)


def test_getitem_name():
    sheet = book[sheetnames[SHEETINDEX]]
    assert xlrd.empty_cell != sheet.cell(0, 0)
    assert xlrd.empty_cell != sheet.cell(NROWS - 1, NCOLS - 1)


def test_iter():
    sheets = [sh.name for sh in book]
    assert sheets == sheetnames
