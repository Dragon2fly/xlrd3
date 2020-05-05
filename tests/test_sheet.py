# Portions Copyright (C) 2010, Manfred Moitzi under a BSD licence

import types
import pytest

import xlrd
from xlrd.timemachine import xrange

from .base import from_this_dir

SHEETINDEX = 0
NROWS = 15
NCOLS = 13

ROW_ERR = NROWS + 10
COL_ERR = NCOLS + 10

sheetnames = ['PROFILEDEF', 'AXISDEF', 'TRAVERSALCHAINAGE', 'AXISDATUMLEVELS', 'PROFILELEVELS']

book = xlrd.open_workbook(from_this_dir('profiles.xls'), formatting_info=True)


def check_sheet_function(function):
    assert function(0, 0) 
    assert function(NROWS - 1, NCOLS - 1) 


def check_sheet_function_index_error(function):
    pytest.raises(IndexError, function, ROW_ERR, 0)
    pytest.raises(IndexError, function, 0, COL_ERR)


def check_col_slice(col_function):
    _slice = col_function(0, 2, NROWS - 2)
    assert len(_slice) == NROWS - 4


def check_row_slice(row_function):
    _slice = row_function(0, 2, NCOLS - 2)
    assert len(_slice) == NCOLS - 4


def test_nrows():
    sheet = book.sheet_by_index(SHEETINDEX)
    assert sheet.nrows == NROWS


def test_ncols():
    sheet = book.sheet_by_index(SHEETINDEX)
    assert sheet.ncols == NCOLS


def test_cell():
    sheet = book.sheet_by_index(SHEETINDEX)
    assert xlrd.empty_cell != sheet.cell(0, 0)
    assert xlrd.empty_cell != sheet.cell(NROWS - 1, NCOLS - 1)


def test_cell_error():
    sheet = book.sheet_by_index(SHEETINDEX)
    check_sheet_function_index_error(sheet.cell)


def test_cell_type():
    sheet = book.sheet_by_index(SHEETINDEX)
    check_sheet_function(sheet.cell_type)


def test_cell_type_error():
    sheet = book.sheet_by_index(SHEETINDEX)
    check_sheet_function_index_error(sheet.cell_type)


def test_cell_value():
    sheet = book.sheet_by_index(SHEETINDEX)
    check_sheet_function(sheet.cell_value)


def test_cell_value_error():
    sheet = book.sheet_by_index(SHEETINDEX)
    check_sheet_function_index_error(sheet.cell_value)


def test_cell_xf_index():
    sheet = book.sheet_by_index(SHEETINDEX)
    check_sheet_function(sheet.cell_xf_index)


def test_cell_xf_index_error():
    sheet = book.sheet_by_index(SHEETINDEX)
    check_sheet_function_index_error(sheet.cell_xf_index)


def test_col():
    sheet = book.sheet_by_index(SHEETINDEX)
    col = sheet.col(0)
    assert len(col) == NROWS


def test_row():
    sheet = book.sheet_by_index(SHEETINDEX)
    row = sheet.row(0)
    assert len(row) == NCOLS


def test_getitem_int():
    sheet = book.sheet_by_index(SHEETINDEX)
    row = sheet[0]
    assert len(row) == NCOLS


def test_getitem_tuple():
    sheet = book.sheet_by_index(SHEETINDEX)
    assert xlrd.empty_cell != sheet[0, 0]
    assert xlrd.empty_cell != sheet[NROWS - 1, NCOLS - 1]


def test_getitem_failure():
    sheet = book.sheet_by_index(SHEETINDEX)
    with pytest.raises(ValueError):
        sheet[0, 0, 0]

    with pytest.raises(TypeError):
        sheet["hi"]


def test_get_rows():
    sheet = book.sheet_by_index(SHEETINDEX)
    rows = sheet.get_rows()
    assert isinstance(rows, types.GeneratorType)
    assert len(list(rows)) == sheet.nrows


def test_iter():
    sheet = book.sheet_by_index(SHEETINDEX)
    rows = []
    # check syntax
    for row in sheet:
        rows.append(row)
    assert len(rows) == sheet.nrows


def test_col_slice():
    sheet = book.sheet_by_index(SHEETINDEX)
    check_col_slice(sheet.col_slice)


def test_col_types():
    sheet = book.sheet_by_index(SHEETINDEX)
    check_col_slice(sheet.col_types)


def test_col_values():
    sheet = book.sheet_by_index(SHEETINDEX)
    check_col_slice(sheet.col_values)


def test_row_slice():
    sheet = book.sheet_by_index(SHEETINDEX)
    check_row_slice(sheet.row_slice)


def test_row_types():
    sheet = book.sheet_by_index(SHEETINDEX)
    check_row_slice(sheet.col_types)


def test_row_values():
    sheet = book.sheet_by_index(SHEETINDEX)
    check_col_slice(sheet.row_values)


def test_read_ragged():
    book = xlrd.open_workbook(from_this_dir('ragged.xls'), ragged_rows=True)
    sheet = book.sheet_by_index(0)
    assert sheet.row_len(0) == 3
    assert sheet.row_len(1) == 2
    assert sheet.row_len(2) == 1
    assert sheet.row_len(3) == 4
    assert sheet.row_len(4) == 4


def test_tidy_dimensions():
    book = xlrd.open_workbook(from_this_dir('merged_cells.xlsx'))
    for sheet in book.sheets():
        for rowx in xrange(sheet.nrows):
            assert sheet.row_len(rowx) == sheet.ncols
