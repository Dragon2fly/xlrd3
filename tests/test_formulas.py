# Portions Copyright (C) 2010, Manfred Moitzi under a BSD licence

import xlrd
from .base import from_this_dir

book = xlrd.open_workbook(from_this_dir('formula_test_sjmachin.xls'))
sheet1 = book.sheet_by_index(0)


def get_value(sh, col, row):
    return ascii(sh.col_values(col)[row])


def test_cell_B2():
    assert get_value(sheet1, 1, 1) == r"'\u041c\u041e\u0421\u041a\u0412\u0410 \u041c\u043e\u0441\u043a\u0432\u0430'"


def test_cell_B3():
    assert get_value(sheet1, 1, 2) == '0.14285714285714285'


def test_cell_B4():
    assert get_value(sheet1, 1, 3) == "'ABCDEF'"


def test_cell_B5():
    assert get_value(sheet1, 1, 4) == "''"


def test_cell_B6():
    assert get_value(sheet1, 1, 5) == '1'


def test_cell_B7():
    assert get_value(sheet1, 1, 6) == '7'


def test_cell_B8():
    assert get_value(sheet1, 1, 7) == r"'\u041c\u041e\u0421\u041a\u0412\u0410 \u041c\u043e\u0441\u043a\u0432\u0430'"


book = xlrd.open_workbook(from_this_dir('formula_test_names.xls'))
sheet2 = book.sheet_by_index(0)


def test_unaryop():
    assert get_value(sheet2, 1, 1) == '-7.0'


def test_attrsum():
    assert get_value(sheet2, 1, 2) == '4.0'


def test_func():
    assert get_value(sheet2, 1, 3) == '6.0'


def test_func_var_args():
    assert get_value(sheet2, 1, 4) == '3.0'


def test_if():
    assert get_value(sheet2, 1, 5) == "'b'"


def test_choose():
    assert get_value(sheet2, 1, 6) == "'C'"
