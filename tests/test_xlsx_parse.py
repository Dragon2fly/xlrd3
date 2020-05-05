###############################################################################
#
# Test the parsing of problematic xlsx files from bug reports.
#

import xlrd

from .base import from_this_dir


# Test parsing of problematic xlsx files. These are usually submitted
# as part of bug reports as noted below.

def test_for_github_issue_75():
    # Test <cell> inlineStr attribute without <si> child.
    # https://github.com/python-excel/xlrd/issues/75
    workbook = xlrd.open_workbook(from_this_dir('apachepoi_52348.xlsx'))
    worksheet = workbook.sheet_by_index(0)

    # Test an empty inlineStr cell.
    cell = worksheet.cell(0, 0)
    assert cell.value == ''
    assert cell.ctype == xlrd.book.XL_CELL_EMPTY

    # Test a non-empty inlineStr cell.
    cell = worksheet.cell(1, 2)
    assert cell.value == 'Category'
    assert cell.ctype == xlrd.book.XL_CELL_TEXT


def test_for_github_issue_96():
    # Test for non-Excel file with forward slash file separator and
    # lowercase names. https://github.com/python-excel/xlrd/issues/96
    workbook = xlrd.open_workbook(from_this_dir('apachepoi_49609.xlsx'))
    worksheet = workbook.sheet_by_index(0)

    # Test reading sample data from the worksheet.
    cell = worksheet.cell(0, 1)
    assert cell.value == 'Cycle'
    assert cell.ctype == xlrd.book.XL_CELL_TEXT

    cell = worksheet.cell(1, 1)
    assert cell.value == 1
    assert cell.ctype == xlrd.book.XL_CELL_NUMBER


def test_for_github_issue_101():
    # Test for non-Excel file with forward slash file separator
    # https://github.com/python-excel/xlrd/issues/101
    workbook = xlrd.open_workbook(from_this_dir('self_evaluation_report_2014-05-19.xlsx'))
    worksheet = workbook.sheet_by_index(0)

    # Test reading sample data from the worksheet.
    cell = worksheet.cell(0, 0)
    assert cell.value == 'one'
    assert cell.ctype == xlrd.book.XL_CELL_TEXT


def test_for_github_issue_150():
    # Test for non-Excel file with a non-lowercase worksheet filename.
    # https://github.com/python-excel/xlrd/issues/150
    workbook = xlrd.open_workbook(from_this_dir('issue150.xlsx'))
    worksheet = workbook.sheet_by_index(0)

    # Test reading sample data from the worksheet.
    cell = worksheet.cell(0, 1)
    assert cell.value == 'Cycle'
    assert cell.ctype == xlrd.book.XL_CELL_TEXT

    cell = worksheet.cell(1, 1)
    assert cell.value == 1
    assert cell.ctype == xlrd.book.XL_CELL_NUMBER
