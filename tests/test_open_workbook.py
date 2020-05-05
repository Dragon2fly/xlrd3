import os
import tempfile
import pytest

from xlrd import open_workbook

from .base import from_this_dir


# test different uses of open_workbook

def test_names_demo():
    # For now, we just check this doesn't raise an error.
    open_workbook(from_this_dir(os.path.join('..', 'examples', 'namesdemo.xls')))


def test_tilde_path_expansion():
    with tempfile.NamedTemporaryFile(suffix='.xlsx', dir=os.path.expanduser('~')) as fp:
        with open(from_this_dir('text_bar.xlsx'), 'rb') as fo:
            fp.write(fo.read())

        # For now, we just check this doesn't raise an error.
        with pytest.raises(PermissionError):
            open_workbook(os.path.join('~', os.path.basename(fp.name)))


def test_ragged_rows_tidied_with_formatting():
    # For now, we just check this doesn't raise an error.
    open_workbook(from_this_dir('issue20.xls'), formatting_info=True)


def test_BYTES_X00():
    # For now, we just check this doesn't raise an error.
    open_workbook(from_this_dir('picture_in_cell.xls'), formatting_info=True)


def test_xlsx_simple():
    # For now, we just check this doesn't raise an error.
    open_workbook(from_this_dir('text_bar.xlsx'))
    # we should make assertions here that data has been
    # correctly processed.


def test_xlsx():
    # For now, we just check this doesn't raise an error.
    open_workbook(from_this_dir('reveng1.xlsx'))
    # we should make assertions here that data has been
    # correctly processed.


def test_err_cell_empty():
    # For cell with type "e" (error) but without inner 'val' tags
    open_workbook(from_this_dir('err_cell_empty.xlsx'))


def test_xlsx_lower_case_cellnames():
    # Check if it opens with lower cell names
    open_workbook(from_this_dir('test_lower_case_cellnames.xlsx'))
