import pytest
import xlrd

from .base import from_this_dir


def test_not_corrupted():
    with pytest.raises(Exception) as context:
        xlrd.open_workbook(from_this_dir('corrupted_error.xls'))

    assert 'Workbook corruption' in str(context)

    xlrd.open_workbook(from_this_dir('corrupted_error.xls'), ignore_workbook_corruption=True)
