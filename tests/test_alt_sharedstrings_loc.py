# Portions Copyright (C) 2010, Manfred Moitzi under a BSD licence
from xlrd import open_workbook
from xlrd.book import Book

from .base import from_this_dir


def test_open_workbook():
    book = open_workbook(from_this_dir('sharedstrings_alt_location.xlsx'))
    # Without the handling of the alternate location for the sharedStrings.xml file, this would pop.
    assert isinstance(book, Book)
