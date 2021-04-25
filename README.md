<!---
[![Build Status](https://travis-ci.org/python-excel/xlrd.svg?branch=master)](https://travis-ci.org/python-excel/xlrd)
[![Coverage Status](https://coveralls.io/repos/github/python-excel/xlrd/badge.svg?branch=master)](https://coveralls.io/github/python-excel/xlrd?branch=master)
[![Documentation Status](https://readthedocs.org/projects/xlrd/badge/?version=latest)](http://xlrd.readthedocs.io/en/latest/?badge=latest)
[![PyPI version](https://badge.fury.io/py/xlrd.svg)](https://badge.fury.io/py/xlrd)
--->
### xlrd3
A fork of original archived [xlrd](https://github.com/python-excel/xlrd) project. 
This fork aims to fix bugs that existing in `xlrd` and improve it features. 
As the name of this fork implies, python2 support is dropped.   

At version 1.0.0, xlrd3 on pair with xlrd version 1.2.0 with following bugs fixed:

* MemoryError: `on_demand` with `mmap` still causes some `xls` to be read the whole file into memory.
* `on_demand` not supported for `xlsx`
* Parsing comments failed for `xlsx` on Windows platform.

### When to use xlrd3
If you just need to **read** and deal with both `xlsx` and `xls`, use `xlrd3`. 
Then if you want to export your data to other excel files, use [OpenPyXL](https://openpyxl.readthedocs.io/en/stable/) or [xlsxWriter](https://github.com/jmcnamara/XlsxWriter).
If you need to **edit** `xlsx` (read and write) and are sure that `xls` never appear in your workflow, you are advised to use [OpenPyXL](https://openpyxl.readthedocs.io/en/stable/) instead.


**Purpose**: Provide a library for developers to use to extract data from Microsoft Excel (tm) spreadsheet files. It is not an end-user tool.

**Original Author**: John Machin

**Licence**: BSD-style (see licences.py)

**Versions of Python supported**: 3.6+.

**Outside scope**: xlrd3 will safely and reliably ignore any of these if present in the file:

*   Charts, Macros, Pictures, any other embedded object. WARNING: currently this includes embedded worksheets.
*   VBA modules
*   Formulas (results of formula calculations are extracted, of course).
*   Comments
*   Hyperlinks
*   Autofilters, advanced filters, pivot tables, conditional formatting, data validation
*   Handling password-protected (encrypted) files.

**Installation**:`$pip install xlrd3`

**Quick start**:

```python
import xlrd3 as xlrd
book = xlrd.open_workbook("myfile.xls")
print("The number of worksheets is {0}".format(book.nsheets))
print("Worksheet name(s): {0}".format(book.sheet_names()))
sh = book.sheet_by_index(0)
print("{0} {1} {2}".format(sh.name, sh.nrows, sh.ncols))
print("Cell D30 is {0}".format(sh.cell_value(rowx=29, colx=3)))
for rx in range(sh.nrows):
    print(sh.row(rx))
```

**Another quick start**: This will show the first, second and last rows of each sheet in each file:

    python PYDIR/scripts/runxlrd.py 3rows *blah*.xls

**Acknowledgements**:

*   This package started life as a translation from C into Python of parts of a utility called "xlreader" developed by David Giffin. "This product includes software developed by David Giffin <david@giffin.org>."
*   OpenOffice.org has truly excellent documentation of the Microsoft Excel file formats and Compound Document file format, authored by Daniel Rentz. See http://sc.openoffice.org
*   U+5F20 U+654F: over a decade of inspiration, support, and interesting decoding opportunities.
*   Ksenia Marasanova: sample Macintosh and non-Latin1 files, alpha testing
*   Backporting to Python 2.1 was partially funded by Journyx - provider of timesheet and project accounting solutions (http://journyx.com/).
*   Provision of formatting information in version 0.6.1 was funded by Simplistix Ltd (http://www.simplistix.co.uk/)

# Changelog
## v1.1.0
* support python3.9
* refactored some underlying code for better maintenance
