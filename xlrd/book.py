# Copyright (c) 2005-2012 Stephen John Machin, Lingfo Pty Ltd
# This module is part of the xlrd package, which is released under a
# BSD-style licence.

import gc
import mmap
import struct

from time import perf_counter
from logging import getLogger

from . import compdoc, formatting, sheet
from .biffh import *
from .formula import *
from .timemachine import *


logger = getLogger(__name__)
unpack = struct.unpack

empty_cell = sheet.empty_cell  # for exposure to the world ...

USE_FANCY_CD = True
TOGGLE_GC = False
# gc.set_debug(gc.DEBUG_STATS)

USE_MMAP = True
MY_EOF = 0xF00BAAA  # not a 16-bit number
SUPBOOK_UNK, SUPBOOK_INTERNAL, SUPBOOK_EXTERNAL, SUPBOOK_ADDIN, SUPBOOK_DDEOLE = range(5)
SUPPORTED_VERSIONS = (80, 70, 50, 45, 40, 30, 21, 20)

_code_from_builtin_name = {
    "Consolidate_Area": "\x00",
    "Auto_Open": "\x01",
    "Auto_Close": "\x02",
    "Extract": "\x03",
    "Database": "\x04",
    "Criteria": "\x05",
    "Print_Area": "\x06",
    "Print_Titles": "\x07",
    "Recorder": "\x08",
    "Data_Form": "\x09",
    "Auto_Activate": "\x0A",
    "Auto_Deactivate": "\x0B",
    "Sheet_Title": "\x0C",
    "_FilterDatabase": "\x0D",
}
builtin_name_from_code = {}
code_from_builtin_name = {}
for _bin, _bic in _code_from_builtin_name.items():
    _bin = UNICODE_LITERAL(_bin)
    _bic = UNICODE_LITERAL(_bic)
    code_from_builtin_name[_bin] = _bic
    builtin_name_from_code[_bic] = _bin
del _bin, _bic, _code_from_builtin_name


def open_workbook_xls(filename=None, verbosity=0, use_mmap=USE_MMAP, file_contents=None,
                      encoding_override=None, formatting_info=False, on_demand=False, ragged_rows=False,
                      ignore_wb_corruption=False):

    t0 = perf_counter()
    if TOGGLE_GC:
        orig_gc_enabled = gc.isenabled()
        if orig_gc_enabled:
            gc.disable()
    bk = Book()
    try:
        bk.biff2_8_load(filename=filename, file_contents=file_contents, verbosity=verbosity,
                        use_mmap=use_mmap, encoding_override=encoding_override, formatting_info=formatting_info,
                        on_demand=on_demand, ragged_rows=ragged_rows, ignore_wb_corruption=ignore_wb_corruption)
        t1 = perf_counter()
        bk.load_time_stage_1 = t1 - t0
        biff_version = bk.getbof(XL_WORKBOOK_GLOBALS)
        if not biff_version:
            raise XLRDError("Can't determine file's BIFF version")
        if biff_version not in SUPPORTED_VERSIONS:
            raise XLRDError(
                "BIFF version %s is not supported"
                % biff_text_from_num[biff_version]
            )
        bk.biff_version = biff_version
        if biff_version <= 40:
            # no workbook globals, only 1 worksheet
            if on_demand:
                logger.warning("on_demand is not supported for this Excel version. Setting on_demand to False")
                bk.on_demand = on_demand = False
            bk.fake_globals_get_sheet()
        elif biff_version == 45:
            # worksheet(s) embedded in global stream
            bk.parse_globals()
            if on_demand:
                logger.warning("on_demand is not supported for this Excel version. Setting on_demand to False")
                bk.on_demand = on_demand = False
        else:
            bk.parse_globals()
            bk._sheet_list = [None for sh in bk._sheet_names]
            if not on_demand:
                bk.get_sheets()
        bk.nsheets = len(bk._sheet_list)
        if biff_version == 45 and bk.nsheets > 1:
            logger.warning(f"Excel 4.0 workbook (.XLW) file contains {bk.nsheets} worksheets. "
                           f"Book-level data will be that of the last worksheet")
        if TOGGLE_GC:
            if orig_gc_enabled:
                gc.enable()
        t2 = perf_counter()
        bk.load_time_stage_2 = t2 - t1
    except:
        bk.release_resources()
        raise
    # normal exit
    if not on_demand:
        bk.release_resources()
    return bk


class Name(BaseObject):
    """
    Information relating to a named reference, formula, macro, etc.

    .. note::

      Name information is **not** extracted from files older than
      Excel 5.0 (``Book.biff_version < 50``)
    """
    _repr_these = ['stack']
    book = None  # parent

    #: 0 = Visible; 1 = Hidden
    hidden = 0

    #: 0 = Command macro; 1 = Function macro. Relevant only if macro == 1
    func = 0

    #: 0 = Sheet macro; 1 = VisualBasic macro. Relevant only if macro == 1
    vbasic = 0

    #: 0 = Standard name; 1 = Macro name
    macro = 0

    #: 0 = Simple formula; 1 = Complex formula (array formula or user defined).
    #:
    #: .. note:: No examples have been sighted.
    complex = 0

    #: 0 = User-defined name; 1 = Built-in name
    #:
    #: Common examples: ``Print_Area``, ``Print_Titles``; see OOo docs for
    #: full list
    builtin = 0

    #: Function group. Relevant only if macro == 1; see OOo docs for values.
    funcgroup = 0

    #: 0 = Formula definition; 1 = Binary data
    #:
    #: .. note:: No examples have been sighted.
    binary = 0

    #: The index of this object in book.name_obj_list
    name_index = 0

    # A Unicode string. If builtin, decoded as per OOo docs.
    name = UNICODE_LITERAL("")

    #: An 8-bit string.
    raw_formula = b''

    #: ``-1``:
    #:    The name is global (visible in all calculation sheets).
    #: ``-2``:
    #:    The name belongs to a macro sheet or VBA sheet.
    #: ``-3``:
    #:    The name is invalid.
    #: ``0 <= scope < book.nsheets``:
    #:    The name is local to the sheet whose index is scope.
    scope = -1

    #: The result of evaluating the formula, if any.
    #: If no formula, or evaluation of the formula encountered problems,
    #: the result is ``None``. Otherwise the result is a single instance of the
    #: :class:`~xlrd.formula.Operand` class.
    #
    result = None

    def cell(self):
        """
        This is a convenience method for the frequent use case where the name
        refers to a single cell.

        :returns: An instance of the :class:`~xlrd.sheet.Cell` class.

        :raises xlrd.biffh.XLRDError:
          The name is not a constant absolute reference
          to a single cell.
        """
        res = self.result
        if res:
            # result should be an instance of the Operand class
            kind = res.kind
            value = res.value
            if kind == oREF and len(value) == 1:
                ref3d = value[0]
                if (0 <= ref3d.shtxlo == ref3d.shtxhi - 1 and
                        ref3d.rowxlo == ref3d.rowxhi - 1 and
                        ref3d.colxlo == ref3d.colxhi - 1):
                    sh = self.book.sheet_by_index(ref3d.shtxlo)
                    return sh.cell(ref3d.rowxlo, ref3d.colxlo)
        self.dump(logger,header="=== Dump of Name object ===", footer="======= End of dump =======")
        raise XLRDError("Not a constant absolute reference to a single cell")

    def area2d(self, clipped=True):
        """
        This is a convenience method for the use case where the name
        refers to one rectangular area in one worksheet.

        :param clipped:
          If ``True``, the default, the returned rectangle is clipped
          to fit in ``(0, sheet.nrows, 0, sheet.ncols)``.
          it is guaranteed that ``0 <= rowxlo <= rowxhi <= sheet.nrows`` and
          that the number of usable rows in the area (which may be zero) is
          ``rowxhi - rowxlo``; likewise for columns.

        :returns: a tuple ``(sheet_object, rowxlo, rowxhi, colxlo, colxhi)``.

        :raises xlrd.biffh.XLRDError:
           The name is not a constant absolute reference
           to a single area in a single sheet.
        """
        res = self.result
        if res:
            # result should be an instance of the Operand class
            kind = res.kind
            value = res.value
            if kind == oREF and len(value) == 1:  # only 1 reference
                ref3d = value[0]
                if 0 <= ref3d.shtxlo == ref3d.shtxhi - 1:  # only 1 usable sheet
                    sh = self.book.sheet_by_index(ref3d.shtxlo)
                    if not clipped:
                        return sh, ref3d.rowxlo, ref3d.rowxhi, ref3d.colxlo, ref3d.colxhi
                    rowxlo = min(ref3d.rowxlo, sh.nrows)
                    rowxhi = max(rowxlo, min(ref3d.rowxhi, sh.nrows))
                    colxlo = min(ref3d.colxlo, sh.ncols)
                    colxhi = max(colxlo, min(ref3d.colxhi, sh.ncols))
                    assert 0 <= rowxlo <= rowxhi <= sh.nrows
                    assert 0 <= colxlo <= colxhi <= sh.ncols
                    return sh, rowxlo, rowxhi, colxlo, colxhi
        self.dump(logger, header="=== Dump of Name object ===", footer="======= End of dump =======")
        raise XLRDError("Not a constant absolute reference to a single area in a single sheet")


class Book(BaseObject):
    """
    Contents of a "workbook".

    .. warning::

      You should not instantiate this class yourself. You use the :class:`Book`
      object that was returned when you called :func:`~xlrd.open_workbook`.
    """

    #: The number of worksheets present in the workbook file.
    #: This information is available even when no sheets have yet been loaded.
    nsheets = 0

    #: Which date system was in force when this file was last saved.
    #:
    #: 0:
    #:   1900 system (the Excel for Windows default).
    #:
    #: 1:
    #:   1904 system (the Excel for Macintosh default).
    #:
    #: Defaults to 0 in case it's not specified in the file.
    date_mode = 0

    #: Version of BIFF (Binary Interchange File Format) used to create the file.
    #: Latest is 8.0 (represented here as 80), introduced with Excel 97.
    #: Earliest supported by this module: 2.0 (represented as 20).
    biff_version = 0

    #: List containing a :class:`Name` object for each ``NAME`` record in the
    #: workbook.
    #:
    #: .. versionadded:: 0.6.0
    name_obj_list = []

    #: An integer denoting the character set used for strings in this file.
    #: For BIFF 8 and later, this will be 1200, meaning Unicode;
    #: more precisely, UTF_16_LE.
    #: For earlier versions, this is used to derive the appropriate Python
    #: encoding to be used to convert to Unicode.
    #: Examples: ``1252 -> 'cp1252'``, ``10000 -> 'mac_roman'``
    codepage = None

    #: The encoding that was derived from the codepage.
    encoding = None

    #: A tuple containing the telephone country code for:
    #:
    #: ``[0]``:
    #:   the user-interface setting when the file was created.
    #:
    #: ``[1]``:
    #:    the regional settings.
    #:
    #: Example: ``(1, 61)`` meaning ``(USA, Australia)``.
    #:
    #: This information may give a clue to the correct encoding for an
    #: unknown codepage. For a long list of observed values, refer to the
    #: OpenOffice.org documentation for the ``COUNTRY`` record.
    countries = (0, 0)

    #: What (if anything) is recorded as the name of the last user to
    #: save the file.
    user_name = UNICODE_LITERAL('')

    #: A list of :class:`~xlrd.formatting.Font` class instances,
    #: each corresponding to a FONT record.
    #:
    #: .. versionadded:: 0.6.1
    font_list = []

    #: A list of :class:`~xlrd.formatting.XF` class instances,
    #: each corresponding to an ``XF`` record.
    #:
    #: .. versionadded:: 0.6.1
    xf_list = []

    #: A list of :class:`~xlrd.formatting.Format` objects, each corresponding to
    #: a ``FORMAT`` record, in the order that they appear in the input file.
    #: It does *not* contain builtin formats.
    #:
    #: If you are creating an output file using (for example) :mod:`xlwt`,
    #: use this list.
    #:
    #: The collection to be used for all visual rendering purposes is
    #: :attr:`format_map`.
    #:
    #: .. versionadded:: 0.6.1
    format_list = []

    ##
    #: The mapping from :attr:`~xlrd.formatting.XF.format_key` to
    #: :class:`~xlrd.formatting.Format` object.
    #:
    #: .. versionadded:: 0.6.1
    format_map = {}

    #: This provides access via name to the extended format information for
    #: both built-in styles and user-defined styles.
    #:
    #: It maps ``name`` to ``(built_in, xf_index)``, where
    #: ``name`` is either the name of a user-defined style,
    #: or the name of one of the built-in styles. Known built-in names are
    #: Normal, RowLevel_1 to RowLevel_7,
    #: ColLevel_1 to ColLevel_7, Comma, Currency, Percent, "Comma [0]",
    #: "Currency [0]", Hyperlink, and "Followed Hyperlink".
    #:
    #: ``built_in`` has the following meanings
    #:
    #: 1:
    #:     built-in style
    #:
    #: 0:
    #:     user-defined
    #:
    #: ``xf_index`` is an index into :attr:`Book.xf_list`.
    #:
    #: References: OOo docs s6.99 (``STYLE`` record); Excel UI Format/Style
    #:
    #: .. versionadded:: 0.6.1
    #:
    #: Extracted only if ``open_workbook(..., formatting_info=True)``
    #:
    #: .. versionadded:: 0.7.4
    style_name_map = {}

    #: This provides definitions for colour indexes. Please refer to
    #: :ref:`palette` for an explanation
    #: of how colours are represented in Excel.
    #:
    #: Colour indexes into the palette map into ``(red, green, blue)`` tuples.
    #: "Magic" indexes e.g. ``0x7FFF`` map to ``None``.
    #:
    #: :attr:`colour_map` is what you need if you want to render cells on screen
    #: or in a PDF file. If you are writing an output XLS file, use
    #: :attr:`palette_record`.
    #:
    #: .. note:: Extracted only if ``open_workbook(..., formatting_info=True)``
    #:
    #: .. versionadded:: 0.6.1
    colour_map = {}

    #: If the user has changed any of the colours in the standard palette, the
    #: XLS file will contain a ``PALETTE`` record with 56 (16 for Excel 4.0 and
    #: earlier) RGB values in it, and this list will be e.g.
    #: ``[(r0, b0, g0), ..., (r55, b55, g55)]``.
    #: Otherwise this list will be empty. This is what you need if you are
    #: writing an output XLS file. If you want to render cells on screen or in a
    #: PDF file, use :attr:`colour_map`.
    #:
    #: .. note:: Extracted only if ``open_workbook(..., formatting_info=True)``
    #:
    #: .. versionadded:: 0.6.1
    palette_record = []

    #: Time in seconds to extract the XLS image as a contiguous string
    #: (or mmap equivalent).
    load_time_stage_1 = -1.0

    #: Time in seconds to parse the data from the contiguous string
    #: (or mmap equivalent).
    load_time_stage_2 = -1.0

    def sheets(self):
        """
        :returns: A list of all sheets in the book.

        All sheets not already loaded will be loaded.
        """
        for sheetx in xrange(self.nsheets):
            if not self._sheet_list[sheetx]:
                self.get_sheet(sheetx)
        return self._sheet_list[:]

    def sheet_by_index(self, sheetx):
        """
        :param sheetx: Sheet index in ``range(nsheets)``
        :returns: A :class:`~xlrd.sheet.Sheet`.
        """
        return self._sheet_list[sheetx] or self.get_sheet(sheetx)

    def __iter__(self):
        """
        Makes iteration through sheets of a book a little more straightforward.
        Don't free resources after use since it can be called like `list(book)`
        """
        for i in range(self.nsheets):
            yield self.sheet_by_index(i)

    def sheet_by_name(self, sheet_name):
        """
        :param sheet_name: Name of the sheet required.
        :returns: A :class:`~xlrd.sheet.Sheet`.
        """
        try:
            sheetx = self._sheet_names.index(sheet_name)
        except ValueError:
            raise XLRDError('No sheet named <%r>' % sheet_name)
        return self.sheet_by_index(sheetx)

    def __getitem__(self, item):
        """
        Allow indexing with sheet name or index.
        :param item: Name or index of sheet enquired upon
        :return: :class:`~xlrd.sheet.Sheet`.
        """
        if isinstance(item, int):
            return self.sheet_by_index(item)
        else:
            return self.sheet_by_name(item)

    def sheet_names(self):
        """
        :returns:
          A list of the names of all the worksheets in the workbook file.
          This information is available even when no sheets have yet been
          loaded.
        """
        return self._sheet_names[:]

    def sheet_loaded(self, sheet_name_or_index):
        """
        :param sheet_name_or_index: Name or index of sheet enquired upon
        :returns: ``True`` if sheet is loaded, ``False`` otherwise.

        .. versionadded:: 0.7.1
        """
        if isinstance(sheet_name_or_index, int):
            sheetx = sheet_name_or_index
        else:
            try:
                sheetx = self._sheet_names.index(sheet_name_or_index)
            except ValueError:
                raise XLRDError('No sheet named <%r>' % sheet_name_or_index)
        return bool(self._sheet_list[sheetx])

    def unload_sheet(self, sheet_name_or_index):
        """
        :param sheet_name_or_index: Name or index of sheet to be unloaded.

        .. versionadded:: 0.7.1
        """
        if isinstance(sheet_name_or_index, int):
            sheetx = sheet_name_or_index
        else:
            try:
                sheetx = self._sheet_names.index(sheet_name_or_index)
            except ValueError:
                raise XLRDError('No sheet named <%r>' % sheet_name_or_index)
        self._sheet_list[sheetx] = None

    def release_resources(self):
        """
        This method has a dual purpose. You can call it to release
        memory-consuming objects and (possibly) a memory-mapped file
        (:class:`mmap.mmap` object) when you have finished loading sheets in
        ``on_demand`` mode, but still require the :class:`Book` object to
        examine the loaded sheets. It is also called automatically (a) when
        :func:`~xlrd.open_workbook`
        raises an exception and (b) if you are using a ``with`` statement, when
        the ``with`` block is exited. Calling this method multiple times on the
        same object has no ill effect.
        """
        self._resources_released = 1
        if hasattr(self.mem, "close"):
            # must be a mmap.mmap object
            self.mem.close()
        self.mem = None
        if hasattr(self.filestr, "close"):
            self.filestr.close()
        self.filestr = None
        self._sharedstrings = None
        self._rich_text_runlist_map = None

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, exc_tb):
        self.release_resources()
        # return false

    #: A mapping from ``(lower_case_name, scope)`` to a single :class:`Name`
    #:  object.
    #:
    #: .. versionadded:: 0.6.0
    name_and_scope_map = {}

    #: A mapping from `lower_case_name` to a list of :class:`Name` objects.
    #: The list is sorted in scope order. Typically there will be one item
    #: (of global scope) in the list.
    #:
    #: .. versionadded:: 0.6.0
    name_map = {}

    def __init__(self):
        self._sheet_list = []
        self._sheet_names = []
        self._sheet_visibility = []  # from BOUNDSHEET record
        self.nsheets = 0
        self._sh_abs_posn = []  # sheet's absolute position in the stream
        self._sharedstrings = []
        self._rich_text_runlist_map = {}
        self.raw_user_name = False
        self._sheethdr_count = 0  # BIFF 4W only
        self.builtinfmtcount = -1  # unknown as yet. BIFF 3, 4S, 4W
        self.initialise_format_info()
        self._all_sheets_count = 0  # includes macro & VBA sheets
        self._supbook_count = 0
        self._supbook_locals_inx = None
        self._supbook_addins_inx = None
        self._all_sheets_map = []  # maps an all_sheets index to a calc-sheets index (or -1)
        self._externsheet_info = []
        self._externsheet_type_b57 = []
        self._extnsht_name_from_num = {}
        self._sheet_num_from_name = {}
        self._extnsht_count = 0
        self._supbook_types = []
        self._resources_released = 0
        self.addin_func_names = []
        self.name_obj_list = []
        self.colour_map = {}
        self.palette_record = []
        self.xf_list = []
        self.style_name_map = {}
        self.mem = b''
        self.filestr = b''

    def biff2_8_load(self, filename=None, file_contents=None, verbosity=0, use_mmap=USE_MMAP, encoding_override=None,
                     formatting_info=False, on_demand=False, ragged_rows=False, ignore_wb_corruption=False):

        self.verbosity = verbosity
        self.use_mmap = use_mmap
        self.encoding_override = encoding_override
        self.formatting_info = formatting_info
        self.on_demand = on_demand
        self.ragged_rows = ragged_rows

        if not file_contents:
            with open(filename, "rb") as f:
                f.seek(0, 2)  # EOF
                size = f.tell()
                f.seek(0, 0)  # BOF
                if size == 0:
                    raise XLRDError("File size is 0 bytes")
                if self.use_mmap:
                    self.filestr = mmap.mmap(f.fileno(), size, access=mmap.ACCESS_READ)
                    self.stream_len = size
                else:
                    self.filestr = f.read()
                    self.stream_len = len(self.filestr)
        else:
            self.filestr = file_contents
            self.stream_len = len(file_contents)

        self.base = 0
        if self.filestr[:8] != compdoc.SIGNATURE:
            # got this one at the antique store
            self.mem = self.filestr
        else:
            cd = compdoc.CompDoc(filename, self.filestr, ignore_workbook_corruption=ignore_wb_corruption)
            if USE_FANCY_CD:
                for qname in ['Workbook', 'Book']:
                    self.mem, self.base, self.stream_len = cd.locate_named_stream(UNICODE_LITERAL(qname))
                    if self.mem:
                        break
                else:
                    raise XLRDError("Can't find workbook in OLE2 compound document")
            else:
                for qname in ['Workbook', 'Book']:
                    self.mem = cd.get_named_stream(UNICODE_LITERAL(qname))
                    if self.mem:
                        break
                else:
                    raise XLRDError("Can't find workbook in OLE2 compound document")
                self.stream_len = len(self.mem)
            del cd
            if self.mem is not self.filestr:
                if hasattr(self.filestr, "close"):
                    self.filestr.close()
                self.filestr = b''
        self._position = self.base
        logger.debug(f"mem: {type(self.mem)}, base: {self.base}, len: {self.stream_len}")

    def initialise_format_info(self):
        # needs to be done once per sheet for BIFF 4W :-(
        self.format_map = {}
        self.format_list = []
        self.xfcount = 0
        self.actualfmtcount = 0  # number of FORMAT records seen so far
        self._xf_index_to_xl_type_map = {0: XL_CELL_NUMBER}
        self._xf_epilogue_done = 0
        self.xf_list = []
        self.font_list = []

    def get2bytes(self):
        pos = self._position
        buff_two = self.mem[pos:pos + 2]
        len_buff = len(buff_two)
        self._position += len_buff
        if len_buff < 2:
            return MY_EOF
        lo, hi = buff_two
        return (BYTES_ORD(hi) << 8) | BYTES_ORD(lo)

    def get_record_parts(self):
        pos = self._position
        mem = self.mem
        code, length = unpack('<HH', mem[pos:pos + 4])
        pos += 4
        data = mem[pos:pos + length]
        self._position = pos + length
        return code, length, data

    def get_record_parts_conditional(self, reqd_record):
        pos = self._position
        mem = self.mem
        code, length = unpack('<HH', mem[pos:pos + 4])
        if code != reqd_record:
            return None, 0, b''
        pos += 4
        data = mem[pos:pos + length]
        self._position = pos + length
        return code, length, data

    def get_sheet(self, sh_number, update_pos=True):
        if self._resources_released:
            raise XLRDError("Can't load sheets after releasing resources.")
        if update_pos:
            self._position = self._sh_abs_posn[sh_number]
        self.getbof(XL_WORKSHEET)
        # assert biff_version == self.biff_version ### FAILS
        # Have an example where book is v7 but sheet reports v8!!!
        # It appears to work OK if the sheet version is ignored.
        # Confirmed by Daniel Rentz: happens when Excel does "save as"
        # creating an old version file; ignore version details on sheet BOF.
        sh = sheet.Sheet(self, self._position, self._sheet_names[sh_number], sh_number)
        sh.read(self)
        self._sheet_list[sh_number] = sh
        return sh

    def get_sheets(self):
        logger.debug(f"GET_SHEETS: {self._sheet_names} {self._sh_abs_posn}")
        for sheetno in xrange(len(self._sheet_names)):
            logger.debug(f"GET_SHEETS: sheetno={sheetno} {self._sheet_names} {self._sh_abs_posn}")
            self.get_sheet(sheetno)

    def fake_globals_get_sheet(self):  # for BIFF 4.0 and earlier
        formatting.initialise_book(self)
        fake_sheet_name = UNICODE_LITERAL('Sheet 1')
        self._sheet_names = [fake_sheet_name]
        self._sh_abs_posn = [0]
        self._sheet_visibility = [0]  # one sheet, visible
        self._sheet_list.append(None)  # get_sheet updates _sheet_list but needs a None beforehand
        self.get_sheets()

    def handle_boundsheet(self, data):
        bv = self.biff_version
        self.derive_encoding()
        logger.debug(f"BOUNDSHEET: bv={bv} data {data}")
        if bv == 45:  # BIFF4W
            #### Not documented in OOo docs ...
            # In fact, the *only* data is the name of the sheet.
            sheet_name = unpack_string(data, 0, self.encoding, lenlen=1)
            visibility = 0
            sheet_type = XL_BOUNDSHEET_WORKSHEET  # guess, patch later
            if len(self._sh_abs_posn) == 0:
                abs_posn = self._sheetsoffset + self.base
                # Note (a) this won't be used
                # (b) it's the position of the SHEETHDR record
                # (c) add 11 to get to the worksheet BOF record
            else:
                abs_posn = -1  # unknown
        else:
            offset, visibility, sheet_type = unpack('<iBB', data[0:6])
            abs_posn = offset + self.base  # because global BOF is always at posn 0 in the stream
            if bv < BIFF_FIRST_UNICODE:
                sheet_name = unpack_string(data, 6, self.encoding, lenlen=1)
            else:
                sheet_name = unpack_unicode(data, 6, lenlen=1)

        logger.debug(f"BOUNDSHEET: inx={self._all_sheets_count} vis={visibility} "
                     f"sheet_name={sheet_name} abs_posn={abs_posn} sheet_type=0x{sheet_type:02x}")
        self._all_sheets_count += 1

        if sheet_type != XL_BOUNDSHEET_WORKSHEET:
            self._all_sheets_map.append(-1)
            descr = {1: 'Macro sheet', 2: 'Chart', 6: 'Visual Basic module'}.get(sheet_type, 'UNKNOWN')
            logger.debug(f"NOTE * Ignoring non-worksheet data named {sheet_name} (type 0x{sheet_type:02x} = {descr})")
        else:
            snum = len(self._sheet_names)
            self._all_sheets_map.append(snum)
            self._sheet_names.append(sheet_name)
            self._sh_abs_posn.append(abs_posn)
            self._sheet_visibility.append(visibility)
            self._sheet_num_from_name[sheet_name] = snum

    def handle_builtinfmtcount(self, data):
        ### N.B. This count appears to be utterly useless.
        builtinfmtcount = unpack('<H', data[0:2])[0]
        logger.debug(f"BUILTINFMTCOUNT: {builtinfmtcount}")
        self.builtinfmtcount = builtinfmtcount

    def derive_encoding(self):
        if self.encoding_override:
            self.encoding = self.encoding_override
        elif self.codepage is None:
            if self.biff_version < 80:
                logger.warning(f"No CODEPAGE record, no encoding_override: will use 'iso-8859-1'")
                self.encoding = 'iso-8859-1'
            else:
                self.codepage = 1200  # utf16le
                logger.warning(f"No CODEPAGE record; assuming 1200 (utf_16_le)")
        else:
            codepage = self.codepage
            if codepage in encoding_from_codepage:
                encoding = encoding_from_codepage[codepage]
            elif 300 <= codepage <= 1999:
                encoding = 'cp' + str(codepage)
            elif self.biff_version >= 80:
                self.codepage = 1200
                encoding = 'utf_16_le'
            else:
                encoding = 'unknown_codepage_' + str(codepage)
            logger.debug(f"CODEPAGE: codepage {codepage} -> encoding {encoding}")
            self.encoding = encoding
        if self.codepage != 1200:  # utf_16_le
            # If we don't have a codec that can decode ASCII into Unicode,
            # we're well & truly stuffed -- let the punter know ASAP.
            try:
                unicode(b'trial', self.encoding)
            except BaseException as e:
                _class = type(e).__name__.split(".")[-1]
                logger.error(f"codepage {self.codepage} -> encoding {self.encoding} -> {_class}: {e}")
                raise
        if self.raw_user_name:
            strg = unpack_string(self.user_name, 0, self.encoding, lenlen=1)
            strg = strg.rstrip()
            # if DEBUG:
            #     print "CODEPAGE: user name decoded from %r to %r" % (self.user_name, strg)
            self.user_name = strg
            self.raw_user_name = False
        return self.encoding

    def handle_codepage(self, data):
        codepage = unpack('<H', data[0:2])[0]
        self.codepage = codepage
        self.derive_encoding()

    def handle_country(self, data):
        countries = unpack('<HH', data[0:4])
        if self.verbosity:
            logger.debug(f"Countries: {countries}")
        # Note: in BIFF7 and earlier, country record was put (redundantly?) in each worksheet.
        assert self.countries == (0, 0) or self.countries == countries
        self.countries = countries

    def handle_datemode(self, data):
        datemode = unpack('<H', data[0:2])[0]
        logger.debug(f"DATEMODE: datemode {datemode}")
        assert datemode in (0, 1)
        self.date_mode = datemode

    def handle_externname(self, data):
        if self.biff_version >= 80:
            option_flags, other_info = unpack("<HI", data[:6])
            pos = 6
            name, pos = unpack_unicode_update_pos(data, pos, lenlen=1)
            extra = data[pos:]
            if self._supbook_types[-1] == SUPBOOK_ADDIN:
                self.addin_func_names.append(name)
            logger.debug(f"EXTERNNAME: sbktype={self._supbook_types[-1]} oflags=0x{option_flags:04x} "
                         f"oinfo=0x{other_info:08x} name={name} extra={extra}")

    def handle_externsheet(self, data):
        self.derive_encoding()  # in case CODEPAGE record missing/out of order/wrong
        self._extnsht_count += 1  # for use as a 1-based index
        if self.biff_version >= 80:
            num_refs = unpack("<H", data[0:2])[0]
            bytes_reqd = num_refs * 6 + 2
            while len(data) < bytes_reqd:
                logger.debug(f"INFO: EXTERNSHEET needs {bytes_reqd} bytes, have {len(data)}")
                code2, length2, data2 = self.get_record_parts()
                if code2 != XL_CONTINUE:
                    raise XLRDError("Missing CONTINUE after EXTERNSHEET record")
                data += data2
            pos = 2
            for k in xrange(num_refs):
                info = unpack("<HHH", data[pos:pos + 6])
                ref_recordx, ref_first_sheetx, ref_last_sheetx = info
                self._externsheet_info.append(info)
                pos += 6
                logger.debug(f"EXTERNSHEET(b8): k={k:2d}, record={ref_recordx:2d}, "
                             f"first_sheet={ref_first_sheetx:5d}, last sheet={ref_last_sheetx:5d}")
        else:
            nc, ty = unpack("<BB", data[:2])
            hex_char_dump(data, 0, len(data), logger=logger, header="EXTERNSHEET(b7-):")
            msg = {1: "Encoded URL", 2: "Current sheet!!", 3: "Specific sheet in own doc't",
                   4: "Nonspecific sheet in own doc't!!", }.get(ty, "Not encoded")
            logger.debug(f"   {nc:3d} chars, type is {ty} ({msg})")
            if ty == 3:
                sheet_name = unicode(data[2:nc + 2], self.encoding)
                self._extnsht_name_from_num[self._extnsht_count] = sheet_name
                logger.debug(self._extnsht_name_from_num)
            if not (1 <= ty <= 4):
                ty = 0
            self._externsheet_type_b57.append(ty)

    def handle_filepass(self, data):
        if self.verbosity >= 2:
            hex_char_dump(data, 0, len(data), base=0, logger=logger, header="FILEPASS:")
            if self.biff_version >= 80:
                kind1, = unpack('<H', data[:2])
                if kind1 == 0:  # weak XOR encryption
                    key, hash_value = unpack('<HH', data[2:])
                    logger.debug(f'weak XOR: key=0x{key:04x} hash=0x{hash_value:04x}')
                elif kind1 == 1:
                    kind2, = unpack('<H', data[4:6])
                    if kind2 == 1:  # BIFF8 standard encryption
                        caption = "BIFF8 std"
                    elif kind2 == 2:
                        caption = "BIFF8 strong"
                    else:
                        caption = "** UNKNOWN ENCRYPTION METHOD **"
                    logger.debug(caption)
        raise XLRDError("Workbook is encrypted")

    def handle_name(self, data):
        bv = self.biff_version
        if bv < 50:
            return
        self.derive_encoding()
        (
            option_flags, kb_shortcut, name_len, fmla_len, extsht_index, sheet_index,
            menu_text_len, description_text_len, help_topic_text_len, status_bar_text_len,
        ) = unpack("<HBBHHH4B", data[0:14])
        nobj = Name()
        nobj.book = self  ### CIRCULAR ###
        name_index = len(self.name_obj_list)
        nobj.name_index = name_index
        self.name_obj_list.append(nobj)
        nobj.option_flags = option_flags
        attrs = [
            ('hidden', 1, 0),
            ('func', 2, 1),
            ('vbasic', 4, 2),
            ('macro', 8, 3),
            ('complex', 0x10, 4),
            ('builtin', 0x20, 5),
            ('funcgroup', 0xFC0, 6),
            ('binary', 0x1000, 12),
        ]
        for attr, mask, nshift in attrs:
            setattr(nobj, attr, (option_flags & mask) >> nshift)

        macro_flag = " M"[nobj.macro]
        if bv < 80:
            internal_name, pos = unpack_string_update_pos(data, 14, self.encoding, known_len=name_len)
        else:
            internal_name, pos = unpack_unicode_update_pos(data, 14, known_len=name_len)
        nobj.extn_sheet_num = extsht_index
        nobj.excel_sheet_index = sheet_index
        nobj.scope = None  # patched up in the names_epilogue() method
        logger.debug(f"NAME[{name_index}]:{macro_flag} oflags={option_flags}, name_len={name_len}, "
                     f"fmla_len={fmla_len}, extsht_index={extsht_index}, sheet_index={sheet_index}, "
                     f"name={internal_name}")
        name = internal_name
        if nobj.builtin:
            name = builtin_name_from_code.get(name, "??Unknown??")
            logger.debug(f"    builtin: {name}")
        nobj.name = name
        nobj.raw_formula = data[pos:]
        nobj.basic_formula_len = fmla_len
        nobj.evaluated = 0
        nobj.dump(logger,
                  header=f"--- handle_name: name[{name_index}] ---",
                  footer="-------------------")

    def names_epilogue(self):
        blah = self.verbosity >= 2
        if blah:
            logger.debug("+++++ names_epilogue +++++")
            logger.debug("_all_sheets_map " + REPR(self._all_sheets_map))
            logger.debug("_extnsht_name_from_num " + REPR(self._extnsht_name_from_num))
            logger.debug("_sheet_num_from_name " + REPR(self._sheet_num_from_name))
        num_names = len(self.name_obj_list)
        for namex in range(num_names):
            nobj = self.name_obj_list[namex]
            # Convert from excel_sheet_index to scope.
            # This is done here because in BIFF7 and earlier, the
            # BOUNDSHEET records (from which _all_sheets_map is derived)
            # come after the NAME records.
            if self.biff_version >= 80:
                sheet_index = nobj.excel_sheet_index
                if sheet_index == 0:
                    intl_sheet_index = -1  # global
                elif 1 <= sheet_index <= len(self._all_sheets_map):
                    intl_sheet_index = self._all_sheets_map[sheet_index - 1]
                    if intl_sheet_index == -1:  # maps to a macro or VBA sheet
                        intl_sheet_index = -2  # valid sheet reference but not useful
                else:
                    # huh?
                    intl_sheet_index = -3  # invalid
            elif 50 <= self.biff_version <= 70:
                sheet_index = nobj.extn_sheet_num
                if sheet_index == 0:
                    intl_sheet_index = -1  # global
                else:
                    sheet_name = self._extnsht_name_from_num[sheet_index]
                    intl_sheet_index = self._sheet_num_from_name.get(sheet_name, -2)
            nobj.scope = intl_sheet_index

        for namex in range(num_names):
            nobj = self.name_obj_list[namex]
            # Parse the formula ...
            if nobj.macro or nobj.binary:
                continue
            if nobj.evaluated:
                continue
            evaluate_name_formula(self, nobj, namex)

        if self.verbosity >= 2:
            logger.debug(f"---------- name object dump ----------")
            for namex in range(num_names):
                nobj = self.name_obj_list[namex]
                nobj.dump(header="--- name[{namex}] ---")
            logger.debug(f"--------------------------------------")
        #
        # Build some dicts for access to the name objects
        #
        name_and_scope_map = {}  # (name.lower(), scope): Name_object
        name_map = {}  # name.lower() : list of Name_objects (sorted in scope order)
        for namex in range(num_names):
            nobj = self.name_obj_list[namex]
            name_lcase = nobj.name.lower()
            key = (name_lcase, nobj.scope)
            if key in name_and_scope_map and self.verbosity:
                logger.debug(f'Duplicate entry {key} in name_and_scope_map')
            name_and_scope_map[key] = nobj
            sort_data = (nobj.scope, namex, nobj)
            # namex (a temp unique ID) ensures the Name objects will not
            # be compared (fatal in py3)
            if name_lcase in name_map:
                name_map[name_lcase].append(sort_data)
            else:
                name_map[name_lcase] = [sort_data]
        for key in name_map.keys():
            a_list = name_map[key]
            a_list.sort()
            name_map[key] = [x[2] for x in a_list]
        self.name_and_scope_map = name_and_scope_map
        self.name_map = name_map

    def handle_obj(self, data):
        # Not doing much handling at all.
        # Worrying about embedded (BOF ... EOF) substreams is done elsewhere.
        # DEBUG = 1
        obj_type, obj_id = unpack('<HI', data[4:10])
        # if DEBUG: print "---> handle_obj type=%d id=0x%08x" % (obj_type, obj_id)

    def handle_supbook(self, data):
        # aka EXTERNALBOOK in OOo docs
        self._supbook_types.append(None)
        hex_char_dump(data, 0, len(data), logger=logger, header="SUPBOOK:")
        num_sheets = unpack("<H", data[0:2])[0]
        logger.debug(f"num_sheets = {num_sheets}")
        sbn = self._supbook_count
        self._supbook_count += 1
        if data[2:4] == b"\x01\x04":
            self._supbook_types[-1] = SUPBOOK_INTERNAL
            self._supbook_locals_inx = self._supbook_count - 1
            logger.debug(f"SUPBOOK[{sbn}]: internal 3D refs; {num_sheets} sheets")
            logger.debug(f"    _all_sheets_map: {self._all_sheets_map}")
            return
        if data[0:4] == b"\x01\x00\x01\x3A":
            self._supbook_types[-1] = SUPBOOK_ADDIN
            self._supbook_addins_inx = self._supbook_count - 1
            logger.debug(f"SUPBOOK[{sbn}]: add-in functions")
            return
        url, pos = unpack_unicode_update_pos(data, 2, lenlen=2)
        if num_sheets == 0:
            self._supbook_types[-1] = SUPBOOK_DDEOLE
            logger.debug(f"SUPBOOK[{sbn}]: DDE/OLE document = {url}")
            return
        self._supbook_types[-1] = SUPBOOK_EXTERNAL
        logger.debug(f"SUPBOOK[{sbn}]: url = {url}")
        sheet_names = []
        for x in range(num_sheets):
            try:
                shname, pos = unpack_unicode_update_pos(data, pos, lenlen=2)
            except struct.error:
                # #### FIX ME ####
                # Should implement handling of CONTINUE record(s) ...
                if self.verbosity:
                    logger.warning(f"Unpack failure in sheet {x} of {num_sheets} in SUPBOOK record for file {url}")
                break
            sheet_names.append(shname)
            logger.debug(f"  sheetx={x} namelen={len(shname)} name={shname} (next pos={pos})")

    def handle_sheethdr(self, data):
        # This a BIFF 4W special.
        # The SHEETHDR record is followed by a (BOF ... EOF) substream containing
        # a worksheet.
        self.derive_encoding()
        sheet_len = unpack('<i', data[:4])[0]
        sheet_name = unpack_string(data, 4, self.encoding, lenlen=1)
        sheetno = self._sheethdr_count
        assert sheet_name == self._sheet_names[sheetno]
        self._sheethdr_count += 1
        BOF_posn = self._position
        posn = BOF_posn - 4 - len(data)
        logger.debug(f'SHEETHDR {sheetno} at posn {posn}: len={sheet_len} name={sheet_name}')

        self.initialise_format_info()
        logger.debug(f'SHEETHDR: xf epilogue flag is {self._xf_epilogue_done}')

        self._sheet_list.append(None)  # get_sheet updates _sheet_list but needs a None beforehand
        self.get_sheet(sheetno, update_pos=False)
        logger.debug(f'SHEETHDR: posn after get_sheet() ={self._position}')

        self._position = BOF_posn + sheet_len

    def handle_sheetsoffset(self, data):
        pos = unpack('<i', data)[0]
        logger.debug(f'SHEETSOFFSET: {pos}')
        self._sheetsoffset = pos

    def handle_sst(self, data):
        logger.debug("SST Processing")
        t0 = perf_counter()
        nbt = len(data)
        strlist = [data]
        uniquestrings = unpack('<i', data[4:8])[0]
        logger.debug(f"SST: unique strings: {uniquestrings}")

        while True:
            code, nb, data = self.get_record_parts_conditional(XL_CONTINUE)
            if code is None:
                break
            nbt += nb
            logger.debug(f"CONTINUE: adding {nb} bytes to SST -> {nbt}")
            strlist.append(data)
        self._sharedstrings, rt_runlist = unpack_SST_table(strlist, uniquestrings)

        if self.formatting_info:
            self._rich_text_runlist_map = rt_runlist

        t1 = perf_counter()
        logger.debug(f"SST processing took {t1-t0:.2f} seconds")

    def handle_writeaccess(self, data):
        if self.biff_version < 80:
            if not self.encoding:
                self.raw_user_name = True
                self.user_name = data
                return
            txt = unpack_string(data, 0, self.encoding, lenlen=1)
        else:
            try:
                txt = unpack_unicode(data, 0, lenlen=2)
            except UnicodeDecodeError:
                # may have invalid trailing characters
                txt = unpack_unicode(data.strip(), 0, lenlen=2)
        logger.debug(f"WRITEACCESS: {len(data)} bytes; raw={self.raw_user_name} {txt}")
        txt = txt.rstrip()
        self.user_name = txt

    def parse_globals(self):
        # no need to position, just start reading (after the BOF)
        formatting.initialise_book(self)
        while True:
            rc, length, data = self.get_record_parts()
            logger.debug(f"parse_globals: record code is 0x{rc:04x}")
            if rc == XL_SST:
                self.handle_sst(data)
            elif rc == XL_FONT or rc == XL_FONT_B3B4:
                self.handle_font(data)
            elif rc == XL_FORMAT:  # XL_FORMAT2 is BIFF <= 3.0, can't appear in globals
                self.handle_format(data)
            elif rc == XL_XF:
                self.handle_xf(data)
            elif rc == XL_BOUNDSHEET:
                self.handle_boundsheet(data)
            elif rc == XL_DATEMODE:
                self.handle_datemode(data)
            elif rc == XL_CODEPAGE:
                self.handle_codepage(data)
            elif rc == XL_COUNTRY:
                self.handle_country(data)
            elif rc == XL_EXTERNNAME:
                self.handle_externname(data)
            elif rc == XL_EXTERNSHEET:
                self.handle_externsheet(data)
            elif rc == XL_FILEPASS:
                self.handle_filepass(data)
            elif rc == XL_WRITEACCESS:
                self.handle_writeaccess(data)
            elif rc == XL_SHEETSOFFSET:
                self.handle_sheetsoffset(data)
            elif rc == XL_SHEETHDR:
                self.handle_sheethdr(data)
            elif rc == XL_SUPBOOK:
                self.handle_supbook(data)
            elif rc == XL_NAME:
                self.handle_name(data)
            elif rc == XL_PALETTE:
                self.handle_palette(data)
            elif rc == XL_STYLE:
                self.handle_style(data)
            elif rc & 0xff == 9 and self.verbosity:
                pos = self._position-length-4
                logger.warning(f"Unexpected BOF at posn {pos}: 0x{rc:04x} len={length} data={data}")
            elif rc == XL_EOF:
                self.xf_epilogue()
                self.names_epilogue()
                self.palette_epilogue()
                if not self.encoding:
                    self.derive_encoding()
                if self.biff_version == 45:
                    logger.debug(f"global EOF: position {self._position}")
                    # if DEBUG:
                    #     pos = self._position - 4
                    #     print repr(self.mem[pos:pos+40])
                return
            else:
                # if DEBUG:
                #     print >> self.logfile, "parse_globals: ignoring record code 0x%04x" % rc
                pass

    def read(self, pos, length):
        data = self.mem[pos:pos + length]
        self._position = pos + len(data)
        return data

    def getbof(self, rqd_stream):
        logger.debug(f"reqd: 0x{rqd_stream:04x}")

        def bof_error(msg):
            raise XLRDError('Unsupported format, or corrupt file: ' + msg)

        savpos = self._position
        opcode = self.get2bytes()
        if opcode == MY_EOF:
            bof_error('Expected BOF record; met end of file')
        if opcode not in bofcodes:
            bof_error(f'Expected BOF record; found {self.mem[savpos:savpos + 8]}')
        length = self.get2bytes()
        if length == MY_EOF:
            bof_error('Incomplete BOF record[1]; met end of file')
        if not (4 <= length <= 20):
            bof_error(
                'Invalid length (%d) for BOF record type 0x%04x'
                % (length, opcode))
        padding = b'\0' * max(0, boflen[opcode] - length)
        data = self.read(self._position, length)
        logger.debug(f"\ngetbof(): data={data}")
        if len(data) < length:
            bof_error('Incomplete BOF record[2]; met end of file')

        data += padding
        version1 = opcode >> 8
        version2, streamtype = unpack('<HH', data[0:4])
        logger.debug(f"getbof(): op=0x{opcode:04x} version2=0x{version2:04x} streamtype=0x{streamtype:04x}")

        bof_offset = self._position - 4 - length
        logger.debug(f"getbof(): BOF found at offset {bof_offset}; savpos={savpos}")

        version = build = year = 0
        if version1 == 0x08:
            build, year = unpack('<HH', data[4:8])
            if version2 == 0x0600:
                version = 80
            elif version2 == 0x0500:
                if year < 1994 or build in (2412, 3218, 3321):
                    version = 50
                else:
                    version = 70
            else:
                # dodgy one, created by a 3rd-party tool
                version = {
                    0x0000: 21,
                    0x0007: 21,
                    0x0200: 21,
                    0x0300: 30,
                    0x0400: 40,
                }.get(version2, 0)
        elif version1 in (0x04, 0x02, 0x00):
            version = {0x04: 40, 0x02: 30, 0x00: 21}[version1]

        if version == 40 and streamtype == XL_WORKBOOK_GLOBALS_4W:
            version = 45  # i.e. 4W

        logger.debug(f"BOF: op=0x{opcode:04x} vers=0x{version2:04x} stream=0x{streamtype:04x} "
                     f"buildid={build} buildyr={year} -> BIFF{version}")
        got_globals = streamtype == XL_WORKBOOK_GLOBALS or (
                version == 45 and streamtype == XL_WORKBOOK_GLOBALS_4W)
        if (rqd_stream == XL_WORKBOOK_GLOBALS and got_globals) or streamtype == rqd_stream:
            return version
        if version < 50 and streamtype == XL_WORKSHEET:
            return version
        if version >= 50 and streamtype == 0x0100:
            bof_error("Workspace file -- no spreadsheet data")
        bof_error(f'BOF not workbook/worksheet: op=0x{opcode:04x} vers=0x{version2:04x} '
                  f'strm=0x{streamtype:04x} build={build} year={year} -> BIFF{version}')


# === helper functions

def expand_cell_address(in_row, in_col):
    # Ref : OOo docs, "4.3.4 Cell Addresses in BIFF8"
    out_row = in_row
    if in_col & 0x8000:
        if out_row >= 32768:
            out_row -= 65536
        rel_row = 1
    else:
        rel_row = 0
    out_col = in_col & 0xFF
    if in_col & 0x4000:
        if out_col >= 128:
            out_col -= 256
        rel_col = 1
    else:
        rel_col = 0
    return out_row, out_col, rel_row, rel_col


def colname(colx, _A2Z="ABCDEFGHIJKLMNOPQRSTUVWXYZ"):
    assert colx >= 0
    name = UNICODE_LITERAL('')
    while True:
        quot, rem = divmod(colx, 26)
        name = _A2Z[rem] + name
        if not quot:
            return name
        colx = quot - 1


def display_cell_address(rowx, colx, relrow, relcol):
    if relrow:
        row_part = "(*%s%d)" % ("+-"[rowx < 0], abs(rowx))
    else:
        row_part = "$%d" % (rowx + 1,)
    if relcol:
        col_part = "(*%s%d)" % ("+-"[colx < 0], abs(colx))
    else:
        col_part = "$" + colname(colx)
    return col_part + row_part


def unpack_SST_table(data_tab, n_strings):
    """Return list of strings"""
    datainx = 0
    ndatas = len(data_tab)
    data = data_tab[0]
    datalen = len(data)
    pos = 8
    strings = []
    richtext_runs = {}
    latin_1 = "latin_1"
    for _unused_i in xrange(n_strings):
        nchars = unpack('<H', data[pos:pos + 2])[0]
        pos += 2
        options = BYTES_ORD(data[pos])
        pos += 1
        rtcount = 0
        phosz = 0

        if options & 0x08:  # richtext
            rtcount = unpack('<H', data[pos:pos + 2])[0]
            pos += 2
        if options & 0x04:  # phonetic
            phosz = unpack('<i', data[pos:pos + 4])[0]
            pos += 4

        accstrg = UNICODE_LITERAL('')
        chars_got = 0
        while True:
            chars_need = nchars - chars_got
            if options & 0x01:
                # Uncompressed UTF-16
                chars_avail = min((datalen - pos) >> 1, chars_need)
                raw_str = data[pos:pos + 2 * chars_avail]
                logger.debug(f"SST U16: nchars={nchars} pos={pos} raw_str=`{raw_str}`")
                try:
                    accstrg += unicode(raw_str, "utf_16_le")
                except:
                    # print "SST U16: nchars=%d pos=%d rawstrg=%r" % (nchars, pos, rawstrg)
                    # Probable cause: dodgy data e.g. unfinished surrogate pair.
                    # E.g. file unicode2.xls in pyExcelerator's examples has cells containing
                    # unichr(i) for i in range(0x100000)
                    # so this will include 0xD800 etc
                    raise
                pos += 2 * chars_avail
            else:
                # Note: this is COMPRESSED (not ASCII!) encoding!!!
                chars_avail = min(datalen - pos, chars_need)
                raw_str = data[pos:pos + chars_avail]
                logger.debug(f"SST CMPRSD: nchars={nchars} pos={pos} raw_str={raw_str}")
                accstrg += unicode(raw_str, latin_1)
                pos += chars_avail
            chars_got += chars_avail
            if chars_got == nchars:
                break
            datainx += 1
            data = data_tab[datainx]
            datalen = len(data)
            options = BYTES_ORD(data[0])
            pos = 1

        if rtcount:
            runs = []
            for runindex in xrange(rtcount):
                if pos == datalen:
                    pos = 0
                    datainx += 1
                    data = data_tab[datainx]
                    datalen = len(data)
                runs.append(unpack("<HH", data[pos:pos + 4]))
                pos += 4
            richtext_runs[len(strings)] = runs

        pos += phosz  # size of the phonetic stuff to skip
        if pos >= datalen:
            # adjust to correct position in next record
            pos = pos - datalen
            datainx += 1
            if datainx < ndatas:
                data = data_tab[datainx]
                datalen = len(data)
            else:
                assert _unused_i == n_strings - 1
        strings.append(accstrg)
    return strings, richtext_runs
