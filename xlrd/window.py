"""Sparse BIFF8 row-window reader.

This module implements :meth:`xlrd.book.Book.sheet_rows_window`, an entry
point that reads a rectangular window of a worksheet without building the
worksheet's complete cell matrix.

A single sequential pass over the sheet's BIFF records parses only the cells
whose ``(row, column)`` address falls inside the requested half-open ranges.
Every other cell record is decoded far enough to move on to the next record
but does not allocate any cell storage.  No complete :class:`~xlrd.sheet.Sheet`
object is ever created as an intermediate result.

The result is a :class:`RowWindow`: a read-only, sheet-like view with
window-local coordinates.  For every position inside the window its cell
access methods return the same :class:`~xlrd.sheet.Cell` objects and values
that slicing the corresponding ``Book.sheet_by_index(...)`` sheet would
return.
"""

from array import array
from struct import unpack

from .biffh import (
    XL_ARRAY, XL_ARRAY2, XL_BLANK, XL_BOOLERR, XL_COLINFO,
    XL_CONTINUE, XL_DEFAULTROWHEIGHT, XL_DEFCOLWIDTH, XL_DIMENSION,
    XL_DIMENSION2, XL_EOF, XL_FORMULA_OPCODES, XL_GCW, XL_LABEL,
    XL_LABELRANGES, XL_LABELSST, XL_MERGEDCELLS, XL_MULBLANK, XL_MULRK,
    XL_NUMBER, XL_RK, XL_ROW, XL_RSTRING, XL_SHRFMLA, XL_STANDARDWIDTH,
    XL_STRING, XL_STRING_B2, XL_TABLEOP, XL_TABLEOP2, XL_TABLEOP_B2,
    XL_WORKSHEET, XL_CELL_BLANK, XL_CELL_BOOLEAN, XL_CELL_EMPTY,
    XL_CELL_ERROR, XL_CELL_TEXT, XLRDError, bofcodes,
    unpack_cell_range_address_list_update_pos, unpack_unicode,
    unpack_unicode_update_pos,
)
from .sheet import Cell, Colinfo, Rowinfo, unpack_RK
from .timemachine import BYTES_ORD, UNICODE_LITERAL, fprintf, unicode, xrange


class _LocalIndex(object):
    """The in-window cells plus the side data needed for formatting fallbacks."""

    __slots__ = (
        'row_lo', 'row_hi', 'col_lo', 'col_hi',
        'rows', 'cell_count',
        'rowinfo_map', 'colinfo_map', 'merged_cells', 'row_label_ranges',
        'col_label_ranges', 'default_row_height',
        'default_row_height_mismatch', 'default_row_hidden',
        'default_additional_space_above', 'default_additional_space_below',
        'defcolwidth', 'standardwidth', 'gcw',
    )

    def __init__(self, row_lo, row_hi, col_lo, col_hi):
        self.row_lo = row_lo
        self.row_hi = row_hi
        self.col_lo = col_lo
        self.col_hi = col_hi
        self.rows = {}
        self.cell_count = 0
        self.rowinfo_map = {}
        self.colinfo_map = {}
        self.merged_cells = []
        self.row_label_ranges = []
        self.col_label_ranges = []
        self.default_row_height = None
        self.default_row_height_mismatch = 0
        self.default_row_hidden = 0
        self.default_additional_space_above = 0
        self.default_additional_space_below = 0
        self.defcolwidth = None
        self.standardwidth = None
        self.gcw = (0, ) * 256

    def add_cell(self, rowx, colx, ctype, value, xf_index):
        row = self.rows.get(rowx)
        if row is None:
            row = self.rows[rowx] = {}
        row[colx] = (ctype, value, xf_index)
        self.cell_count += 1

    def add_rowinfo(self, rowx, r):
        self.rowinfo_map[rowx] = r

    def add_colinfo(self, colx, c):
        self.colinfo_map[colx] = c


# Record types which, when immediately following a FORMULA whose cached
# result is a special value, may legally appear before the STRING record.
_FORMULA_SUB_RECORD_CODES = frozenset(
    XL_FORMULA_OPCODES + (
        XL_SHRFMLA, XL_ARRAY, XL_ARRAY2, XL_TABLEOP, XL_TABLEOP2,
        XL_TABLEOP_B2,
    )
)


class _WindowScanner(object):
    """One sequential BIFF8 pass that materialises only window cells."""

    def __init__(self, book, sheetx, row_lo, row_hi, col_lo, col_hi,
                 formatting_info):
        self.book = book
        self.sheetx = sheetx
        self.row_lo = row_lo
        self.row_hi = row_hi
        self.col_lo = col_lo
        self.col_hi = col_hi
        # Keep BLANK records whenever the caller wants XF indexes; this
        # mirrors Sheet.read, which only skips them when formatting is off.
        self.keep_blanks = formatting_info
        self.fmt_info = formatting_info
        self.index = _LocalIndex(row_lo, row_hi, col_lo, col_hi)
        self.dimnrows = None
        self.dimncols = None
        self.logfile = book.logfile
        self._rowinfo_sharing_dict = {}

    # -- window membership ------------------------------------------------

    def _in_rows(self, rowx):
        return self.row_lo <= rowx < self.row_hi

    def _in_cols(self, colx):
        return self.col_lo <= colx < self.col_hi

    def _in_window(self, rowx, colx):
        return self._in_rows(rowx) and self._in_cols(colx)

    def _put(self, rowx, colx, ctype, value, xf_index):
        if not self._in_window(rowx, colx):
            return
        if ctype is None:
            # Number/RK records leave the type to the XF's format type.
            ctype = self.book._xf_index_to_xl_type_map[xf_index]
        self.index.add_cell(rowx, colx, ctype, value, xf_index)

    # -- the single BIFF pass ---------------------------------------------

    def scan(self, start_position):
        book = self.book
        old_position = book._position
        book._position = start_position
        try:
            while 1:
                rc, data_len, data = book.get_record_parts()
                if rc == XL_EOF:
                    break
                self._dispatch(rc, data_len, data)
        finally:
            book._position = old_position
        return self.index

    def _dispatch(self, rc, data_len, data):
        if rc == XL_NUMBER:
            rowx, colx, xf_index, d = unpack('<HHHd', data[:14])
            self._put(rowx, colx, None, d, xf_index)
        elif rc == XL_LABELSST:
            rowx, colx, xf_index, sstindex = unpack('<HHHi', data)
            # SST lookup is skipped for out-of-window cells, so shared
            # strings outside the window cost no work beyond this record.
            if self._in_window(rowx, colx):
                self._put(
                    rowx, colx, XL_CELL_TEXT,
                    self.book._sharedstrings[sstindex], xf_index,
                )
        elif rc == XL_LABEL:
            rowx, colx, xf_index = unpack('<HHH', data[0:6])
            if self._in_window(rowx, colx):
                strg = unpack_unicode(data, 6, lenlen=2)
                self._put(rowx, colx, XL_CELL_TEXT, strg, xf_index)
        elif rc == XL_RSTRING:
            rowx, colx, xf_index = unpack('<HHH', data[0:6])
            if self._in_window(rowx, colx):
                strg, _pos = unpack_unicode_update_pos(data, 6, lenlen=2)
                self._put(rowx, colx, XL_CELL_TEXT, strg, xf_index)
        elif rc == XL_RK:
            rowx, colx, xf_index = unpack('<HHH', data[:6])
            if self._in_window(rowx, colx):
                d = unpack_RK(data[6:10])
                self._put(rowx, colx, None, d, xf_index)
        elif rc == XL_MULRK:
            mulrk_row, mulrk_first = unpack('<HH', data[0:4])
            mulrk_last, = unpack('<H', data[-2:])
            if self._in_rows(mulrk_row):
                pos = 4
                for colx in xrange(mulrk_first, mulrk_last + 1):
                    xf_index, = unpack('<H', data[pos:pos + 2])
                    if self._in_cols(colx):
                        d = unpack_RK(data[pos + 2:pos + 6])
                        self._put(mulrk_row, colx, None, d, xf_index)
                    pos += 6
        elif rc in XL_FORMULA_OPCODES:
            self._handle_formula(data)
        elif rc == XL_BOOLERR:
            rowx, colx, xf_index, value, is_err = unpack('<HHHBB', data[:8])
            if self._in_window(rowx, colx):
                cellty = (XL_CELL_BOOLEAN, XL_CELL_ERROR)[is_err]
                self._put(rowx, colx, cellty, value, xf_index)
        elif rc == XL_BLANK:
            if not self.keep_blanks:
                return
            rowx, colx, xf_index = unpack('<HHH', data[:6])
            if self._in_window(rowx, colx):
                self._put(rowx, colx, XL_CELL_BLANK, '', xf_index)
        elif rc == XL_MULBLANK:
            if not self.keep_blanks:
                return
            nitems = data_len >> 1
            result = unpack("<%dH" % nitems, data)
            rowx, mul_first = result[:2]
            mul_last = result[-1]
            if self._in_rows(rowx):
                pos = 2
                for colx in xrange(mul_first, mul_last + 1):
                    if self._in_cols(colx):
                        self._put(rowx, colx, XL_CELL_BLANK, '', result[pos])
                    pos += 1
        elif rc == XL_DIMENSION or rc == XL_DIMENSION2:
            if data_len < 12:
                # Empty DIMENSIONS records occur in the wild; see issue 64.
                return
            # BIFF8 layout: <rowlo i><rowhi i><collo H><colhi H><unused H>
            row_hi, col_hi = unpack('<4xixxH', data[:12])
            self.dimnrows = row_hi
            self.dimncols = col_hi
        elif rc == XL_ROW:
            rowx, bits1, bits2 = unpack('<H4xH4xi', data[0:16])
            if not self._in_rows(rowx):
                return
            key = (bits1, bits2)
            r = self._rowinfo_sharing_dict.get(key)
            if r is None:
                self._rowinfo_sharing_dict[key] = r = Rowinfo()
                r.height = bits1 & 0x7fff
                r.has_default_height = (bits1 >> 15) & 1
                r.outline_level = bits2 & 7
                r.outline_group_starts_ends = (bits2 >> 4) & 1
                r.hidden = (bits2 >> 5) & 1
                r.height_mismatch = (bits2 >> 6) & 1
                r.has_default_xf_index = (bits2 >> 7) & 1
                r.xf_index = (bits2 >> 16) & 0xfff
                r.additional_space_above = (bits2 >> 28) & 1
                r.additional_space_below = (bits2 >> 29) & 1
                if not r.has_default_xf_index:
                    r.xf_index = -1
            self.index.add_rowinfo(rowx, r)
        elif rc == XL_COLINFO:
            c = Colinfo()
            first_colx, last_colx, c.width, c.xf_index, flags = \
                unpack("<HHHHH", data[:10])
            if not (0 <= first_colx <= last_colx <= 256):
                return
            c.hidden = flags & 0x0001
            c.bit1_flag = (flags & 0x0002) >> 1
            c.outline_level = (flags & 0x0700) >> 8
            c.collapsed = (flags & 0x1000) >> 12
            for colx in xrange(first_colx, last_colx + 1):
                if colx > 255:
                    break
                if self._in_cols(colx):
                    self.index.add_colinfo(colx, c)
        elif rc in (XL_MERGEDCELLS, XL_LABELRANGES, XL_DEFAULTROWHEIGHT,
                    XL_DEFCOLWIDTH, XL_STANDARDWIDTH, XL_GCW):
            self._handle_saved_metadata(rc, data_len, data)
        elif rc in bofcodes:
            # Embedded chart BOF: skip records until its matching EOF.
            self._skip_embedded_bof()
        # All other record types (drawing objects, notes, hyperlinks,
        # conditional formats, ...) are irrelevant to a cell window and
        # are skipped without inspection.

    # -- record helpers -----------------------------------------------------

    def _handle_formula(self, data):
        # BIFF8 FORMULA: <rowx H><colx H><xf H><result 8s><flags H>...
        rowx, colx, xf_index, result_str, _flags = \
            unpack('<HHH8sH', data[0:16])
        if result_str[6:8] == b"\xFF\xFF":
            first_byte = BYTES_ORD(result_str[0])
            if first_byte == 0:
                # String result: the value lives in a following STRING
                # record, possibly after SHRFMLA/ARRAY/TABLEOP records.
                self._handle_string_formula(rowx, colx, xf_index)
            elif first_byte == 1:
                value = BYTES_ORD(result_str[2])
                self._put(rowx, colx, XL_CELL_BOOLEAN, value, xf_index)
            elif first_byte == 2:
                value = BYTES_ORD(result_str[2])
                self._put(rowx, colx, XL_CELL_ERROR, value, xf_index)
            elif first_byte == 3:
                # Empty (zero-length) string, NOT an empty cell.
                self._put(rowx, colx, XL_CELL_TEXT, "", xf_index)
            else:
                raise XLRDError(
                    "unexpected special case (0x%02x) in FORMULA" % first_byte)
        else:
            d = unpack('<d', result_str)[0]
            self._put(rowx, colx, None, d, xf_index)

    def _handle_string_formula(self, rowx, colx, xf_index):
        book = self.book
        gotstring = False
        # An optional SHRFMLA or ARRAY etc record may precede STRING.
        rc2, _len2, data2 = book.get_record_parts()
        if rc2 in (XL_STRING, XL_STRING_B2):
            gotstring = True
        elif rc2 not in _FORMULA_SUB_RECORD_CODES:
            raise XLRDError(
                "Expected SHRFMLA, ARRAY, TABLEOP* or STRING record; "
                "found 0x%04x" % rc2)
        if not gotstring:
            rc2, _len2, data2 = book.get_record_parts()
            if rc2 not in (XL_STRING, XL_STRING_B2):
                raise XLRDError(
                    "Expected STRING record; found 0x%04x" % rc2)
        if self._in_window(rowx, colx):
            strg = self._string_record_contents(data2)
            self._put(rowx, colx, XL_CELL_TEXT, strg, xf_index)

    def _string_record_contents(self, data):
        # Same decoding as Sheet.string_record_contents, BIFF8 only.
        nchars_expected = unpack("<H", data[:2])[0]
        offset = 2
        nchars_found = 0
        result = UNICODE_LITERAL("")
        while 1:
            flag = BYTES_ORD(data[offset]) & 1
            enc = ("latin_1", "utf_16_le")[flag]
            offset += 1
            chunk = unicode(data[offset:], enc)
            result += chunk
            nchars_found += len(chunk)
            if nchars_found == nchars_expected:
                return result
            if nchars_found > nchars_expected:
                raise XLRDError(
                    "STRING/CONTINUE: expected %d chars, found %d"
                    % (nchars_expected, nchars_found))
            rc, _unused_len, data = self.book.get_record_parts()
            if rc != XL_CONTINUE:
                raise XLRDError(
                    "Expected CONTINUE record; found record-type 0x%04X" % rc)
            offset = 0

    def _skip_embedded_bof(self):
        # Embedded chart sub-stream: consume records up to its EOF.
        while 1:
            rc, _len, _data = self.book.get_record_parts()
            if rc == XL_EOF:
                break

    def _handle_saved_metadata(self, rc, data_len, data):
        index = self.index
        if rc == XL_MERGEDCELLS:
            pos = unpack_cell_range_address_list_update_pos(
                index.merged_cells, data, 0, 80, addr_size=8)
            assert pos == data_len, \
                "MERGEDCELLS: pos=%d data_len=%d" % (pos, data_len)
        elif rc == XL_LABELRANGES:
            pos = unpack_cell_range_address_list_update_pos(
                index.row_label_ranges, data, 0, 80, addr_size=8)
            pos = unpack_cell_range_address_list_update_pos(
                index.col_label_ranges, data, pos, 80, addr_size=8)
            assert pos == data_len
        elif rc == XL_DEFAULTROWHEIGHT:
            if data_len == 4:
                bits, index.default_row_height = unpack("<HH", data[:4])
            elif data_len == 2:
                index.default_row_height, = unpack("<H", data)
                bits = 0
                fprintf(self.logfile,
                    "*** WARNING: DEFAULTROWHEIGHT record len is 2, "
                    "should be 4; assuming BIFF2 format\n")
            else:
                bits = 0
                fprintf(self.logfile,
                    "*** WARNING: DEFAULTROWHEIGHT record len is %d, "
                    "should be 4; ignoring this record\n", data_len)
            index.default_row_height_mismatch = bits & 1
            index.default_row_hidden = (bits >> 1) & 1
            index.default_additional_space_above = (bits >> 2) & 1
            index.default_additional_space_below = (bits >> 3) & 1
        elif rc == XL_DEFCOLWIDTH:
            index.defcolwidth, = unpack("<H", data[:2])
        elif rc == XL_STANDARDWIDTH:
            index.standardwidth, = unpack("<H", data[:2])
        elif rc == XL_GCW:
            if data_len != 34:
                return
            iguff = unpack("<8i", data[2:34])
            gcw = []
            for bits in iguff:
                for _unused in xrange(32):
                    gcw.append(bits & 1)
                    bits >>= 1
            index.gcw = tuple(gcw)


class RowWindow(object):
    """A sparse, read-only rectangular view over one worksheet.

    Constructed by :meth:`xlrd.book.Book.sheet_rows_window`; instances are
    not meant to be created directly.

    Coordinates used by the access methods are relative to the requested
    window: row 0 corresponds to the workbook's ``row_lo`` row, column 0 to
    ``col_lo``.  Formatting-related attributes (``rowinfo_map``,
    ``colinfo_map``, :meth:`cell_xf_index`, ...) keep absolute workbook
    coordinates so the values remain valid indexes into
    :attr:`xlrd.book.Book.xf_list`.

    The cell access methods return the same :class:`~xlrd.sheet.Cell`
    objects, types and values as
    ``book.sheet_by_index(sheetx).row_slice(row_lo + r, col_lo, col_hi)``.
    Positions inside the window that have no cell record are represented as
    :data:`~xlrd.biffh.XL_CELL_EMPTY` cells, exactly as in a fully loaded,
    non-ragged sheet.
    """

    def __init__(self, book, sheetx, name, index, dimnrows, dimncols,
                 formatting_info):
        self.book = book
        self.number = sheetx
        self.name = name
        self.formatting_info = formatting_info
        self.ragged_rows = book.ragged_rows
        self.biff_version = book.biff_version
        self.row_lo = index.row_lo
        self.row_hi = index.row_hi
        self.col_lo = index.col_lo
        self.col_hi = index.col_hi
        self.nrows = index.row_hi - index.row_lo
        self.ncols = index.col_hi - index.col_lo
        self._dimnrows = dimnrows
        self._dimncols = dimncols
        self.cell_count = index.cell_count
        self.rowinfo_map = index.rowinfo_map
        self.colinfo_map = index.colinfo_map
        self.merged_cells = index.merged_cells
        self.row_label_ranges = index.row_label_ranges
        self.col_label_ranges = index.col_label_ranges
        self.default_row_height = index.default_row_height
        self.default_row_height_mismatch = index.default_row_height_mismatch
        self.default_row_hidden = index.default_row_hidden
        self.default_additional_space_above = \
            index.default_additional_space_above
        self.default_additional_space_below = \
            index.default_additional_space_below
        self.defcolwidth = index.defcolwidth
        self.standardwidth = index.standardwidth
        self.gcw = index.gcw
        self.visibility = book._sheet_visibility[sheetx]
        self._xf_index_stats = [0, 0, 0, 0]
        self._build_matrices(index)

    def _build_matrices(self, index):
        nrows = self.nrows
        ncols = self.ncols
        empty_row_values = [UNICODE_LITERAL('')] * ncols
        self._cell_values = [list(empty_row_values) for _ in xrange(nrows)]
        self._cell_types = [
            array('B', [XL_CELL_EMPTY]) * ncols for _ in xrange(nrows)
        ]
        if self.formatting_info:
            self._cell_xf_indexes = [
                array('h', [-1]) * ncols for _ in xrange(nrows)
            ]
        else:
            self._cell_xf_indexes = None
        for growx, row in index.rows.items():
            lrowx = growx - self.row_lo
            types_row = self._cell_types[lrowx]
            values_row = self._cell_values[lrowx]
            if self.formatting_info:
                xf_row = self._cell_xf_indexes[lrowx]
            for gcolx, (ctype, value, xf_index) in row.items():
                lcolx = gcolx - self.col_lo
                types_row[lcolx] = ctype
                values_row[lcolx] = value
                if self.formatting_info:
                    xf_row[lcolx] = xf_index

    # -- cell access (Sheet-compatible API, window-local coordinates) -------

    def cell(self, rowx, colx):
        """
        :class:`~xlrd.sheet.Cell` at window position ``(rowx, colx)``.
        """
        if self.formatting_info:
            xfx = self.cell_xf_index(rowx, colx)
        else:
            xfx = None
        return Cell(
            self._cell_types[rowx][colx],
            self._cell_values[rowx][colx],
            xfx,
        )

    def cell_value(self, rowx, colx):
        "Value of the cell at window position ``(rowx, colx)``."
        return self._cell_values[rowx][colx]

    def cell_type(self, rowx, colx):
        "Type of the cell at window position ``(rowx, colx)``."
        return self._cell_types[rowx][colx]

    def cell_xf_index(self, rowx, colx):
        """
        XF index of the cell at window position ``(rowx, colx)``.

        The fallback rules (cell XF -> row XF -> column XF -> global
        default 15) mirror :meth:`xlrd.sheet.Sheet.cell_xf_index`.

        .. versionadded:: 2.1.0
        """
        if not self.formatting_info:
            raise XLRDError(
                "Feature requires open_workbook(..., formatting_info=True)")
        xfx = self._cell_xf_indexes[rowx][colx]
        if xfx > -1:
            self._xf_index_stats[0] += 1
            return xfx
        growx = rowx + self.row_lo
        r = self.rowinfo_map.get(growx)
        if r is not None:
            if r.xf_index > -1:
                self._xf_index_stats[1] += 1
                return r.xf_index
        gcolx = colx + self.col_lo
        c = self.colinfo_map.get(gcolx)
        if c is not None:
            self._xf_index_stats[2] += 1
            if c.xf_index == -1:
                return 15
            return c.xf_index
        self._xf_index_stats[3] += 1
        return 15

    def row_len(self, rowx):
        """Always returns :attr:`ncols`: window rows are not ragged."""
        return self.ncols

    def row(self, rowx):
        "Sequence of :class:`~xlrd.sheet.Cell` objects for one window row."
        return [self.cell(rowx, colx) for colx in xrange(self.ncols)]

    def __getitem__(self, item):
        try:
            rowix, colix = item
        except TypeError:
            return self.row(item)
        return self.cell(rowix, colix)

    def get_rows(self):
        "Generator yielding each window row as a list of cells."
        return (self.row(index) for index in range(self.nrows))

    __iter__ = get_rows

    def row_types(self, rowx, start_colx=0, end_colx=None):
        "Types slice for the given window row."
        if end_colx is None:
            return self._cell_types[rowx][start_colx:]
        return self._cell_types[rowx][start_colx:end_colx]

    def row_values(self, rowx, start_colx=0, end_colx=None):
        "Values slice for the given window row."
        if end_colx is None:
            return self._cell_values[rowx][start_colx:]
        return self._cell_values[rowx][start_colx:end_colx]

    def row_slice(self, rowx, start_colx=0, end_colx=None):
        ":class:`~xlrd.sheet.Cell` slice for the given window row."
        nc = self.ncols
        if start_colx < 0:
            start_colx += nc
            if start_colx < 0:
                start_colx = 0
        if end_colx is None or end_colx > nc:
            end_colx = nc
        elif end_colx < 0:
            end_colx += nc
        return [
            self.cell(rowx, colx)
            for colx in xrange(start_colx, end_colx)
        ]

    def col_slice(self, colx, start_rowx=0, end_rowx=None):
        ":class:`~xlrd.sheet.Cell` slice for the given window column."
        nr = self.nrows
        if start_rowx < 0:
            start_rowx += nr
            if start_rowx < 0:
                start_rowx = 0
        if end_rowx is None or end_rowx > nr:
            end_rowx = nr
        elif end_rowx < 0:
            end_rowx += nr
        return [
            self.cell(rowx, colx)
            for rowx in xrange(start_rowx, end_rowx)
        ]

    col = col_slice

    def col_values(self, colx, start_rowx=0, end_rowx=None):
        "Values slice for the given window column."
        nr = self.nrows
        if start_rowx < 0:
            start_rowx += nr
            if start_rowx < 0:
                start_rowx = 0
        if end_rowx is None or end_rowx > nr:
            end_rowx = nr
        elif end_rowx < 0:
            end_rowx += nr
        return [
            self._cell_values[rowx][colx]
            for rowx in xrange(start_rowx, end_rowx)
        ]

    def col_types(self, colx, start_rowx=0, end_rowx=None):
        "Types slice for the given window column."
        nr = self.nrows
        if start_rowx < 0:
            start_rowx += nr
            if start_rowx < 0:
                start_rowx = 0
        if end_rowx is None or end_rowx > nr:
            end_rowx = nr
        elif end_rowx < 0:
            end_rowx += nr
        return [
            self._cell_types[rowx][colx]
            for rowx in xrange(start_rowx, end_rowx)
        ]

    def computed_column_width(self, colx):
        """
        Display width for window column ``colx``, using the same
        BIFF8 fallback rules as :meth:`xlrd.sheet.Sheet.computed_column_width`.
        """
        if not self.formatting_info:
            raise XLRDError(
                "Feature requires open_workbook(..., formatting_info=True)")
        c = self.colinfo_map.get(colx + self.col_lo)
        if c is not None:
            return c.width
        if self.standardwidth is not None:
            return self.standardwidth
        if self.defcolwidth is not None:
            return self.defcolwidth * 256
        return 8 * 256

    def __repr__(self):
        return "RowWindow(sheetx=%d, name=%r, rows=[%d:%d), cols=[%d:%d), cells=%d)" % (
            self.number, self.name, self.row_lo, self.row_hi,
            self.col_lo, self.col_hi, self.cell_count,
        )


def read_row_window(book, sheetx, row_lo, row_hi, col_lo, col_hi,
                    formatting_info=None):
    """Read one worksheet window without materialising its full cell matrix.

    The function performs a single sequential pass over the worksheet's
    BIFF8 records.  Cell records outside the window are skipped in place;
    only cells whose address is in ``row_lo <= rowx < row_hi`` and
    ``col_lo <= colx < col_hi`` are stored.  Workbook-level shared metadata
    (SST, XF/FORMAT/FONT tables, ...) loaded by :func:`xlrd.open_workbook`
    is reused as-is.

    :param book: the :class:`~xlrd.book.Book` returned by
        :func:`xlrd.open_workbook`, opened with ``on_demand=True`` so that
        the workbook stream is still available.
    :param sheetx: sheet index in ``range(book.nsheets)``.
    :param row_lo: first workbook row of the window (inclusive).
    :param row_hi: one past the last workbook row of the window.
    :param col_lo: first workbook column of the window (inclusive).
    :param col_hi: one past the last workbook column of the window.
    :param formatting_info: when true (defaulting to the book-level
        ``formatting_info`` flag), ``BLANK``/``MULBLANK`` cells are kept,
        per-cell XF indexes are retained and :class:`RowWindow` exposes the
        row/column formatting maps.  When false, blank records are skipped,
        matching :func:`xlrd.open_workbook` with ``formatting_info=False``.
    :returns: a :class:`RowWindow`.
    """
    if not (0 <= sheetx < book.nsheets):
        raise XLRDError(
            "Sheet index %d out of range (0..%d)"
            % (sheetx, book.nsheets - 1))
    if book.biff_version < 80:
        raise XLRDError(
            "Row-window reads are only supported for BIFF8 workbooks; "
            "this workbook reports BIFF version %d. "
            "Use sheet_by_index() for older workbooks."
            % book.biff_version)
    if not (0 <= row_lo <= row_hi <= 65536):
        raise ValueError(
            "row window [%d, %d) must satisfy 0 <= row_lo <= row_hi <= 65536"
            % (row_lo, row_hi))
    if not (0 <= col_lo <= col_hi <= 256):
        raise ValueError(
            "column window [%d, %d) must satisfy 0 <= col_lo <= col_hi <= 256"
            % (col_lo, col_hi))
    if book._resources_released:
        raise XLRDError(
            "Can't read a sheet window after releasing resources; "
            "open the workbook with on_demand=True.")
    if formatting_info is None:
        formatting_info = book.formatting_info

    # Consume the sheet BOF exactly like Book.get_sheet does, but never
    # register a full Sheet in book._sheet_list.
    saved_position = book._position
    book._position = book._sh_abs_posn[sheetx]
    try:
        book.getbof(XL_WORKSHEET)
        start_position = book._position
        scanner = _WindowScanner(
            book, sheetx, row_lo, row_hi, col_lo, col_hi, formatting_info)
        index = scanner.scan(start_position)
    finally:
        book._position = saved_position
    return RowWindow(
        book, sheetx, book._sheet_names[sheetx], index,
        scanner.dimnrows, scanner.dimncols, formatting_info,
    )
