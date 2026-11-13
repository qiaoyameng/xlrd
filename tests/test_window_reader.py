"""Tests for the BIFF8 sparse row-window reader."""

from struct import pack
from unittest import TestCase

import xlrd
from xlrd.biffh import (
    XL_CELL_BLANK, XL_CELL_BOOLEAN, XL_CELL_EMPTY, XL_CELL_ERROR,
    XL_CELL_NUMBER, XL_CELL_TEXT,
)
from xlrd.window import RowWindow


def rec(code, data=b''):
    return pack('<HH', code, len(data)) + data


def bof(stream_type):
    return rec(0x0809, pack('<HHHHI', 0x0600, stream_type, 0x0DBB,
                            0x07CC, 0x00000041))


def sst_record(strings):
    body = pack('<ii', len(strings), len(strings))
    for s in strings:
        raw = s.encode('latin_1')
        body += pack('<HB', len(raw), 0x00) + raw
    return rec(0x00FC, body)


def boundsheet(offset, name):
    nm = name.encode('latin_1')
    return rec(0x0085, pack('<iBB', offset, 0, 0) +
               pack('<B', len(nm)) + b'\x00' + nm)


def xf_record():
    # <font H><format H><type H><align B><rot B><align2 B><used B><b1 I><b2 i><b3 H>
    return rec(0x00E0, b'\x00' * 20)


def number_record(row, col, xf, value):
    return rec(0x0203, pack('<HHHd', row, col, xf, value))


def labelsst_record(row, col, xf, sst_idx):
    return rec(0x00FD, pack('<HHHi', row, col, xf, sst_idx))


def rk_record(row, col, xf, rk):
    return rec(0x027E, pack('<HHHI', row, col, xf, rk))


def mulrk_record(row, first_col, items):
    body = pack('<HH', row, first_col)
    for xf, rk in items:
        body += pack('<HI', xf, rk)
    body += pack('<H', first_col + len(items) - 1)
    return rec(0x00BD, body)


def blank_record(row, col, xf):
    return rec(0x0201, pack('<HHH', row, col, xf))


def mulblank_record(row, first_col, xfs):
    body = pack('<HH', row, first_col)
    for xf in xfs:
        body += pack('<H', xf)
    body += pack('<H', first_col + len(xfs) - 1)
    return rec(0x00BE, body)


def boolerr_record(row, col, xf, value, is_err):
    return rec(0x0205, pack('<HHHBB', row, col, xf, value, is_err))


def formula_record(row, col, xf, result_8, flags=0):
    return rec(0x0006, pack('<HHH8sH', row, col, xf, result_8, flags))


def string_record(text):
    raw = text.encode('latin_1')
    return rec(0x0207, pack('<HB', len(raw), 0) + raw)


def row_record(row, xf_index=0):
    # BIFF8 ROW record is 16 bytes. xlrd reads it as <row H>4x<height H>4x<bits i>.
    # bit 7 of bits = has_default_xf_index; bits 16..27 = the XF index.
    bits = 0x0080 | ((xf_index & 0x0FFF) << 16)
    return rec(0x0208, pack('<HHHH4xi', row, 0, 0, 0x00FF, bits))


def dimension_record(nrows, ncols):
    return rec(0x0200, pack('<iiHHH', 0, nrows, 0, ncols, 0))


def colinfo_record(first_col, last_col, width, xf, flags=0):
    return rec(0x007D, pack('<HHHHH', first_col, last_col, width, xf, flags))


def build_workbook(cell_records, nrows=100, ncols=8, sst_strings=()):
    globals_stream = b''.join((
        bof(0x0005),
        sst_record(list(sst_strings)),
        xf_record(),
        xf_record(),
    ))
    boundsheet_len = len(boundsheet(0, 'WindowSheet')) + 4  # + globals EOF
    sheet_offset = len(globals_stream) + boundsheet_len
    globals_stream += boundsheet(sheet_offset, 'WindowSheet')
    globals_stream += rec(0x000A)  # globals EOF

    sheet_stream = b''.join((
        bof(0x0010),
        dimension_record(nrows, ncols),
        cell_records,
        rec(0x000A),  # sheet EOF
    ))
    return globals_stream + sheet_stream

# ---------------------------------------------------------------------------
# Test helpers: minimal synthetic BIFF8 workbook construction
# ---------------------------------------------------------------------------

NROWS = 200
NCOLS = 8


def _ordered_number_grid(skip=()):
    parts = []
    for r in range(1, NROWS):
        for c in range(NCOLS):
            if (r, c) in skip:
                continue
            parts.append(number_record(r, c, 0, float(r * NCOLS + c)))
    return b''.join(parts)


def _window_rows(book, **kw):
    return book.sheet_rows_window(0, 2, 7, 1, 5, **kw)


class UnorderedRecordsTest(TestCase):
    """A late record for an early row must still land in the window."""

    def setUp(self):
        records = _ordered_number_grid(skip={(2, 1)})
        # Deliberately emit rows out of order: this cell for row 2 is
        # written AFTER every record of rows 3..NROWS-1.
        records += number_record(2, 1, 0, 999.0)
        records += blank_record(0, 1, 1)      # late, outside the window
        records += number_record(0, 0, 0, 4242.0)  # late, outside window
        self.stream = build_workbook(
            records, nrows=NROWS, ncols=NCOLS)

    def test_late_in_window_cell_is_materialised(self):
        bk = xlrd.open_workbook(file_contents=self.stream,
                                formatting_info=True, on_demand=True)
        win = _window_rows(bk)
        # Window row 0, col 0 corresponds to workbook cell (2, 1).
        self.assertEqual(win.cell_value(0, 0), 999.0)
        self.assertEqual(win.cell_type(0, 0), XL_CELL_NUMBER)
        bk.release_resources()

    def test_matches_sheet_by_index_slice(self):
        full = xlrd.open_workbook(
            file_contents=self.stream, formatting_info=True,
            on_demand=True).sheet_by_index(0)
        win = xlrd.open_workbook(
            file_contents=self.stream, formatting_info=True,
            on_demand=True).sheet_rows_window(0, 2, 7, 1, 5)
        self.assertEqual(win.nrows, 5)
        self.assertEqual(win.ncols, 4)
        for r in range(win.nrows):
            self.assertEqual(
                win.row_values(r), full.row_values(2 + r, 1, 5))
            self.assertEqual(
                list(win.row_types(r)),
                list(full.row_types(2 + r, 1, 5)))
            self.assertEqual(
                [c.xf_index for c in win.row_slice(r)],
                [c.xf_index for c in full.row_slice(2 + r, 1, 5)])

    def test_only_window_cells_are_instantiated(self):
        full = xlrd.open_workbook(
            file_contents=self.stream, formatting_info=True,
            on_demand=True).sheet_by_index(0)
        full_entries = sum(len(row) for row in full._cell_types)
        self.assertEqual(full_entries, NROWS * NCOLS)

        bk = xlrd.open_workbook(
            file_contents=self.stream, formatting_info=True, on_demand=True)
        win = _window_rows(bk)
        # 5 rows * 4 cols = 20 cells, even though the sheet holds 1600.
        self.assertEqual(win.cell_count, 20)
        self.assertEqual(
            sum(len(row) for row in win._cell_types), win.cell_count)
        self.assertLess(win.cell_count, full_entries)
        bk.release_resources()

    def test_out_of_window_late_cells_are_not_built(self):
        bk = xlrd.open_workbook(
            file_contents=self.stream, formatting_info=True, on_demand=True)
        win = _window_rows(bk)
        # The two late out-of-window records must not have been stored:
        # 20 in-window cells, not 22.
        self.assertEqual(win.cell_count, 20)
        # The dense matrix only spans the window, not the whole sheet.
        self.assertEqual(len(win._cell_values), 5)
        self.assertTrue(all(len(row) == 4 for row in win._cell_values))
        bk.release_resources()


class CellRecordTypeTest(TestCase):
    """All BIFF8 cell record kinds resolve identically to a full Sheet."""

    def setUp(self):
        records = b''.join((
            labelsst_record(0, 2, 0, 0),
            mulrk_record(1, 1,
                         [(0, 0x00000008), (0, 0x0000000A), (1, 0x0000000C)]),
            boolerr_record(2, 0, 0, 1, 0),
            boolerr_record(2, 1, 0, 7, 1),
            formula_record(3, 0, 0, pack('<d', 3.25)),
            formula_record(3, 1, 0, b'\x01\x00\x01\x00\x00\x00\xff\xff'),
            formula_record(3, 2, 0, b'\x02\x00\x07\x00\x00\x00\xff\xff'),
            (formula_record(3, 3, 0, b'\x00' * 6 + b'\xff\xff')
             + string_record('fmla text')),
            mulblank_record(4, 0, [1, 1]),
            blank_record(4, 3, 1),
            rk_record(5, 0, 0, 0x00000010),
            rec(0x0204, pack('<HHH', 5, 1, 0) + pack('<HB', 3, 0) + b'abc'),
        ))
        self.stream = build_workbook(
            records, nrows=10, ncols=8, sst_strings=['shared hi'])

    def test_types_and_values_match_full_sheet(self):
        for fmt in (False, True):
            full = xlrd.open_workbook(
                file_contents=self.stream, formatting_info=fmt,
                on_demand=True).sheet_by_index(0)
            win = xlrd.open_workbook(
                file_contents=self.stream, formatting_info=fmt,
                on_demand=True).sheet_rows_window(0, 0, 10, 0, 5)
            self.assertIsInstance(win, RowWindow)
            for r in range(10):
                for c in range(5):
                    if r < full.nrows and c < full.ncols:
                        self.assertEqual(
                            win.cell_value(r, c), full.cell_value(r, c))
                        self.assertEqual(
                            win.cell_type(r, c), full.cell_type(r, c))
                        if fmt:
                            self.assertEqual(
                                win.cell_xf_index(r, c),
                                full.cell_xf_index(r, c))
                    else:
                        self.assertEqual(win.cell_type(r, c),
                                         XL_CELL_EMPTY)
                        self.assertEqual(win.cell_value(r, c), '')

    def test_formula_cached_values(self):
        win = xlrd.open_workbook(
            file_contents=self.stream, formatting_info=True,
            on_demand=True).sheet_rows_window(0, 3, 4, 0, 4)
        self.assertEqual(win.row_values(0),
                         [3.25, 1, 7, 'fmla text'])
        self.assertEqual(
            list(win.row_types(0)),
            [XL_CELL_NUMBER, XL_CELL_BOOLEAN, XL_CELL_ERROR, XL_CELL_TEXT])

    def test_blank_records_follow_formatting_flag(self):
        win = xlrd.open_workbook(
            file_contents=self.stream, formatting_info=True,
            on_demand=True).sheet_rows_window(0, 4, 5, 0, 4)
        self.assertEqual(list(win.row_types(0)),
                         [XL_CELL_BLANK, XL_CELL_BLANK,
                          XL_CELL_EMPTY, XL_CELL_BLANK])
        self.assertEqual(win.cell_xf_index(0, 0), 1)

        win_no_fmt = xlrd.open_workbook(
            file_contents=self.stream, formatting_info=False,
            on_demand=True).sheet_rows_window(0, 4, 5, 0, 4)
        self.assertTrue(
            all(t == XL_CELL_EMPTY for t in win_no_fmt.row_types(0)))
        self.assertEqual(win_no_fmt.cell_count, 0)

    def test_sparse_rows_are_empty(self):
        win = xlrd.open_workbook(
            file_contents=self.stream, on_demand=True
        ).sheet_rows_window(0, 6, 10, 0, 5)
        self.assertEqual(win.nrows, 4)
        for r in range(win.nrows):
            self.assertTrue(
                all(t == XL_CELL_EMPTY for t in win.row_types(r)))
            self.assertEqual(win.row_values(r), ['', '', '', '', ''])


class ResourceAndOnDemandTest(TestCase):
    """on_demand, unload_sheet and release_resources interoperation."""

    def setUp(self):
        records = b''.join(
            number_record(r, c, 0, float(r * 4 + c))
            for r in range(50) for c in range(4))
        self.stream = build_workbook(records, nrows=50, ncols=4)

    def test_window_never_registers_sheet(self):
        bk = xlrd.open_workbook(
            file_contents=self.stream, on_demand=True)
        self.assertFalse(bk.sheet_loaded(0))
        bk.sheet_rows_window(0, 0, 10, 0, 2)
        self.assertFalse(bk.sheet_loaded(0))
        bk.sheet_rows_window(0, 40, 50, 2, 4)
        self.assertFalse(bk.sheet_loaded(0))
        full = bk.sheet_by_index(0)
        self.assertTrue(bk.sheet_loaded(0))
        bk.unload_sheet(0)
        self.assertFalse(bk.sheet_loaded(0))
        # Resources retained: another window read still works.
        win = bk.sheet_rows_window(0, 5, 6, 0, 1)
        self.assertEqual(win.cell_value(0, 0), full.cell_value(5, 0))
        bk.release_resources()

    def test_repeated_window_reads_are_independent(self):
        bk = xlrd.open_workbook(
            file_contents=self.stream, on_demand=True)
        windows = [
            bk.sheet_rows_window(0, r, r + 5, 0, 4)
            for r in range(0, 25, 5)
        ]
        full = bk.sheet_by_index(0)
        for i, win in enumerate(windows):
            base = i * 5
            self.assertEqual(
                win.row_values(0), full.row_values(base, 0, 4))
            self.assertEqual(win.cell_count, 20)
        bk.release_resources()

    def test_release_resources_blocks_further_reads(self):
        bk = xlrd.open_workbook(
            file_contents=self.stream, on_demand=True)
        win = bk.sheet_rows_window(0, 0, 3, 0, 4)
        bk.release_resources()
        with self.assertRaises(xlrd.XLRDError):
            bk.sheet_rows_window(0, 0, 1, 0, 1)
        # Previously returned windows stay usable.
        self.assertEqual(win.cell_value(2, 3), 11.0)

    def test_non_on_demand_books_release_at_open(self):
        bk = xlrd.open_workbook(
            file_contents=self.stream, on_demand=False)
        self.assertTrue(bk._resources_released)
        with self.assertRaises(xlrd.XLRDError):
            bk.sheet_rows_window(0, 0, 1, 0, 1)

    def test_context_manager(self):
        with xlrd.open_workbook(
                file_contents=self.stream, on_demand=True) as bk:
            win = bk.sheet_rows_window(0, 0, 3, 0, 4)
            self.assertEqual(win.nrows, 3)
        self.assertTrue(bk._resources_released)
        self.assertEqual(win.cell_value(2, 3), 11.0)


class BoundsAndGuardsTest(TestCase):

    def setUp(self):
        self.stream = build_workbook(
            number_record(0, 0, 0, 1.0), nrows=1, ncols=1)

    def test_bad_ranges(self):
        bk = xlrd.open_workbook(
            file_contents=self.stream, on_demand=True)
        bad_ranges = [
            (-1, 1, 0, 1),
            (2, 1, 0, 1),
            (0, 1, -1, 1),
            (0, 65537, 0, 1),
            (0, 1, 0, 257),
        ]
        for rng in bad_ranges:
            with self.assertRaises(ValueError):
                bk.sheet_rows_window(0, *rng)
        with self.assertRaises(xlrd.XLRDError):
            bk.sheet_rows_window(9, 0, 1, 0, 1)
        bk.release_resources()

    def test_empty_window(self):
        bk = xlrd.open_workbook(
            file_contents=self.stream, on_demand=True)
        win = bk.sheet_rows_window(0, 5, 8, 0, 2)
        self.assertEqual((win.nrows, win.ncols), (3, 2))
        self.assertEqual(win.cell_count, 0)
        self.assertTrue(all(
            win.cell_type(r, c) == XL_CELL_EMPTY
            for r in range(3) for c in range(2)))
        bk.release_resources()

    def test_pre_biff8_guard(self):
        from .helpers import from_sample
        bk = xlrd.open_workbook(
            from_sample('biff4_no_format_no_window2.xls'))
        with self.assertRaisesRegex(xlrd.XLRDError, 'BIFF8'):
            bk.sheet_rows_window(0, 0, 1, 0, 1)
        bk.release_resources()


class SampleParityTest(TestCase):
    """Window reads match full-sheet slices on real shipped workbooks."""

    def _compare(self, path, fmt):
        bk = xlrd.open_workbook(
            path, formatting_info=fmt, on_demand=True)
        if bk.biff_version < 80:
            bk.release_resources()
            return
        for shx in range(bk.nsheets):
            full = bk.sheet_by_index(shx)
            rlo = full.nrows // 3
            rhi = max(rlo + 1, full.nrows - max(1, full.nrows // 4))
            clo = full.ncols // 3
            chi = max(clo + 1, full.ncols)
            rhi = min(rhi, full.nrows)
            chi = min(chi, full.ncols)
            bk.unload_sheet(shx)
            win = bk.sheet_rows_window(shx, rlo, rhi, clo, chi)
            for r in range(win.nrows):
                self.assertEqual(
                    win.row_values(r),
                    full.row_values(rlo + r, clo, chi))
                self.assertEqual(
                    list(win.row_types(r)),
                    list(full.row_types(rlo + r, clo, chi)))
                if fmt:
                    self.assertEqual(
                        [c.xf_index for c in win.row_slice(r)],
                        [c.xf_index
                         for c in full.row_slice(rlo + r, clo, chi)])
        bk.release_resources()

    def test_profiles(self):
        from .helpers import from_sample
        for fmt in (False, True):
            self._compare(from_sample('profiles.xls'), fmt)

    def test_formula_workbooks(self):
        from .helpers import from_sample
        for name in ('formula_test_names.xls', 'formula_test_sjmachin.xls',
                     'invalid_formula.xls', 'xf_class.xls'):
            for fmt in (False, True):
                self._compare(from_sample(name), fmt)


class FormattingOverrideTest(TestCase):

    def setUp(self):
        self.stream = build_workbook(
            blank_record(0, 0, 1) + number_record(0, 1, 0, 3.0),
            nrows=1, ncols=2)

    def test_window_can_disable_book_formatting(self):
        bk = xlrd.open_workbook(
            file_contents=self.stream, formatting_info=True, on_demand=True)
        win = bk.sheet_rows_window(
            0, 0, 1, 0, 2, formatting_info=False)
        self.assertEqual(list(win.row_types(0)),
                         [XL_CELL_EMPTY, XL_CELL_NUMBER])
        self.assertIsNone(win.cell(0, 0).xf_index)
        bk.release_resources()

    def test_window_can_enable_book_formatting(self):
        bk = xlrd.open_workbook(
            file_contents=self.stream, formatting_info=False, on_demand=True)
        win = bk.sheet_rows_window(
            0, 0, 1, 0, 2, formatting_info=True)
        self.assertEqual(list(win.row_types(0)),
                         [XL_CELL_BLANK, XL_CELL_NUMBER])
        self.assertEqual(win.cell_xf_index(0, 0), 1)
        bk.release_resources()


class FormattingMetadataTest(TestCase):
    """ROW/COLINFO fallbacks keep the same XF indexes as a full Sheet."""

    def setUp(self):
        records = b''.join((
            row_record(0, xf_index=2),
            colinfo_record(0, 1, 0x0BB8, 3),  # width 3000, fallback XF 3
            blank_record(0, 0, 1),
            number_record(1, 2, 0, 1.25),
            row_record(1, xf_index=2),
        ))
        self.stream = build_workbook(records, nrows=4, ncols=4)

    def test_row_and_col_fallbacks(self):
        full = xlrd.open_workbook(
            file_contents=self.stream, formatting_info=True,
            on_demand=True).sheet_by_index(0)
        win = xlrd.open_workbook(
            file_contents=self.stream, formatting_info=True,
            on_demand=True).sheet_rows_window(0, 0, 2, 0, 4)
        for r in range(2):
            for c in range(full.ncols):
                self.assertEqual(
                    win.cell_xf_index(r, c), full.cell_xf_index(r, c))
        # Empty cells beyond the materialised sheet width still use the
        # same row/column fallback chain as a padded full Sheet.
        self.assertEqual(win.cell_xf_index(0, 3), 2)  # row XF
        self.assertEqual(win.cell_xf_index(1, 0), 2)  # row XF beats COLINFO

    def test_computed_column_width(self):
        full = xlrd.open_workbook(
            file_contents=self.stream, formatting_info=True,
            on_demand=True).sheet_by_index(0)
        win = xlrd.open_workbook(
            file_contents=self.stream, formatting_info=True,
            on_demand=True).sheet_rows_window(0, 0, 2, 0, 4)
        for c in range(4):
            self.assertEqual(win.computed_column_width(c),
                             full.computed_column_width(c))

    def test_merged_cells_preserved(self):
        merged = rec(0x00E5, pack('<H', 1) + pack('<HHHH', 0, 1, 0, 2))
        stream = build_workbook(
            number_record(0, 0, 0, 1.0) + merged, nrows=2, ncols=2)
        win = xlrd.open_workbook(
            file_contents=stream, formatting_info=True,
            on_demand=True).sheet_rows_window(0, 0, 2, 0, 2)
        self.assertEqual(win.merged_cells, [(0, 2, 0, 3)])
