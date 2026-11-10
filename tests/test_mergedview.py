# -*- coding: utf-8 -*-

import struct
from unittest import TestCase

import pytest

import xlrd
from xlrd import (
    XL_CELL_BLANK, XL_CELL_EMPTY, XL_CELL_TEXT,
    MergedCellsError, MergedCellView,
)

from .helpers import from_sample


# --- Minimal in-memory BIFF8 workbook construction ---------------------

def _rec(code, data):
    return struct.pack('<HH', code, len(data)) + data


def _bof(stream_type):
    return _rec(0x0809, struct.pack('<HHHHII', 0x0600, stream_type,
                                    0x0DBB, 0x07CC, 0x41, 0x06))


def _boundsheet(pos, name):
    encoded = name.encode('utf_16_le')
    return _rec(0x0085, struct.pack('<iBB', pos, 0, 0)
                + struct.pack('<BB', len(name), 1) + encoded)


def _label(rowx, colx, text, xf=0):
    encoded = text.encode('utf_16_le')
    return _rec(0x0204, struct.pack('<HHH', rowx, colx, xf)
                + struct.pack('<HB', len(text), 1) + encoded)


def _number(rowx, colx, value, xf=0):
    return _rec(0x0203, struct.pack('<HHHd', rowx, colx, xf, value))


def _blank(rowx, colx, xf=0):
    return _rec(0x0201, struct.pack('<HHH', rowx, colx, xf))


def _row_record(rowx, hidden=False):
    grbit = 0x20 if hidden else 0
    return _rec(0x0208, struct.pack('<HHHHHHHH',
                                    rowx, 0, 16, 255, 0, 0, grbit, 0x0F))


def _colinfo(first, last, hidden=False):
    flags = 1 if hidden else 0
    return _rec(0x007D, struct.pack('<HHHHHH',
                                    first, last, 2048, 0x0F, flags, 0))


def _dimensions(nrows, ncols):
    return _rec(0x0200, struct.pack('<IIHHH', 0, nrows, 0, ncols, 0))


def _mergedcells(ranges):
    data = struct.pack('<H', len(ranges))
    for rlo, rhi, clo, chi in ranges:
        data += struct.pack('<HHHH', rlo, rhi - 1, clo, chi - 1)
    return _rec(0x00E5, data)


def make_book_bytes(sheet_records, nrows, ncols, sheet_name='S1'):
    sheet_stream = b''.join(
        [_bof(0x0010), _dimensions(nrows, ncols)]
        + list(sheet_records)
        + [_rec(0x000A, b'')]
    )
    globals_head = b''.join([
        _bof(0x0005),
        _rec(0x0042, struct.pack('<H', 1200)),
    ])
    offset = len(globals_head) + len(_boundsheet(0, sheet_name)) + 4
    globals_stream = b''.join([
        globals_head,
        _boundsheet(offset, sheet_name),
        _rec(0x000A, b''),
    ])
    return globals_stream + sheet_stream


def open_synthetic(sheet_records, nrows, ncols, **kwargs):
    kwargs.setdefault('formatting_info', True)
    return xlrd.open_workbook(
        file_contents=make_book_bytes(sheet_records, nrows, ncols),
        **kwargs)


# --- Tests against the shipped XLS fixtures ----------------------------

class TestFixtureMergedView(TestCase):

    def setUp(self):
        self.book = xlrd.open_workbook(from_sample('xf_class.xls'),
                                       formatting_info=True)
        self.sheet = self.book.sheet_by_name('table2')
        # Fixture merge: rows 3..6, cols 2..4, anchor (3, 2) == 'MERGED'
        self.view = self.sheet.merged_view()

    def test_anchor_and_covered_positions(self):
        anchor = self.view.info(3, 2)
        self.assertEqual(anchor.value, 'MERGED')
        self.assertEqual(anchor.ctype, XL_CELL_TEXT)
        self.assertTrue(anchor.is_merged)
        self.assertEqual((anchor.anchor_rowx, anchor.anchor_colx), (3, 2))

        covered = self.view.info(5, 4)
        self.assertEqual(covered.value, 'MERGED')
        # The position's own raw cell type is preserved (blank here).
        self.assertEqual(covered.ctype, XL_CELL_BLANK)
        self.assertTrue(covered.is_merged)
        self.assertEqual((covered.anchor_rowx, covered.anchor_colx), (3, 2))

    def test_unmerged_position_is_its_own_anchor(self):
        info = self.view.info(0, 0)
        self.assertFalse(info.is_merged)
        self.assertEqual((info.anchor_rowx, info.anchor_colx), (0, 0))
        self.assertEqual(info.value, self.sheet.cell_value(0, 0))

    def test_raw_sheet_api_untouched(self):
        before = [self.sheet.row_values(r) for r in range(self.sheet.nrows)]
        # Exercise the view thoroughly.
        list(self.view.iter_window(0, self.sheet.nrows, 0, self.sheet.ncols))
        after = [self.sheet.row_values(r) for r in range(self.sheet.nrows)]
        self.assertEqual(before, after)
        # The covered cell still reads empty through the classic API.
        self.assertEqual(self.sheet.cell_value(4, 3), '')
        self.assertEqual(self.sheet.row_values(4)[3], '')

    def test_window_excluding_anchor_still_gets_value(self):
        # Rows 4..6 of the merge: the anchor row 3 is outside.
        values = self.view.window_values(4, 7, 2, 5)
        self.assertEqual(values, [['MERGED'] * 3] * 3)

    def test_window_clipped_to_sheet_bounds(self):
        cells = list(self.view.iter_window(-10, 999, -10, 999))
        self.assertEqual(len(cells), self.sheet.nrows * self.sheet.ncols)
        self.assertEqual((cells[0].rowx, cells[0].colx), (0, 0))
        self.assertEqual((cells[-1].rowx, cells[-1].colx),
                         (self.sheet.nrows - 1, self.sheet.ncols - 1))
        empty = list(self.view.iter_window(100, 200, 0, 3))
        self.assertEqual(empty, [])

    def test_anchor_and_position_formats_distinguished(self):
        info = self.view.info(4, 3)
        self.assertEqual(info.anchor_xf_index,
                         self.sheet.cell_xf_index(3, 2))
        self.assertEqual(info.xf_index, self.sheet.cell_xf_index(4, 3))

    def test_formate_fixture(self):
        book = xlrd.open_workbook(from_sample('Formate.xls'),
                                  formatting_info=True)
        sheet = book.sheet_by_name('ÖÄÜ')
        view = sheet.merged_view()
        self.assertEqual(view.info(4, 4).value, 'MERGED CELLS')
        self.assertEqual((view.info(4, 4).anchor_rowx,
                          view.info(4, 4).anchor_colx), (2, 2))


# --- Tests against synthetic in-memory BIFF8 workbooks -----------------

class TestSyntheticMergedView(TestCase):

    def test_empty_anchor(self):
        # Anchor is a BLANK record: the logical value is '' everywhere.
        book = open_synthetic(
            [_blank(1, 1, xf=7),
             _mergedcells([(1, 3, 1, 4)])],
            nrows=4, ncols=5)
        view = book.sheet_by_index(0).merged_view()
        info = view.info(2, 3)
        self.assertTrue(info.is_merged)
        self.assertEqual(info.value, '')
        self.assertEqual(info.ctype, XL_CELL_EMPTY)
        self.assertEqual(info.anchor_xf_index, 7)

    def test_anchor_and_position_formats_differ(self):
        book = open_synthetic(
            [_label(0, 0, 'Title', xf=11),
             _blank(1, 1, xf=22),
             _mergedcells([(0, 2, 0, 2)])],
            nrows=3, ncols=3)
        view = book.sheet_by_index(0).merged_view()
        info = view.info(1, 1)
        self.assertEqual(info.value, 'Title')
        self.assertEqual(info.anchor_xf_index, 11)
        self.assertEqual(info.xf_index, 22)

    def test_hidden_rows_and_columns(self):
        book = open_synthetic(
            [_row_record(1, hidden=True),
             _colinfo(2, 2, hidden=True),
             _label(0, 0, 'Head'),
             _mergedcells([(0, 3, 0, 3)])],
            nrows=4, ncols=4)
        view = book.sheet_by_index(0).merged_view()
        hidden = view.info(1, 2)
        self.assertTrue(hidden.row_hidden)
        self.assertTrue(hidden.col_hidden)
        self.assertEqual(hidden.value, 'Head')
        visible = view.info(0, 0)
        self.assertFalse(visible.row_hidden)
        self.assertFalse(visible.col_hidden)

    def test_overlapping_ranges_raise_at_construction(self):
        for ranges in ([(0, 3, 0, 3), (2, 5, 2, 5)],
                       [(2, 5, 2, 5), (0, 3, 0, 3)]):
            book = open_synthetic([_mergedcells(ranges)],
                                  nrows=6, ncols=6)
            sheet = book.sheet_by_index(0)
            with pytest.raises(MergedCellsError, match='overlap'):
                sheet.merged_view()

    def test_out_of_bounds_ranges_raise_at_construction(self):
        # Column past the 256-column BIFF8 grid, via the file itself.
        book = open_synthetic([_mergedcells([(0, 1, 250, 300)])],
                              nrows=2, ncols=4)
        sheet = book.sheet_by_index(0)
        with pytest.raises(MergedCellsError, match='outside'):
            sheet.merged_view()
        # Row past the grid and degenerate ranges, injected post-load
        # (xlrd's loader would reject these while tidying dimensions).
        for bad in [(0, 65537, 0, 1), (2, 2, 0, 1), (1, 2, 3, 3)]:
            sheet.merged_cells = [bad]
            with pytest.raises(MergedCellsError):
                MergedCellView(sheet)

    def test_ragged_rows(self):
        # Row 1 is short; the merge reaches past its last stored cell.
        book = open_synthetic(
            [_label(0, 0, 'Wide'),
             _number(1, 0, 1.5),
             _mergedcells([(0, 3, 0, 4)])],
            nrows=3, ncols=4, ragged_rows=True)
        sheet = book.sheet_by_index(0)
        # Row 1 really is ragged (only one stored cell).
        self.assertEqual(sheet.row_len(1), 1)
        view = sheet.merged_view()
        info = view.info(1, 3)
        self.assertEqual(info.value, 'Wide')
        self.assertEqual(info.ctype, XL_CELL_EMPTY)
        self.assertIsInstance(info.xf_index, int)

    def test_on_demand_view_survives_release_resources(self):
        book = open_synthetic(
            [_label(0, 0, 'Persist'),
             _mergedcells([(0, 2, 0, 2)])],
            nrows=3, ncols=3, on_demand=True)
        sheet = book.sheet_by_index(0)
        view = sheet.merged_view()
        book.release_resources()
        # The view only needs the loaded sheet's in-memory arrays.
        self.assertEqual(view.info(1, 1).value, 'Persist')
        self.assertEqual(view.window_values(0, 2, 0, 2),
                         [['Persist', 'Persist'], ['Persist', 'Persist']])
        self.assertEqual(sheet.cell_value(1, 1), '')

    def test_view_survives_with_block_exit(self):
        with open_synthetic(
                [_label(0, 0, 'Ctx'),
                 _mergedcells([(0, 2, 0, 2)])],
                nrows=2, ncols=2) as book:
            view = book.sheet_by_index(0).merged_view()
        self.assertEqual(view.info(1, 0).value, 'Ctx')

    def test_without_formatting_info(self):
        book = open_synthetic(
            [_label(0, 0, 'Plain')],
            nrows=2, ncols=2, formatting_info=False)
        view = book.sheet_by_index(0).merged_view()
        info = view.info(0, 0)
        self.assertEqual(info.value, 'Plain')
        self.assertIsNone(info.xf_index)
        self.assertIsNone(info.anchor_xf_index)
        self.assertFalse(info.row_hidden)
        self.assertFalse(info.col_hidden)

    def test_index_reused_across_many_windows(self):
        ranges = [(r, r + 1, 0, 8) for r in range(0, 40, 2)]
        records = [_label(r, 0, 'Band%d' % r) for r in range(0, 40, 2)]
        records.append(_blank(39, 0))
        records.append(_mergedcells(ranges))
        book = open_synthetic(records, nrows=40, ncols=8)
        view = book.sheet_by_index(0).merged_view()
        # Repeated window queries hit the prebuilt per-row index.
        for _unused in range(50):
            for r in range(0, 40, 2):
                values = view.window_values(r, r + 2, 0, 8)
                self.assertEqual(values[0], ['Band%d' % r] * 8)
                self.assertEqual(values[1], [''] * 8)
