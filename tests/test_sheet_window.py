import struct
from unittest import TestCase

import xlrd
from xlrd import sheet as sheet_module
from xlrd.biffh import XLRDError

from .helpers import from_sample


ROW_COUNT = 400
COL_COUNT = 4


def _record(code, data=b''):
    return struct.pack('<HH', code, len(data)) + data


def _bof(stream_type):
    return _record(
        0x0809,
        struct.pack('<HHHHHH', 0x0600, stream_type, 0, 0, 0, 0),
    )


def _number(rowx, colx, xf_index, value):
    return _record(
        0x0203,
        struct.pack('<HHHd', rowx, colx, xf_index, value),
    )


def _label(rowx, colx, xf_index, value):
    raw = value.encode('latin_1')
    return _record(
        0x0204,
        struct.pack('<HHH', rowx, colx, xf_index) +
        struct.pack('<H', len(raw)) + b'\x00' + raw,
    )


def _blank(rowx, colx, xf_index):
    return _record(0x0201, struct.pack('<HHH', rowx, colx, xf_index))


def _mulblank(rowx, first_colx, last_colx, xf_indexes):
    return _record(
        0x00BE,
        struct.pack('<HH', rowx, first_colx) +
        b''.join(struct.pack('<H', xfx) for xfx in xf_indexes) +
        struct.pack('<H', last_colx),
    )


def _formula_number(rowx, colx, xf_index, value):
    return _record(
        0x0006,
        struct.pack('<HHH', rowx, colx, xf_index) +
        struct.pack('<d', value) + struct.pack('<H', 0) + b'\x00' * 6,
    )


def _formula_text(rowx, colx, xf_index, value):
    raw = value.encode('latin_1')
    formula = (
        struct.pack('<HHH', rowx, colx, xf_index) +
        b'\x00\x00\x00\x00\x00\x00\xff\xff' +
        struct.pack('<H', 0) + b'\x00' * 6
    )
    string = _record(
        0x0207,
        struct.pack('<H', len(raw)) + b'\x00' + raw,
    )
    return _record(0x0006, formula) + string


def make_biff8_workbook():
    """
    Build a small raw BIFF8 workbook.

    Row 0 is written after all other cell records, so any reader that assumes
    records are sorted by row will miss it.  Rows 1 and 2 exercise formula
    cached values, BLANK/MULBLANK records and a sparse row.
    """
    name = b'Sheet1'
    globals_bof = _bof(0x0005)
    placeholder = (
        struct.pack('<iBB', 0, 0, 0) +
        bytes([len(name), 0]) + name
    )
    sheet_offset = len(
        globals_bof + _record(0x0085, placeholder) + _record(0x000A))
    boundsheet = (
        struct.pack('<iBB', sheet_offset, 0, 0) +
        bytes([len(name), 0]) + name
    )
    globals_stream = (
        globals_bof + _record(0x0085, boundsheet) + _record(0x000A)
    )

    records = [
        _bof(0x0010),
        _record(0x0200, struct.pack(
            '<IIHHH', 0, ROW_COUNT, 0, COL_COUNT, 0)),
    ]
    for rowx in range(1, ROW_COUNT):
        records.append(_number(rowx, 0, 0, float(rowx)))

    records.append(_number(0, 0, 0, 12345.0))
    records.append(_blank(1, 1, 0))
    records.append(_mulblank(1, 2, 3, [0, 0]))
    records.append(_formula_number(1, 0, 0, 99.0))
    records.append(_formula_text(2, 0, 0, 'cached'))
    records.append(_label(2, 3, 0, 'sparse'))
    records.append(_record(0x000A))
    return globals_stream + b''.join(records)


class TestSheetWindow(TestCase):

    def setUp(self):
        self.contents = make_biff8_workbook()

    def _open_window_book(self, formatting_info=True):
        return xlrd.open_workbook(
            file_contents=self.contents,
            on_demand=True,
            formatting_info=formatting_info,
        )

    def _open_full_book(self, formatting_info=True):
        return xlrd.open_workbook(
            file_contents=self.contents,
            formatting_info=formatting_info,
        )

    def test_window_matches_full_sheet_slices(self):
        for formatting_info in (False, True):
            window_book = self._open_window_book(formatting_info)
            full_sheet = self._open_full_book(
                formatting_info).sheet_by_index(0)
            window = window_book.sheet_by_index_window(0, 0, 3, 0, 4)
            self.assertIsInstance(window, list)
            self.assertEqual((window.nrows, window.ncols), (3, 4))
            self.assertEqual(
                window,
                [full_sheet.row_slice(rowx, 0, 4) for rowx in range(3)],
            )
            for rowx in range(3):
                for colx in range(4):
                    got = window.cell(rowx, colx)
                    expected = full_sheet.cell(rowx, colx)
                    self.assertEqual(got.ctype, expected.ctype)
                    self.assertEqual(got.value, expected.value)
                    self.assertEqual(got.xf_index, expected.xf_index)

        self.assertEqual(window.cell_value(0, 0), 12345.0)
        self.assertEqual(window.cell_value(1, 0), 99.0)
        self.assertEqual(window.cell_value(2, 0), 'cached')
        self.assertEqual(window.cell_value(2, 3), 'sparse')
        self.assertEqual(window.cell_type(2, 1), xlrd.XL_CELL_EMPTY)
        self.assertEqual(window.cell_type(2, 2), xlrd.XL_CELL_EMPTY)

    def test_only_window_cells_are_instantiated(self):
        book = self._open_window_book(formatting_info=False)
        created = []
        original_cell = sheet_module.Cell

        class CountingCell(original_cell):
            def __init__(self, *args, **kwargs):
                original_cell.__init__(self, *args, **kwargs)
                created.append(self)

        sheet_module.Cell = CountingCell
        try:
            window = book.sheet_by_index_window(0, 1, 3, 0, 2)
        finally:
            sheet_module.Cell = original_cell

        self.assertEqual(len(created), 2 * 2)
        self.assertEqual((window.nrows, window.ncols), (2, 2))
        self.assertEqual(window.cell_value(0, 0), 99.0)
        self.assertEqual(window.cell_value(1, 1), '')

    def test_multiple_windows_and_resource_release(self):
        book = self._open_window_book()
        first = book.sheet_by_index_window(0, 0, 1, 0, 1)
        second = book.sheet_by_index_window(0, 5, 7, 0, 1)
        third = book.sheet_by_index_window(0, 0, 1, 0, 1)
        self.assertEqual(first.cell_value(0, 0), 12345.0)
        self.assertEqual(second.cell_value(0, 0), 5.0)
        self.assertEqual(third.cell_value(0, 0), 12345.0)
        book.release_resources()
        with self.assertRaises(XLRDError):
            book.sheet_by_index_window(0, 0, 1, 0, 1)

    def test_loaded_sheet_fallback_after_release(self):
        book = self._open_full_book(formatting_info=True)
        window = book.sheet_by_index_window(0, 1, 3, 0, 2)
        self.assertEqual(window.cell_value(0, 0), 99.0)
        self.assertEqual(window.cell_value(1, 0), 'cached')
        self.assertEqual(window.cell_value(0, 1), '')

    def test_rejects_non_biff8(self):
        book = xlrd.open_workbook(
            from_sample('biff4_no_format_no_window2.xls'), on_demand=True)
        with self.assertRaises(XLRDError):
            book.sheet_by_index_window(0, 0, 1, 0, 1)
