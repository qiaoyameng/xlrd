# -*- coding: utf-8 -*-
# Portions Copyright (C) 2010, Manfred Moitzi under a BSD licence

import struct
from unittest import TestCase

import xlrd
from xlrd.formula import FormulaRef

from .helpers import from_sample

try:
    ascii
except NameError:
    # For Python 2
    def ascii(s):
        a = repr(s)
        if a.startswith(('u"', "u'")):
            a = a[1:]
        return a

class TestFormulas(TestCase):

    def setUp(self):
        book = xlrd.open_workbook(from_sample('formula_test_sjmachin.xls'))
        self.sheet = book.sheet_by_index(0)

    def get_value(self, col, row):
        return ascii(self.sheet.col_values(col)[row])

    def test_cell_B2(self):
        self.assertEqual(
            self.get_value(1, 1),
            r"'\u041c\u041e\u0421\u041a\u0412\u0410 \u041c\u043e\u0441\u043a\u0432\u0430'",
        )

    def test_cell_B3(self):
        self.assertEqual(self.get_value(1, 2), '0.14285714285714285')

    def test_cell_B4(self):
        self.assertEqual(self.get_value(1, 3), "'ABCDEF'")

    def test_cell_B5(self):
        self.assertEqual(self.get_value(1, 4), "''")

    def test_cell_B6(self):
        self.assertEqual(self.get_value(1, 5), '1')

    def test_cell_B7(self):
        self.assertEqual(self.get_value(1, 6), '7')

    def test_cell_B8(self):
        self.assertEqual(
            self.get_value(1, 7),
            r"'\u041c\u041e\u0421\u041a\u0412\u0410 \u041c\u043e\u0441\u043a\u0432\u0430'",
        )

class TestNameFormulas(TestCase):

    def setUp(self):
        book = xlrd.open_workbook(from_sample('formula_test_names.xls'))
        self.sheet = book.sheet_by_index(0)

    def get_value(self, col, row):
        return ascii(self.sheet.col_values(col)[row])

    def test_unaryop(self):
        self.assertEqual(self.get_value(1, 1), '-7.0')

    def test_attrsum(self):
        self.assertEqual(self.get_value(1, 2), '4.0')

    def test_func(self):
        self.assertEqual(self.get_value(1, 3), '6.0')

    def test_func_var_args(self):
        self.assertEqual(self.get_value(1, 4), '3.0')

    def test_if(self):
        self.assertEqual(self.get_value(1, 5), "'b'")

    def test_choose(self):
        self.assertEqual(self.get_value(1, 6), "'C'")

    def test_evaluate_name_formula_with_invalid_operand(self):
        book = xlrd.open_workbook(from_sample('invalid_formula.xls'))
        sheet = book.sheet_by_index(0)
        cell = sheet.cell(0, 0)

        self.assertEqual(cell.ctype, xlrd.XL_CELL_ERROR)
        self.assertIn(cell.value, xlrd.error_text_from_code)


def _biff8_record(rtype, data):
    return struct.pack('<HH', rtype, len(data)) + data


def _biff8_bof(btype):
    return _biff8_record(0x0809, struct.pack('<HHHHII', 0x0600, btype, 0, 0, 0, 0))


def _biff8_ustr(text, lenlen=1):
    raw = text.encode('latin1')
    if lenlen == 1:
        return struct.pack('<BB', len(raw), 0) + raw
    return struct.pack('<HB', len(raw), 0) + raw


def _formula_record(rowx, colx, result_str, tokens):
    data = struct.pack('<HHH8sH', rowx, colx, 0, result_str, 0)
    data += b'\x00' * 4
    data += struct.pack('<H', len(tokens)) + tokens
    return _biff8_record(0x0006, data)


def _number_result(value):
    return struct.pack('<d', value)


def _special_result(first_byte, value=0):
    result = bytearray(8)
    result[0] = first_byte
    result[2] = value
    result[6] = 0xFF
    result[7] = 0xFF
    return bytes(result)


def make_shared_formula_book_bytes():
    # A minimal in-memory BIFF8 workbook (raw stream, no OLE wrapper):
    # Sheet1!B1:B2 share the formula =A<this row>*2, defined by a
    # SHRFMLA record attached to B1; B3..B6 exercise the string, boolean,
    # error and empty-string cached result forms; B7 holds a deleted
    # reference; B8 references Sheet2!A1; B9 uses a name defined in an
    # external workbook.
    texp_b1 = struct.pack('<BHH', 0x01, 0, 1)
    # ptgRefN(same row, one column left, both relative) ptgInt 2 ptgMul
    shared_tokens = (
        bytes([0x2C]) + struct.pack('<hH', 0, 0xC0FF)
        + bytes([0x1E]) + struct.pack('<H', 2)
        + bytes([0x05])
    )
    sheet1 = b''
    sheet1 += _biff8_bof(0x0010)
    sheet1 += _biff8_record(0x0200, struct.pack('<IIHHH', 0, 9, 0, 2, 0))
    sheet1 += _biff8_record(0x0203, struct.pack('<HHHd', 0, 0, 0, 1.0))
    sheet1 += _biff8_record(0x0203, struct.pack('<HHHd', 1, 0, 0, 2.0))
    sheet1 += _formula_record(0, 1, _number_result(2.0), texp_b1)
    sheet1 += _biff8_record(
        0x04BC,
        struct.pack('<HHBBxBH', 0, 1, 1, 1, 2, len(shared_tokens))
        + shared_tokens,
    )
    sheet1 += _formula_record(1, 1, _number_result(4.0), texp_b1)
    sheet1 += _formula_record(
        2, 1, _special_result(0), bytes([0x17, 0x02, 0x00]) + b'hi')
    sheet1 += _biff8_record(0x0207, struct.pack('<HB', 2, 0) + b'hi')
    sheet1 += _formula_record(3, 1, _special_result(1, 1), bytes([0x1D, 0x01]))
    sheet1 += _formula_record(
        4, 1, _special_result(2, 0x07), bytes([0x1C, 0x07]))
    sheet1 += _formula_record(
        5, 1, _special_result(3), bytes([0x17, 0x00, 0x00]))
    sheet1 += _formula_record(
        6, 1, _special_result(2, 0x17), bytes([0x2A]) + b'\xff' * 4)
    sheet1 += _formula_record(
        7, 1, _number_result(42.0),
        bytes([0x3A]) + struct.pack('<HHH', 0, 0, 0x0000))
    sheet1 += _formula_record(
        8, 1, _special_result(2, 0x1D),
        bytes([0x39]) + struct.pack('<HHH', 1, 1, 0))
    sheet1 += _biff8_record(0x000A, b'')

    sheet2 = b''
    sheet2 += _biff8_bof(0x0010)
    sheet2 += _biff8_record(0x0200, struct.pack('<IIHHH', 0, 1, 0, 1, 0))
    sheet2 += _biff8_record(0x0203, struct.pack('<HHHd', 0, 0, 0, 42.0))
    sheet2 += _biff8_record(0x000A, b'')

    boundsheet1 = _biff8_record(
        0x0085, struct.pack('<iBB', 0, 0, 0) + _biff8_ustr('Sheet1'))
    boundsheet2 = _biff8_record(
        0x0085, struct.pack('<iBB', 0, 0, 0) + _biff8_ustr('Sheet2'))
    supbook_internal = _biff8_record(0x01AE, struct.pack('<HH', 2, 0x0401))
    supbook_external = _biff8_record(
        0x01AE,
        struct.pack('<H', 1)
        + _biff8_ustr('other.xls', lenlen=2)
        + _biff8_ustr('Sheet1', lenlen=2),
    )
    externsheet = _biff8_record(
        0x0017,
        struct.pack('<H', 2)
        + struct.pack('<HHH', 0, 1, 1)
        + struct.pack('<HHH', 1, 0, 0),
    )
    globals_eof = _biff8_record(0x000A, b'')

    globals_head = _biff8_bof(0x0005)
    sheet1_offset = sum(map(len, [
        globals_head, boundsheet1, boundsheet2,
        supbook_internal, supbook_external, externsheet, globals_eof,
    ]))
    sheet2_offset = sheet1_offset + len(sheet1)
    boundsheet1 = _biff8_record(
        0x0085, struct.pack('<iBB', sheet1_offset, 0, 0)
        + _biff8_ustr('Sheet1'))
    boundsheet2 = _biff8_record(
        0x0085, struct.pack('<iBB', sheet2_offset, 0, 0)
        + _biff8_ustr('Sheet2'))

    return b''.join([
        globals_head, boundsheet1, boundsheet2,
        supbook_internal, supbook_external, externsheet, globals_eof,
        sheet1, sheet2,
    ])


class TestFormulaInfoShared(TestCase):

    @classmethod
    def setUpClass(cls):
        cls.book_bytes = make_shared_formula_book_bytes()

    def open_sheet(self, **kwargs):
        book = xlrd.open_workbook(file_contents=self.book_bytes, **kwargs)
        return book.sheet_by_index(0)

    def test_default_mode_collects_nothing(self):
        sheet = self.open_sheet()
        self.assertEqual(sheet.formula_info_map, {})
        self.assertIsNone(sheet.cell_formula_info(0, 1))

    def test_cached_values_unchanged_by_option(self):
        plain = self.open_sheet()
        preserved = self.open_sheet(preserve_formula_info=True)
        for rowx in range(plain.nrows):
            for colx in range(plain.ncols):
                self.assertEqual(
                    plain.cell_type(rowx, colx),
                    preserved.cell_type(rowx, colx),
                )
                self.assertEqual(
                    plain.cell_value(rowx, colx),
                    preserved.cell_value(rowx, colx),
                )

    def test_shared_formula_expands_per_cell(self):
        sheet = self.open_sheet(preserve_formula_info=True)
        info1 = sheet.cell_formula_info(0, 1)
        info2 = sheet.cell_formula_info(1, 1)
        self.assertEqual(info1.shared_range, (0, 1, 1, 1))
        self.assertEqual(info2.shared_range, (0, 1, 1, 1))
        self.assertEqual(info1.shared_anchor, (0, 1))
        self.assertEqual(info2.shared_anchor, (0, 1))
        self.assertEqual(info1.shared_tokens, info2.shared_tokens)
        # Same shared definition, different relative references.
        self.assertEqual(
            info1.references,
            (FormulaRef('cell', 0, 0, 0, 0, 0, 0),),
        )
        self.assertEqual(
            info2.references,
            (FormulaRef('cell', 0, 0, 1, 1, 0, 0),),
        )
        self.assertEqual(info1.unresolved, ())
        self.assertEqual(info2.unresolved, ())

    def test_cached_result_types(self):
        sheet = self.open_sheet(preserve_formula_info=True)
        self.assertEqual(sheet.cell_formula_info(0, 1).cached_type, xlrd.XL_CELL_NUMBER)
        self.assertEqual(sheet.cell_formula_info(2, 1).cached_type, xlrd.XL_CELL_TEXT)
        self.assertEqual(sheet.cell_formula_info(2, 1).rowx, 2)
        self.assertEqual(sheet.cell_value(2, 1), 'hi')
        self.assertEqual(sheet.cell_formula_info(3, 1).cached_type, xlrd.XL_CELL_BOOLEAN)
        self.assertEqual(sheet.cell_formula_info(4, 1).cached_type, xlrd.XL_CELL_ERROR)
        self.assertEqual(sheet.cell_formula_info(5, 1).cached_type, xlrd.XL_CELL_TEXT)
        self.assertEqual(sheet.cell_value(5, 1), '')

    def test_deleted_reference_is_unresolved(self):
        sheet = self.open_sheet(preserve_formula_info=True)
        info = sheet.cell_formula_info(6, 1)
        self.assertEqual(info.references, ())
        self.assertNotEqual(info.unresolved, ())

    def test_cross_sheet_reference_keeps_sheet_identity(self):
        sheet = self.open_sheet(preserve_formula_info=True)
        info = sheet.cell_formula_info(7, 1)
        self.assertEqual(
            info.references,
            (FormulaRef('cell', 1, 1, 0, 0, 0, 0),),
        )
        self.assertEqual(info.unresolved, ())

    def test_external_name_is_unresolved(self):
        sheet = self.open_sheet(preserve_formula_info=True)
        info = sheet.cell_formula_info(8, 1)
        self.assertEqual(info.references, ())
        self.assertNotEqual(info.unresolved, ())


class TestFormulaInfoExistingFiles(TestCase):

    def check_values_unchanged(self, filename):
        plain = xlrd.open_workbook(from_sample(filename))
        preserved = xlrd.open_workbook(
            from_sample(filename), preserve_formula_info=True)
        self.assertEqual(plain.nsheets, preserved.nsheets)
        for sheetx in range(plain.nsheets):
            plain_sheet = plain.sheet_by_index(sheetx)
            preserved_sheet = preserved.sheet_by_index(sheetx)
            self.assertEqual(plain_sheet.formula_info_map, {})
            for rowx in range(plain_sheet.nrows):
                for colx in range(plain_sheet.ncols):
                    self.assertEqual(
                        plain_sheet.cell_type(rowx, colx),
                        preserved_sheet.cell_type(rowx, colx),
                    )
                    self.assertEqual(
                        plain_sheet.cell_value(rowx, colx),
                        preserved_sheet.cell_value(rowx, colx),
                    )
        return preserved

    def test_sjmachin_fixture(self):
        book = self.check_values_unchanged('formula_test_sjmachin.xls')
        sheet = book.sheet_by_index(0)
        self.assertNotEqual(sheet.formula_info_map, {})
        for info in sheet.formula_info_map.values():
            self.assertEqual(info.unresolved, ())

    def test_names_fixture(self):
        book = self.check_values_unchanged('formula_test_names.xls')
        sheet = book.sheet_by_index(0)
        info = sheet.cell_formula_info(1, 1)
        self.assertEqual(
            info.references,
            (FormulaRef('name', name='unaryminus'),),
        )
        self.assertEqual(info.unresolved, ())
