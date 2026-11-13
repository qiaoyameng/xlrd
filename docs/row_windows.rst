Reading row windows
===================

.. currentmodule:: xlrd.book

New in version 2.1.0.

Backends that only need a few hundred rows of an old ``.xls`` file should
not have to pay for materialising the whole worksheet.  The
:meth:`Book.sheet_rows_window` method reads a rectangular window of a
worksheet in a single sequential pass over its BIFF8 records and
materialises only the cells whose ``(row, column)`` address falls inside
the requested half-open ranges::

    import xlrd

    with xlrd.open_workbook('huge.xls', on_demand=True) as book:
        window = book.sheet_rows_window(0, 100, 200, 0, 8)
        for rowx in range(window.nrows):
            print(window.row_values(rowx))

The row range ``[row_lo, row_hi)`` and column range ``[col_lo, col_hi)``
use absolute workbook coordinates; the returned
:class:`xlrd.window.RowWindow` uses window-local coordinates, so row 0 of
the window is workbook row ``row_lo``.

Behaviour notes:

* Only BIFF8 workbooks are supported; older files raise
  :class:`xlrd.XLRDError`.
* Workbook-level shared metadata (the shared string table, XF/FORMAT/FONT
  tables, palette, ...) is loaded exactly as for a normal open; the memory
  saving applies to cells outside the window.
* BIFF records are not assumed to be sorted by row: cells that appear
  later in the stream but fall inside the window are still returned, as
  are formula cached values, ``BLANK``/``MULBLANK`` records and sparse
  rows.
* Cell types, values and (with ``formatting_info=True``) XF indexes match
  what slicing the corresponding ``Book.sheet_by_index`` sheet would
  produce.
* The workbook must be opened with ``on_demand=True`` (and
  :meth:`Book.release_resources` not yet called) so the workbook stream is
  still available for the scan.  No :class:`~xlrd.sheet.Sheet` object is
  created or registered, so window reads compose freely with
  :meth:`Book.unload_sheet` and with further window reads of the same
  sheet.
