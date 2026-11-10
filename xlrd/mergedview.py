# -*- coding: utf-8 -*-
# This module is part of the xlrd package, which is released under a
# BSD-style licence.

"""
An optional, read-only merged-cell view over a loaded worksheet.

When a merged title spans several rows and columns, the raw
:meth:`~xlrd.sheet.Sheet.cell` / :meth:`~xlrd.sheet.Sheet.row_values`
API only reports the value in the top-left (anchor) cell; every other
position in the merged rectangle reads as an empty cell. This module
provides :class:`MergedCellView`, which resolves each position in a
rectangular query window to the logical value of the covering merged
range's anchor, while still reporting the position's own raw cell type
and keeping both formatting sources (anchor vs. position) apart.

The view never writes back into the sheet's underlying arrays, so
``cell()`` and ``row_values()`` on the sheet are unaffected.
"""

from __future__ import print_function

from bisect import bisect_right

from .biffh import XL_CELL_EMPTY, XLRDError
from .timemachine import UNICODE_LITERAL, xrange


class MergedCellsError(XLRDError):
    """
    Raised when a :class:`MergedCellView` is built over a sheet whose
    ``merged_cells`` ranges are structurally invalid: degenerate,
    outside the sheet's addressable grid, or overlapping each other.

    The error is raised eagerly when the view is constructed, so the
    outcome never depends on traversal order.
    """


class MergedCellInfo(object):
    """
    Resolved information for one grid position, as produced by
    :class:`MergedCellView`.

    .. versionadded:: 2.1.0
    """

    __slots__ = [
        'rowx', 'colx',
        'value', 'ctype',
        'anchor_rowx', 'anchor_colx', 'is_merged',
        'xf_index', 'anchor_xf_index',
        'row_hidden', 'col_hidden',
    ]

    def __init__(self, rowx, colx, value, ctype,
                 anchor_rowx, anchor_colx, is_merged,
                 xf_index, anchor_xf_index,
                 row_hidden, col_hidden):
        self.rowx = rowx
        self.colx = colx
        #: Logical value: the anchor cell's value if this position is
        #: covered by a merged range, otherwise the position's own value.
        self.value = value
        #: The raw cell type recorded at this exact position
        #: (e.g. ``XL_CELL_EMPTY`` for a covered non-anchor cell).
        self.ctype = ctype
        #: Row index of the anchor (top-left) cell of the covering
        #: merged range; equal to ``rowx`` when not merged.
        self.anchor_rowx = anchor_rowx
        #: Column index of the anchor cell; equal to ``colx`` when
        #: not merged.
        self.anchor_colx = anchor_colx
        #: ``True`` if this position is covered by a merged range.
        self.is_merged = is_merged
        #: XF index of the cell at this exact position, or ``None``
        #: unless the book was opened with ``formatting_info=True``.
        self.xf_index = xf_index
        #: XF index of the anchor cell (the formatting a renderer
        #: applies to the whole merged rectangle), or ``None`` unless
        #: the book was opened with ``formatting_info=True``.
        self.anchor_xf_index = anchor_xf_index
        #: ``True`` if the row is hidden (always ``False`` without
        #: ``formatting_info=True``).
        self.row_hidden = row_hidden
        #: ``True`` if the column is hidden (always ``False`` without
        #: ``formatting_info=True``).
        self.col_hidden = col_hidden

    def __repr__(self):
        return ("%s(%d, %d, value=%r, ctype=%d, anchor=(%d, %d))"
                % (self.__class__.__name__, self.rowx, self.colx,
                   self.value, self.ctype,
                   self.anchor_rowx, self.anchor_colx))


class MergedCellView(object):
    """
    A read-only view resolving merged cells over a loaded
    :class:`~xlrd.sheet.Sheet`.

    Create via :meth:`xlrd.sheet.Sheet.merged_view` or directly with
    ``MergedCellView(sheet)``. Construction validates every range in
    ``sheet.merged_cells`` and raises :class:`MergedCellsError` on
    overlapping or out-of-bounds ranges; a per-row index is built once
    and reused by all queries, so repeated window queries do not
    rescan the whole merged-range table per cell.

    .. note::
       The view reads only the sheet's in-memory cell arrays, so it
       stays usable for an already-loaded sheet after
       :meth:`xlrd.book.Book.release_resources` (including exiting an
       ``open_workbook`` ``with`` block). It is bound to the
       :class:`~xlrd.sheet.Sheet` object it was built from: if the
       book later unloads and reloads that sheet (``on_demand`` mode),
       build a fresh view from the new sheet object.

    .. versionadded:: 2.1.0
    """

    def __init__(self, sheet):
        self._sheet = sheet
        self._formatting_info = bool(sheet.formatting_info)
        self._cell_values = sheet._cell_values
        self._cell_types = sheet._cell_types
        self._cell_xf_indexes = sheet._cell_xf_indexes
        ranges = self._validate(sheet)
        #: Validated merged ranges, sorted by ``(rlo, rhi, clo, chi)``.
        self.merged_ranges = sorted(ranges)
        # Reusable per-row index: rowx -> list of
        # (clo, chi, anchor_rowx, anchor_colx) sorted by clo, plus a
        # parallel list of the clo keys for bisecting.
        buckets = {}
        for rlo, rhi, clo, chi in self.merged_ranges:
            entry = (clo, chi, rlo, clo)
            for rowx in xrange(rlo, rhi):
                bucket = buckets.get(rowx)
                if bucket is None:
                    buckets[rowx] = [entry]
                else:
                    bucket.append(entry)
        self._row_spans = {}
        self._row_span_los = {}
        for rowx, entries in buckets.items():
            entries.sort()
            self._row_spans[rowx] = entries
            self._row_span_los[rowx] = [e[0] for e in entries]

    @staticmethod
    def _validate(sheet):
        ranges = list(sheet.merged_cells)
        umaxrows = sheet.utter_max_rows
        umaxcols = sheet.utter_max_cols
        for crange in ranges:
            rlo, rhi, clo, chi = crange
            if not (0 <= rlo < rhi <= umaxrows
                    and 0 <= clo < chi <= umaxcols):
                raise MergedCellsError(
                    "sheet #%d (%r): merged range %r is degenerate or "
                    "outside the addressable grid "
                    "(rows 0..%d, cols 0..%d)"
                    % (sheet.number, sheet.name, crange,
                       umaxrows, umaxcols))
        # Sweep-line overlap check, independent of input order.
        ordered = sorted(ranges)
        active = []  # ranges whose rhi is beyond the current rlo
        for crange in ordered:
            rlo, rhi, clo, chi = crange
            active = [other for other in active if other[1] > rlo]
            for other in active:
                if other[2] < chi and clo < other[3]:
                    raise MergedCellsError(
                        "sheet #%d (%r): merged ranges %r and %r overlap"
                        % (sheet.number, sheet.name, other, crange))
            active.append(crange)
        return ranges

    def _anchor_for(self, rowx, colx):
        # Return (anchor_rowx, anchor_colx) covering (rowx, colx),
        # or None if the position is not inside a merged range.
        entries = self._row_spans.get(rowx)
        if not entries:
            return None
        los = self._row_span_los[rowx]
        # Walk backwards through entries with clo <= colx; entries are
        # non-overlapping so the first match is the only possible one.
        for i in xrange(bisect_right(los, colx) - 1, -1, -1):
            entry = entries[i]
            if colx < entry[1]:
                return entry[2], entry[3]
        return None

    def _raw_type_value(self, rowx, colx):
        values_row = self._cell_values[rowx]
        if colx < len(values_row):
            # Safe for ragged rows: positions past the end of a short
            # row read as empty, without touching the sheet's arrays.
            return self._cell_types[rowx][colx], values_row[colx]
        return XL_CELL_EMPTY, UNICODE_LITERAL('')

    def _xf_index_at(self, rowx, colx):
        if not self._formatting_info:
            return None
        sheet = self._sheet
        if colx < len(self._cell_xf_indexes[rowx]):
            return sheet.cell_xf_index(rowx, colx)
        # Ragged row shorter than colx: no cell-level XF, apply the
        # same row/column/default fallback chain as Sheet.cell_xf_index.
        try:
            xfx = sheet.rowinfo_map[rowx].xf_index
            if xfx > -1:
                return xfx
        except KeyError:
            pass
        try:
            xfx = sheet.colinfo_map[colx].xf_index
            if xfx == -1:
                xfx = 15
            return xfx
        except KeyError:
            return 15

    def _hidden_at(self, rowx, colx):
        if not self._formatting_info:
            return False, False
        sheet = self._sheet
        rowinfo = sheet.rowinfo_map.get(rowx)
        if rowinfo is None:
            row_hidden = bool(sheet.default_row_hidden)
        else:
            row_hidden = bool(rowinfo.hidden)
        colinfo = sheet.colinfo_map.get(colx)
        col_hidden = bool(colinfo.hidden) if colinfo is not None else False
        return row_hidden, col_hidden

    def info(self, rowx, colx):
        """
        Resolve a single grid position.

        :param rowx: Row index, ``0 <= rowx < sheet.nrows``.
        :param colx: Column index, ``0 <= colx < sheet.ncols``.
        :returns: A :class:`MergedCellInfo` whose ``value`` comes from
          the anchor cell of any covering merged range (found via the
          index, even when the anchor lies outside any window being
          iterated), while ``ctype`` and ``xf_index`` describe the
          position itself and ``anchor_xf_index`` the anchor cell.
        """
        anchor = self._anchor_for(rowx, colx)
        ctype, own_value = self._raw_type_value(rowx, colx)
        if anchor is None:
            anchor_rowx, anchor_colx = rowx, colx
            value = own_value
            is_merged = False
        else:
            anchor_rowx, anchor_colx = anchor
            _anchor_type, value = self._raw_type_value(
                anchor_rowx, anchor_colx)
            is_merged = True
        row_hidden, col_hidden = self._hidden_at(rowx, colx)
        return MergedCellInfo(
            rowx, colx, value, ctype,
            anchor_rowx, anchor_colx, is_merged,
            self._xf_index_at(rowx, colx),
            self._xf_index_at(anchor_rowx, anchor_colx),
            row_hidden, col_hidden,
        )

    def iter_window(self, rlo, rhi, clo, chi):
        """
        Iterate row-major over a rectangular query window, yielding a
        :class:`MergedCellInfo` for every position.

        The window is clipped to the sheet's populated bounds
        (``[0, nrows) x [0, ncols)``); negative or oversized edges are
        trimmed rather than raising. Positions covered by a merged
        range resolve to the anchor's value even when the anchor lies
        outside the window. With ``ragged_rows=True``, positions beyond
        a short row's length yield empty raw cells.
        """
        sheet = self._sheet
        rlo = max(rlo, 0)
        rhi = min(rhi, sheet.nrows)
        clo = max(clo, 0)
        chi = min(chi, sheet.ncols)
        for rowx in xrange(rlo, rhi):
            for colx in xrange(clo, chi):
                yield self.info(rowx, colx)

    def window_values(self, rlo, rhi, clo, chi):
        """
        Return the logical values of a query window as a list of rows
        (lists), after the same clipping as :meth:`iter_window`.
        """
        sheet = self._sheet
        rlo = max(rlo, 0)
        rhi = min(rhi, sheet.nrows)
        clo = max(clo, 0)
        chi = min(chi, sheet.ncols)
        return [
            [self.info(rowx, colx).value for colx in xrange(clo, chi)]
            for rowx in xrange(rlo, rhi)
        ]
