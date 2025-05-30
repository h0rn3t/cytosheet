# cython: language_level=3
"""Worksheet module implemented in Cython.

This module defines the :class:`Worksheet` class with method
signatures matching OpenPyXL 2.4.0.
"""

from __future__ import annotations
from typing import Iterable, Iterator

cdef class Workbook  # forward declaration
cdef class Cell
cdef class Chart
cdef class Image
cdef class SheetProperties
cdef class PageSetup
cdef class PrintOptions

cdef class Worksheet:
    """Representation of a single worksheet."""

    def __cinit__(self, Workbook parent, str title):
        """Create a worksheet bound to ``parent``."""
        pass

    def __getitem__(self, key: str) -> Cell:
        """Return cell or range by key."""
        return <Cell>None

    def cell(self, int row, int column, value=None) -> Cell:
        """Return or create a cell by numeric coordinates."""
        return <Cell>None

    def iter_rows(self, int min_row=1, int max_row=None, int min_col=1, int max_col=None) -> Iterator:
        """Iterate over rows of cells."""
        return iter([])

    def iter_cols(self, int min_col=1, int max_col=None, int min_row=1, int max_row=None) -> Iterator:
        """Iterate over columns of cells."""
        return iter([])

    def append(self, Iterable iterable) -> None:
        """Append a row of values."""
        pass

    def merge_cells(self, range_string: str) -> None:
        """Merge the specified cell range."""
        pass

    def unmerge_cells(self, range_string: str) -> None:
        """Unmerge the specified cell range."""
        pass

    def add_chart(self, Chart chart, anchor: str) -> None:
        """Add a chart at ``anchor``."""
        pass

    def add_image(self, Image img, anchor: str) -> None:
        """Add an image at ``anchor``."""
        pass

    property title:
        def __get__(self) -> str:
            """Get worksheet title."""
            return ""
        def __set__(self, value: str) -> None:
            """Set worksheet title."""
            pass

    property sheet_properties:
        def __get__(self) -> SheetProperties:
            """Return sheet properties."""
            return <SheetProperties>None

    property page_setup:
        def __get__(self) -> PageSetup:
            """Return page setup object."""
            return <PageSetup>None

    property print_options:
        def __get__(self) -> PrintOptions:
            """Return print options."""
            return <PrintOptions>None
