# cython: language_level=3
"""Cell module implemented in Cython.

This module defines the :class:`Cell` class with minimal
attribute accessors and method signatures.
"""

from __future__ import annotations

cdef class Font
cdef class PatternFill
cdef class Border
cdef class Alignment
cdef class Comment
cdef class Hyperlink
cdef class Worksheet  # forward declaration

cdef class Cell:
    """Represents a single cell in a worksheet."""

    def __cinit__(self, Worksheet parent, int row, int column, value=None):
        """Create a cell at ``row``/``column`` owned by ``parent``."""
        pass

    property value:
        def __get__(self):
            """Return the cell value."""
            return None
        def __set__(self, v):
            """Set the cell value with type conversion."""
            pass

    property data_type:
        def __get__(self) -> str:
            """Return the cell data type."""
            return ""

    property number_format:
        def __get__(self) -> str:
            """Return the number format string."""
            return ""
        def __set__(self, fmt: str) -> None:
            """Set the number format string."""
            pass

    property font:
        def __get__(self) -> Font:
            """Return the cell font."""
            return <Font>None

    property fill:
        def __get__(self) -> PatternFill:
            """Return the fill settings."""
            return <PatternFill>None

    property border:
        def __get__(self) -> Border:
            """Return the border settings."""
            return <Border>None

    property alignment:
        def __get__(self) -> Alignment:
            """Return the alignment settings."""
            return <Alignment>None

    property comment:
        def __get__(self) -> Comment:
            """Return the cell comment."""
            return <Comment>None
        def __set__(self, Comment c) -> None:
            """Set the cell comment."""
            pass

    property hyperlink:
        def __get__(self) -> Hyperlink:
            """Return the hyperlink object."""
            return <Hyperlink>None

    def offset(self, int row_offset, int col_offset) -> "Cell":
        """Return a cell offset from the current one."""
        return <Cell>None
