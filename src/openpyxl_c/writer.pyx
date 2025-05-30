# cython: language_level=3
"""Writer module for producing OOXML workbooks."""

from __future__ import annotations

cdef class Workbook
cdef class Worksheet

cdef class Writer:
    """Serialize a :class:`Workbook` to disk."""

    def __cinit__(self, Workbook workbook, str filename):
        pass

    def write(self) -> None:
        """Write the workbook as a ZIP archive."""
        pass

    def _write_workbook(self) -> bytes:
        """Serialize workbook metadata."""
        return b""

    def _write_worksheet(self, Worksheet ws) -> bytes:
        """Serialize a worksheet."""
        return b""

    def _write_shared_strings(self) -> bytes:
        """Serialize shared strings table."""
        return b""

    def _write_styles(self) -> bytes:
        """Serialize styles."""
        return b""

    def _write_charts(self) -> bytes:
        """Serialize chart objects."""
        return b""

    def _write_comments(self) -> bytes:
        """Serialize comments."""
        return b""

    def _write_data_validations(self) -> bytes:
        """Serialize data validation rules."""
        return b""

    def _write_conditional_formatting(self) -> bytes:
        """Serialize conditional formatting rules."""
        return b""
