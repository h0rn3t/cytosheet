# cython: language_level=3
"""Reader module for parsing OOXML workbooks."""

from __future__ import annotations

cdef class Workbook
cdef class Worksheet

cdef class Reader:
    """Parse an Excel file into a :class:`Workbook`."""

    def __cinit__(self, filename: str, bint read_only=False):
        pass

    def _unpack_archive(self) -> None:
        """Unpack the OOXML ZIP archive."""
        pass

    def _parse_workbook(self) -> None:
        """Parse the workbook description."""
        pass

    def _parse_worksheet(self, sheet_path: str) -> Worksheet:
        """Parse a single worksheet."""
        return <Worksheet>None

    def _parse_shared_strings(self) -> None:
        """Parse the shared strings table."""
        pass

    def _parse_styles(self) -> None:
        """Parse workbook styles."""
        pass

    def _parse_charts(self) -> None:
        """Parse chart data."""
        pass

    def _parse_comments(self) -> None:
        """Parse worksheet comments."""
        pass

    def _parse_data_validations(self) -> None:
        """Parse data validation rules."""
        pass

    def _parse_conditional_formatting(self) -> None:
        """Parse conditional formatting rules."""
        pass
