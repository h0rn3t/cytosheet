# cython: language_level=3
"""Workbook module implemented in Cython.

This module provides a minimal skeleton of the :class:`Workbook` class
compatible with the OpenPyXL 2.4.0 API. The implementation only
contains method signatures and basic docstrings.
"""

from __future__ import annotations
from typing import Iterator

cdef class Worksheet  # forward declaration

cdef class Workbook:
    """In-memory representation of an Excel workbook."""

    def __cinit__(self, bint write_only=False, bint guess_types=False):
        """Create a new workbook instance."""
        pass

    @classmethod
    def load_workbook(cls, filename: str, bint read_only=False) -> "Workbook":
        """Load a workbook from ``filename``."""
        return cls()

    def save(self, filename: str, bint as_template=False, bint keep_vba=False) -> None:
        """Save the workbook to ``filename``."""
        pass

    property active:
        def __get__(self) -> Worksheet:
            """Return the active worksheet."""
            return <Worksheet>None

    property sheetnames:
        def __get__(self) -> list:
            """Return a list of worksheet titles."""
            return []

    def create_sheet(self, title: str=None, index: int=None) -> Worksheet:
        """Create a new worksheet."""
        return <Worksheet>None

    def remove(self, sheet: Worksheet) -> None:
        """Remove ``sheet`` from the workbook."""
        pass

    def __getitem__(self, key: str) -> Worksheet:
        """Return worksheet by name."""
        return <Worksheet>None

    def __iter__(self) -> Iterator[Worksheet]:
        """Iterate over worksheets."""
        return iter([])
