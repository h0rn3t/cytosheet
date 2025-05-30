# cython: language_level=3
"""Data validation rules."""

from __future__ import annotations

cdef class XMLNode

cdef class DataValidation:
    """Represents data validation settings."""

    def __cinit__(self, type:str, formula1:str=None, formula2:str=None, bint allow_blank=False):
        pass

    def add(self, cell_range: str) -> None:
        """Add a cell range to the validation."""
        pass

    def to_tree(self) -> XMLNode:
        """Serialize validation to XML."""
        return <XMLNode>None
