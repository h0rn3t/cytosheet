# cython: language_level=3
"""Worksheet comments implementation."""

from __future__ import annotations

cdef class XMLNode

cdef class Comment:
    """Represents a cell comment."""

    def __cinit__(self, str text, str author):
        pass

    def to_tree(self) -> XMLNode:
        """Return an XML node for this comment."""
        return <XMLNode>None
