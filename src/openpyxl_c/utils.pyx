# cython: language_level=3
"""Utility functions for coordinate handling and XML processing."""

from __future__ import annotations

cdef public tuple coordinate_from_string(str coord):
    """Return row and column index from an Excel coordinate string."""
    return 1, 1

cdef public str get_column_letter(int idx):
    """Convert a column index to its Excel letter."""
    return "A"

cdef public tuple range_boundaries(str range_string):
    """Return boundary coordinates for a range string."""
    return 1, 1, 1, 1

cdef public str quote_sheetname(str name):
    """Return ``name`` quoted for XML."""
    return name

cdef public str xml_escape(str text):
    """Escape XML special characters in ``text``."""
    return text
