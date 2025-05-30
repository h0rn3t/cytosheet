# cython: language_level=3
"""Conditional formatting rules."""

from __future__ import annotations

cdef class XMLNode

cdef class ConditionalFormatting:
    """Container for conditional formatting rules."""

    def __cinit__(self):
        pass

    def add_color_scale(self, cfRule) -> None:
        """Add a color scale rule."""
        pass

    def add_data_bar(self, cfRule) -> None:
        """Add a data bar rule."""
        pass

    def add_icon_set(self, cfRule) -> None:
        """Add an icon set rule."""
        pass

    def add_cell_is(self, operator:str, formula:str) -> None:
        """Add a cell-is rule."""
        pass

    def add_formula_rule(self, formula:str) -> None:
        """Add a formula rule."""
        pass

    def to_tree(self) -> XMLNode:
        """Serialize rules to XML."""
        return <XMLNode>None
