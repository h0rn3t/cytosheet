# cython: language_level=3
"""Styles module providing basic formatting classes."""

from __future__ import annotations

cdef class Side

cdef class Font:
    """Font description."""

    def __cinit__(self, name:str='Calibri', double sz=11, bint bold=False, bint italic=False, color:str=None):
        pass

cdef class PatternFill:
    """Cell fill pattern."""

    def __cinit__(self, fill_type:str='none', fgColor:str=None, bgColor:str=None):
        pass

cdef class Border:
    """Border settings."""

    def __cinit__(self, Side left, Side right, Side top, Side bottom, Side diagonal=None):
        pass

cdef class Alignment:
    """Cell alignment options."""

    def __cinit__(self, horizontal:str=None, vertical:str=None, bint wrapText=False):
        pass

cdef class Protection:
    """Cell protection options."""

    def __cinit__(self, bint locked=True, bint hidden=False):
        pass
