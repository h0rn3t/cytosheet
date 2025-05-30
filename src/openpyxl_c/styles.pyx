cdef class Font:
    def __cinit__(self, name:str='Calibri', sz:float=11, bold: bint=False,
                  italic: bint=False, color:str=None):
        pass

cdef class PatternFill:
    def __cinit__(self, fill_type:str='none', fgColor:str=None, bgColor:str=None):
        pass

cdef class Border:
    def __cinit__(self, left, right, top, bottom, diagonal=None):
        pass

cdef class Alignment:
    def __cinit__(self, horizontal:str=None, vertical:str=None, wrapText:bint=False):
        pass

cdef class Protection:
    def __cinit__(self, locked:bint=True, hidden:bint=False):
        pass
