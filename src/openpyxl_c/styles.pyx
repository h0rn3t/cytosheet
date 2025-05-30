cdef class Font:
    def __cinit__(self, name:str='Calibri', sz:float=11, bold: bint=False,
                  italic: bint=False, color:str=None):
        self.name = name
        self.sz = sz
        self.bold = bold
        self.italic = italic
        self.color = color

cdef class PatternFill:
    def __cinit__(self, fill_type:str='none', fgColor:str=None, bgColor:str=None):
        self.fill_type = fill_type
        self.fgColor = fgColor
        self.bgColor = bgColor

cdef class Border:
    def __cinit__(self, left, right, top, bottom, diagonal=None):
        self.left = left
        self.right = right
        self.top = top
        self.bottom = bottom
        self.diagonal = diagonal

cdef class Alignment:
    def __cinit__(self, horizontal:str=None, vertical:str=None, wrapText:bint=False):
        self.horizontal = horizontal
        self.vertical = vertical
        self.wrapText = wrapText

cdef class Protection:
    def __cinit__(self, locked:bint=True, hidden:bint=False):
        self.locked = locked
        self.hidden = hidden
