from cytosheet.cell import Cell as CSCell


cdef class Cell:
    """Wrapper around :class:`cytosheet.cell.Cell`."""

    cdef CSCell _cell
    cdef object parent
    cdef int row
    cdef int column

    def __cinit__(self, parent, int row, int column, value=None):
        self.parent = parent
        self.row = row
        self.column = column
        self._cell = parent._ws.cell(row, column, value)

    @classmethod
    cdef Cell _wrap(cls, parent, CSCell cs, int row, int column):
        obj = cls.__new__(cls)
        obj.parent = parent
        obj.row = row
        obj.column = column
        obj._cell = cs
        return obj

    @property
    def value(self):
        return self._cell.value

    @value.setter
    def value(self, v):
        self._cell.value = v

    @property
    def data_type(self):
        if isinstance(self._cell.value, str):
            return "s"
        elif isinstance(self._cell.value, bool):
            return "b"
        elif self._cell.value is None:
            return "n"
        else:
            return "n"

    @property
    def number_format(self):
        return getattr(self._cell, "number_format", "General")

    @number_format.setter
    def number_format(self, fmt: str) -> None:
        setattr(self._cell, "number_format", fmt)

    @property
    def font(self):
        return getattr(self._cell, "font", None)

    @property
    def fill(self):
        return getattr(self._cell, "fill", None)

    @property
    def border(self):
        return getattr(self._cell, "border", None)

    @property
    def alignment(self):
        return getattr(self._cell, "alignment", None)

    @property
    def comment(self):
        return getattr(self._cell, "comment", None)

    @comment.setter
    def comment(self, c):
        self._cell.comment = c

    @property
    def hyperlink(self):
        return getattr(self._cell, "hyperlink", None)

    def offset(self, int row_offset, int col_offset):
        return self.parent.cell(self.row + row_offset, self.column + col_offset)

