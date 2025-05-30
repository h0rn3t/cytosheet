cdef class Cell:
    """Minimal Cell representation."""

    def __cinit__(self, parent, int row, int column, value=None):
        self.parent = parent
        self.row = row
        self.column = column
        self._value = value

    @property
    def value(self):
        return self._value

    @value.setter
    def value(self, v):
        self._value = v

    @property
    def data_type(self):
        raise NotImplementedError

    @property
    def number_format(self):
        raise NotImplementedError

    @number_format.setter
    def number_format(self, fmt: str) -> None:
        raise NotImplementedError

    @property
    def font(self):
        raise NotImplementedError

    @property
    def fill(self):
        raise NotImplementedError

    @property
    def border(self):
        raise NotImplementedError

    @property
    def alignment(self):
        raise NotImplementedError

    @property
    def comment(self):
        raise NotImplementedError

    @comment.setter
    def comment(self, c):
        raise NotImplementedError

    @property
    def hyperlink(self):
        raise NotImplementedError

    def offset(self, int row_offset, int col_offset):
        raise NotImplementedError
