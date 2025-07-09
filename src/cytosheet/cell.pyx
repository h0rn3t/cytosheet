cdef class Cell:
    cdef str position
    cdef public object value
    cdef public object style
    cdef public object parent
    cdef public bint is_merged_cell
    cdef public str merged_range

    def __init__(self, str position=None, object value=None, object style=None, object parent=None):
        self.position = position if position is not None else ""
        self.value = value
        self.style = style
        self.parent = parent
        self.is_merged_cell = False
        self.merged_range = None

    def __repr__(self):
        if self.is_merged_cell:
            return f"<MergedCell position={self.position}, merged_range={self.merged_range}>"
        return f"<Cell position={self.position}, value={self.value}>"

    cpdef str get_position(self):
        """Быстрый доступ к позиции ячейки"""
        return self.position

    cpdef void set_value(self, object value):
        """Быстрая установка значения"""
        self.value = value

    cpdef object get_value(self):
        """Быстрый доступ к значению"""
        return self.value

    cpdef void set_style(self, object style):
        """Быстрая установка стиля"""
        self.style = style

    cpdef object get_style(self):
        """Быстрый доступ к стилю"""
        return self.style
