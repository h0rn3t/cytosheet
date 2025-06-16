cdef class Cell:
    cdef str position
    cdef public object value

    def __init__(self, str position=None, object value=None):
        self.position = position if position is not None else ""
        self.value = value

    def __repr__(self):
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