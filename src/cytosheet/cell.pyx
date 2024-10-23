cdef class Cell:
    cdef public object value

    def __init__(self, object value=None):
        self.value = value

    def __repr__(self):
        return f"<Cell value={self.value}>"
