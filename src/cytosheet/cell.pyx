cdef class Cell:
    cdef str position
    cdef public object value

    def __init__(self, str position=None, object value=None):
        self.position = position
        self.value = value

    def __repr__(self):
        return f"<Cell position={self.position}, value={self.value}>"
