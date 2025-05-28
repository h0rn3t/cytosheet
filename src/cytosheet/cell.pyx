cdef class Cell:
    # attributes declared in cell.pxd

    def __init__(self, str position=None, object value=None):
        self.position = position
        self.value = value

    def __repr__(self):
        return f"<Cell position={self.position}, value={self.value}>"
