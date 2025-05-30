cdef class Cell:
    # attributes declared in cell.pxd

    def __init__(self, str position=None, object value=None, object formula=None, object style_id=None):
        self.position = position
        self.value = value
        self.formula = formula
        self.style_id = style_id

    def __repr__(self):
        """Return a string representation of the Cell."""
        return f"<Cell position={self.position}, value={self.value}, formula={self.formula}>"
