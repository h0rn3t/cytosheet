class Cell:
    """Simple cell object used when Cython extensions are not built."""

    def __init__(self, position=None, value=None):
        self.position = position
        self.value = value

    def __repr__(self):
        return f"<Cell position={self.position}, value={self.value}>"
