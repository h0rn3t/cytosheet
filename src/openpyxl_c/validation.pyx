cdef class XMLNode:
    pass

cdef class DataValidation:
    def __cinit__(self, type:str, formula1:str=None, formula2:str=None, allow_blank:bint=False):
        self.type = type
        self.formula1 = formula1
        self.formula2 = formula2
        self.allow_blank = allow_blank
        self.ranges = []

    def add(self, cell_range: str) -> None:
        self.ranges.append(cell_range)

    def to_tree(self) -> XMLNode:
        raise NotImplementedError
