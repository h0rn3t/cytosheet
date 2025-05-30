cdef class XMLNode:
    pass

cdef class ConditionalFormatting:
    def __cinit__(self):
        self.rules = []

    def add_color_scale(self, cfRule) -> None:
        raise NotImplementedError

    def add_data_bar(self, cfRule) -> None:
        raise NotImplementedError

    def add_icon_set(self, cfRule) -> None:
        raise NotImplementedError

    def add_cell_is(self, operator:str, formula:str) -> None:
        raise NotImplementedError

    def add_formula_rule(self, formula:str) -> None:
        raise NotImplementedError

    def to_tree(self) -> XMLNode:
        raise NotImplementedError
