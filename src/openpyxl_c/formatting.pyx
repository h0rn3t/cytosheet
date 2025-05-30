cdef class XMLNode:
    pass

cdef class ConditionalFormatting:
    def __cinit__(self):
        self.rules = []

    def add_color_scale(self, cfRule) -> None:
        self.rules.append(cfRule)

    def add_data_bar(self, cfRule) -> None:
        self.rules.append(cfRule)

    def add_icon_set(self, cfRule) -> None:
        self.rules.append(cfRule)

    def add_cell_is(self, operator:str, formula:str) -> None:
        self.rules.append((operator, formula))

    def add_formula_rule(self, formula:str) -> None:
        self.rules.append(formula)

    def to_tree(self) -> XMLNode:
        return XMLNode()
