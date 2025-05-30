# cython: language_level=3

cdef class XMLNode:
    pass

cdef class AreaChart:
    def __cinit__(self):
        pass

    def add_data(self, series, bint titles_from_data=False) -> None:
        raise NotImplementedError

    def set_categories(self, cats) -> None:
        raise NotImplementedError

    def to_tree(self) -> XMLNode:
        raise NotImplementedError

cdef class BarChart:
    def __cinit__(self):
        pass

    def add_data(self, series, bint titles_from_data=False) -> None:
        raise NotImplementedError

    def set_categories(self, cats) -> None:
        raise NotImplementedError

    def to_tree(self) -> XMLNode:
        raise NotImplementedError

cdef class LineChart:
    def __cinit__(self):
        pass

    def add_data(self, series, bint titles_from_data=False) -> None:
        raise NotImplementedError

    def set_categories(self, cats) -> None:
        raise NotImplementedError

    def to_tree(self) -> XMLNode:
        raise NotImplementedError

cdef class PieChart:
    def __cinit__(self):
        pass

    def add_data(self, series, bint titles_from_data=False) -> None:
        raise NotImplementedError

    def set_categories(self, cats) -> None:
        raise NotImplementedError

    def to_tree(self) -> XMLNode:
        raise NotImplementedError

cdef class DoughnutChart:
    def __cinit__(self):
        pass

    def add_data(self, series, bint titles_from_data=False) -> None:
        raise NotImplementedError

    def set_categories(self, cats) -> None:
        raise NotImplementedError

    def to_tree(self) -> XMLNode:
        raise NotImplementedError

cdef class ScatterChart:
    def __cinit__(self):
        pass

    def add_data(self, series, bint titles_from_data=False) -> None:
        raise NotImplementedError

    def set_categories(self, cats) -> None:
        raise NotImplementedError

    def to_tree(self) -> XMLNode:
        raise NotImplementedError

cdef class BubbleChart:
    def __cinit__(self):
        pass

    def add_data(self, series, bint titles_from_data=False) -> None:
        raise NotImplementedError

    def set_categories(self, cats) -> None:
        raise NotImplementedError

    def to_tree(self) -> XMLNode:
        raise NotImplementedError

cdef class RadarChart:
    def __cinit__(self):
        pass

    def add_data(self, series, bint titles_from_data=False) -> None:
        raise NotImplementedError

    def set_categories(self, cats) -> None:
        raise NotImplementedError

    def to_tree(self) -> XMLNode:
        raise NotImplementedError

cdef class StockChart:
    def __cinit__(self):
        pass

    def add_data(self, series, bint titles_from_data=False) -> None:
        raise NotImplementedError

    def set_categories(self, cats) -> None:
        raise NotImplementedError

    def to_tree(self) -> XMLNode:
        raise NotImplementedError

cdef class SurfaceChart:
    def __cinit__(self):
        pass

    def add_data(self, series, bint titles_from_data=False) -> None:
        raise NotImplementedError

    def set_categories(self, cats) -> None:
        raise NotImplementedError

    def to_tree(self) -> XMLNode:
        raise NotImplementedError
