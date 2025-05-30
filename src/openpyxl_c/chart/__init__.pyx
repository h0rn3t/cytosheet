# cython: language_level=3

cdef class XMLNode:
    pass

cdef class AreaChart:
    def __cinit__(self):
        self.series = []
        self.cats = None

    def add_data(self, series, bint titles_from_data=False) -> None:
        self.series.append(series)

    def set_categories(self, cats) -> None:
        self.cats = cats

    def to_tree(self) -> XMLNode:
        return XMLNode()

cdef class BarChart:
    def __cinit__(self):
        self.series = []
        self.cats = None

    def add_data(self, series, bint titles_from_data=False) -> None:
        self.series.append(series)

    def set_categories(self, cats) -> None:
        self.cats = cats

    def to_tree(self) -> XMLNode:
        return XMLNode()

cdef class LineChart:
    def __cinit__(self):
        self.series = []
        self.cats = None

    def add_data(self, series, bint titles_from_data=False) -> None:
        self.series.append(series)

    def set_categories(self, cats) -> None:
        self.cats = cats

    def to_tree(self) -> XMLNode:
        return XMLNode()

cdef class PieChart:
    def __cinit__(self):
        self.series = []
        self.cats = None

    def add_data(self, series, bint titles_from_data=False) -> None:
        self.series.append(series)

    def set_categories(self, cats) -> None:
        self.cats = cats

    def to_tree(self) -> XMLNode:
        return XMLNode()

cdef class DoughnutChart:
    def __cinit__(self):
        self.series = []
        self.cats = None

    def add_data(self, series, bint titles_from_data=False) -> None:
        self.series.append(series)

    def set_categories(self, cats) -> None:
        self.cats = cats

    def to_tree(self) -> XMLNode:
        return XMLNode()

cdef class ScatterChart:
    def __cinit__(self):
        self.series = []
        self.cats = None

    def add_data(self, series, bint titles_from_data=False) -> None:
        self.series.append(series)

    def set_categories(self, cats) -> None:
        self.cats = cats

    def to_tree(self) -> XMLNode:
        return XMLNode()

cdef class BubbleChart:
    def __cinit__(self):
        self.series = []
        self.cats = None

    def add_data(self, series, bint titles_from_data=False) -> None:
        self.series.append(series)

    def set_categories(self, cats) -> None:
        self.cats = cats

    def to_tree(self) -> XMLNode:
        return XMLNode()

cdef class RadarChart:
    def __cinit__(self):
        self.series = []
        self.cats = None

    def add_data(self, series, bint titles_from_data=False) -> None:
        self.series.append(series)

    def set_categories(self, cats) -> None:
        self.cats = cats

    def to_tree(self) -> XMLNode:
        return XMLNode()

cdef class StockChart:
    def __cinit__(self):
        self.series = []
        self.cats = None

    def add_data(self, series, bint titles_from_data=False) -> None:
        self.series.append(series)

    def set_categories(self, cats) -> None:
        self.cats = cats

    def to_tree(self) -> XMLNode:
        return XMLNode()

cdef class SurfaceChart:
    def __cinit__(self):
        self.series = []
        self.cats = None

    def add_data(self, series, bint titles_from_data=False) -> None:
        self.series.append(series)

    def set_categories(self, cats) -> None:
        self.cats = cats

    def to_tree(self) -> XMLNode:
        return XMLNode()
