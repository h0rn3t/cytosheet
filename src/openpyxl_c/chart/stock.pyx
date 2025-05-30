# cython: language_level=3
"""Stock chart type."""

cdef class XMLNode

cdef class StockChart:
    """Skeleton implementation of StockChart."""

    def __cinit__(self):
        pass

    def add_data(self, series, bint titles_from_data=False) -> None:
        """Add data series to the chart."""
        pass

    def set_categories(self, cats) -> None:
        """Set chart categories."""
        pass

    def to_tree(self) -> XMLNode:
        """Return an XML node representing the chart."""
        return <XMLNode>None
