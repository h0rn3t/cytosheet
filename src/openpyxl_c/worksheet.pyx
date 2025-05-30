cdef class Worksheet:
    """Placeholder Worksheet object."""

    def __cinit__(self, parent, title: str):
        self.parent = parent
        self._title = title

    def __getitem__(self, key: str):
        raise NotImplementedError

    def cell(self, int row, int column, value=None):
        raise NotImplementedError

    def iter_rows(self, int min_row=1, int max_row=None, int min_col=1, int max_col=None):
        raise NotImplementedError

    def iter_cols(self, int min_col=1, int max_col=None, int min_row=1, int max_row=None):
        raise NotImplementedError

    def append(self, iterable):
        raise NotImplementedError

    def merge_cells(self, range_string: str):
        raise NotImplementedError

    def unmerge_cells(self, range_string: str):
        raise NotImplementedError

    def add_chart(self, chart, anchor: str):
        raise NotImplementedError

    def add_image(self, img, anchor: str):
        raise NotImplementedError

    @property
    def title(self) -> str:
        return self._title

    @title.setter
    def title(self, value: str) -> None:
        self._title = value

    @property
    def sheet_properties(self):
        raise NotImplementedError

    @property
    def page_setup(self):
        raise NotImplementedError

    @property
    def print_options(self):
        raise NotImplementedError
