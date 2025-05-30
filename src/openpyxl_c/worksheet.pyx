from cytosheet.worksheet import Worksheet as CSWorksheet
from .cell import Cell


cdef class Worksheet:
    """Wrapper around :class:`cytosheet.worksheet.Worksheet`."""

    cdef CSWorksheet _ws
    cdef object parent
    cdef str _title

    def __cinit__(self, parent, title: str):
        self.parent = parent
        self._ws = CSWorksheet(parent._wb._shared_strings, title)
        self._title = title

    @classmethod
    cdef Worksheet _wrap(cls, parent, CSWorksheet cs_ws):
        obj = cls.__new__(cls)
        obj.parent = parent
        obj._ws = cs_ws
        obj._title = cs_ws.title
        return obj

    def __getitem__(self, key: str):
        c_cell = self._ws[key]
        col = ord(key[0].upper()) - ord('A') + 1
        row = int(key[1:]) if len(key) > 1 else 0
        return Cell._wrap(self, c_cell, row, col)

    def cell(self, int row, int column, value=None):
        c_cell = self._ws.cell(row, column, value)
        return Cell._wrap(self, c_cell, row, column)

    def iter_rows(self, int min_row=1, int max_row=None, int min_col=1, int max_col=None):
        for row in self._ws.iter_rows(min_row=min_row, max_row=max_row,
                                      min_col=min_col, max_col=max_col):
            yield tuple(
                Cell._wrap(self, c, int(c.position[1:]),
                           ord(c.position[0].upper()) - ord('A') + 1)
                for c in row
            )

    def iter_cols(self, int min_col=1, int max_col=None, int min_row=1, int max_row=None):
        for col in self._ws.iter_cols(min_col=min_col, max_col=max_col,
                                      min_row=min_row, max_row=max_row):
            yield tuple(
                Cell._wrap(self, c, int(c.position[1:]),
                           ord(c.position[0].upper()) - ord('A') + 1)
                for c in col
            )

    def append(self, iterable):
        self._ws.append(iterable)

    def merge_cells(self, range_string: str):
        self._ws.merge_cells(range_string)

    def unmerge_cells(self, range_string: str):
        self._ws.unmerge_cells(range_string)

    def add_chart(self, chart, anchor: str):
        pass

    def add_image(self, img, anchor: str):
        pass

    @property
    def title(self) -> str:
        return self._title

    @title.setter
    def title(self, value: str) -> None:
        self._title = value



