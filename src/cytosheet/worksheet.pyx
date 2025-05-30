import io

from lxml import etree

from .cell import Cell as PyCell
from .cell cimport Cell

NS_MAIN = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

cdef class Worksheet:
    cdef public dict _cells
    cdef public str title
    cdef public list _shared_strings
    cdef public dict _data
    cdef public set _merged_cells


    def __init__(self, list shared_strings=None, str title="Sheet"):
        if shared_strings is None:
            shared_strings = []
        self.title = title
        self._cells = {}
        self._shared_strings = shared_strings
        self._data = {}
        self._merged_cells = set()

    def __getitem__(self, key: str):
        if key in self._cells:
            return self._cells[key]
        else:
            self._cells[key] = PyCell(position=key)
            return self._cells[key]

    def __setitem__(self, cell: str, value):
        # Создаем новую ячейку, если её ещё нет
        if cell not in self._cells:
            self._cells[cell] = PyCell(position=cell)
        # Устанавливаем значение ячейки
        self._cells[cell] = value

    def merge_cells(self, range_string: str):
        self._merged_cells.add(range_string)

    def unmerge_cells(self, range_string: str):
        self._merged_cells.discard(range_string)

    @property
    def merged_cells(self):
        return self._merged_cells

    cpdef cell(self, int row, int column, value=None):
        """Return or create a cell by numeric coordinates (openpyxl compatibility)."""
        cdef str col_letter = chr(ord('A') + column - 1)
        cdef str position = f"{col_letter}{row}"
        cdef object c = self[position]
        if value is not None:
            c.value = value
        return c


    cpdef void _parse_sheet(self, bytes xml_data):
        """Parse worksheet XML using a single ``iterparse`` loop."""
        cdef object context = etree.iterparse(
            io.BytesIO(xml_data),
            events=('end',),
            tag=(NS_MAIN + 'c', NS_MAIN + 'mergeCell'),
        )
        cdef object event
        cdef object elem
        cdef str col_ref
        cdef object value_elem
        cdef object formula_elem
        cdef object value
        cdef object style_id
        cdef str ref
        cdef int shared_string_index
        cdef int ss_len = len(self._shared_strings)
        for event, elem in context:
            if elem.tag == NS_MAIN + 'mergeCell':
                ref = elem.attrib.get('ref')
                if ref:
                    self._merged_cells.add(ref)
            else:
                col_ref = elem.attrib['r']
                cell_type = elem.attrib.get('t')
                value = None
                value_elem = elem.find(NS_MAIN + 'v')
                formula_elem = elem.find(NS_MAIN + 'f')
                style_id = elem.attrib.get('s')
                if cell_type == 's' and value_elem is not None:
                    shared_string_index = int(value_elem.text)
                    if shared_string_index < ss_len:
                        value = self._shared_strings[shared_string_index]
                elif value_elem is not None:
                    value = str(value_elem.text)
                self._cells[col_ref] = PyCell(
                    position=col_ref,
                    value=value,
                    formula=formula_elem.text if formula_elem is not None else None,
                    style_id=style_id,
                )
            elem.clear()

    def get_xml_data(self) -> bytes:
        """
        Generates XML data for the worksheet.
        :return:
        """
        rows_data = {}
        cdef object cell
        cdef str cell_position
        cdef int row
        cdef str column
        cdef str style_attr
        for cell_position, cell in self._cells.items():
            if cell.value is not None or cell.formula is not None:
                column, row = cell_position[0], int(cell_position[1:])
                if row not in rows_data:
                    rows_data[row] = []
                parts = []
                if cell.formula is not None:
                    parts.append(f'<f>{cell.formula}</f>')
                if cell.value is not None:
                    parts.append(f'<v>{cell.value}</v>')
                style_attr = f' s="{cell.style_id}"' if cell.style_id is not None else ''
                rows_data[row].append(
                    f'<c r="{cell_position}"{style_attr} t="str">{"".join(parts)}</c>'
                )

        # Генерируем строки XML с каждой строкой, содержащей свои ячейки
        rows_xml = [
            f'<row r="{row}">{" ".join(cells)}</row>'
            for row, cells in sorted(rows_data.items())
        ]
        merge_xml = ''
        if self._merged_cells:
            merge_elems = [f'<mergeCell ref="{rng}"/>' for rng in sorted(self._merged_cells)]
            merge_xml = f'<mergeCells count="{len(self._merged_cells)}">{" ".join(merge_elems)}</mergeCells>'

        xml_content = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheetData>
                {" ".join(rows_xml)}
            </sheetData>
            {merge_xml}
        </worksheet>"""

        return xml_content.encode('utf-8')

    # ------------------------------------------------------------------
    # Additional openpyxl-like helpers

    cdef str _column_letter(int idx):
        """Convert a 1-based column index to its Excel column letter."""
        cdef list letters = []
        cdef int num = idx
        while num > 0:
            num -= 1
            letters.append(chr(num % 26 + ord('A')))
            num //= 26
        letters.reverse()
        return "".join(letters)

    cdef int _column_index_from_key(str key):
        """Extract 1-based column index from cell key like 'A1'."""
        cdef int i = 0
        cdef int result = 0
        cdef str ch
        for ch in key:
            if ch.isalpha():
                result = result * 26 + (ord(ch.upper()) - ord('A') + 1)
                i += 1
            else:
                break
        return result

    cdef int _row_index_from_key(str key):
        """Extract row index from cell key like 'A1'."""
        cdef int i = 0
        for i in range(len(key)):
            if not key[i].isalpha():
                break
        if i < len(key):
            return int(key[i:])
        return 0

    @property
    def max_row(self):
        """Return the maximum row index that contains data."""
        if not self._cells:
            return 0
        return max(self._row_index_from_key(k) for k in self._cells.keys())

    @property
    def max_column(self):
        """Return the maximum column index that contains data."""
        if not self._cells:
            return 0
        return max(self._column_index_from_key(k) for k in self._cells.keys())

    cpdef append(self, list values):
        """Append a list of values to the next available row."""
        cdef int row_idx = self.max_row + 1
        cdef int col_idx
        cdef str col_letter
        cdef object cell
        for col_idx, value in enumerate(values, start=1):
            col_letter = self._column_letter(col_idx)
            cell = self[f"{col_letter}{row_idx}"]
            cell.value = value

    cpdef iter_rows(self, int min_row=1, int max_row=None, int min_col=1, int max_col=None):
        """Yield rows of cells from the worksheet."""
        if max_row is None:
            max_row = self.max_row
        if max_col is None:
            max_col = self.max_column
        cdef int row
        cdef int col
        for row in range(min_row, max_row + 1):
            cdef list row_cells = []
            for col in range(min_col, max_col + 1):
                row_cells.append(self.cell(row, col))
            yield row_cells

    cpdef iter_cols(self, int min_col=1, int max_col=None, int min_row=1, int max_row=None):
        """Yield columns of cells from the worksheet."""
        if max_col is None:
            max_col = self.max_column
        if max_row is None:
            max_row = self.max_row
        cdef int col
        cdef int row
        for col in range(min_col, max_col + 1):
            cdef list col_cells = []
            for row in range(min_row, max_row + 1):
                col_cells.append(self.cell(row, col))
            yield col_cells
