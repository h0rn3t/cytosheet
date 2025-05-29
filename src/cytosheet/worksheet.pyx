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
        # Створюємо нову комірку, якщо її ще немає
        if cell not in self._cells:
            self._cells[cell] = PyCell(position=cell)
        # Встановлюємо значення комірки
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
        """Parse worksheet XML using iterparse for better performance."""
        cdef object root
        cdef object context
        cdef object event
        cdef object cell
        cdef str col_ref
        cdef object value_elem
        cdef object formula_elem
        cdef object value
        cdef object style_id
        cdef int shared_string_index
        cdef int ss_len

        xml_stream = io.BytesIO(xml_data)

        root = etree.fromstring(xml_data)
        for m in root.findall('.//' + NS_MAIN + 'mergeCell'):
            ref = m.attrib.get('ref')
            if ref:
                self._merged_cells.add(ref)

        context = etree.iterparse(io.BytesIO(xml_data), events=('end',), tag=NS_MAIN + 'c')
        ss_len = len(self._shared_strings)
        for event, cell in context:
            col_ref = cell.attrib['r']
            cell_type = cell.attrib.get('t')
            value = None
            value_elem = cell.find(NS_MAIN + 'v')
            formula_elem = cell.find(NS_MAIN + 'f')
            style_id = cell.attrib.get('s')
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
            cell.clear()

    def get_xml_data(self) -> bytes:
        """
        Generates XML data for the worksheet.
        :return:
        """
        cdef object cell
        cdef str cell_position
        cdef int row
        cdef str column
        cdef list parts
        cdef str style_attr
        rows_data = {}

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

        # Генеруємо рядки XML з кожним рядком, що містить свої комірки
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