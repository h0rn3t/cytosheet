import io

from lxml import etree

from .cell import Cell

NS_MAIN = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

cdef class Worksheet:
    cdef public dict _cells
    cdef public str title
    cdef public list _shared_strings
    cdef public dict _data


    def __init__(self, list shared_strings=None, str title="Sheet"):
        if shared_strings is None:
            shared_strings = []
        self.title = title
        self._cells = {}
        self._shared_strings = shared_strings
        self._data = {}

    def __getitem__(self, key: str):
        if key in self._cells:
            return self._cells[key]
        else:
            self._cells[key] = Cell(position=key)
            return self._cells[key]

    def __setitem__(self, cell: str, value):
        # Создаем новую ячейку, если её ещё нет
        if cell not in self._cells:
            self._cells[cell] = Cell(position=cell)
        # Устанавливаем значение ячейки
        self._cells[cell] = value

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
        xml_stream = io.BytesIO(xml_data)
        cdef object context = etree.iterparse(xml_stream, events=('end',), tag=NS_MAIN + 'c')
        cdef object event
        cdef object cell
        cdef str col_ref
        cdef object value_elem
        cdef object value
        cdef int shared_string_index
        cdef int ss_len = len(self._shared_strings)
        for event, cell in context:
            col_ref = cell.attrib['r']
            cell_type = cell.attrib.get('t')
            value = None
            value_elem = cell.find(NS_MAIN + 'v')
            if cell_type == 's' and value_elem is not None:
                shared_string_index = int(value_elem.text)
                if shared_string_index < ss_len:
                    value = self._shared_strings[shared_string_index]
            elif value_elem is not None:
                value = str(value_elem.text)
            self._cells[col_ref] = Cell(position=col_ref, value=value)
            cell.clear()

    def get_xml_data(self) -> bytes:
        """
        Generates XML data for the worksheet.
        :return:
        """
        rows_data = {}
        for cell_position, cell in self._cells.items():
            if cell.value is not None:
                # Разделим позицию на букву колонки и номер строки
                column, row = cell_position[0], int(cell_position[1:])
                if row not in rows_data:
                    rows_data[row] = []
                # Добавляем ячейку с указанием типа данных (строка)
                rows_data[row].append(
                    f'<c r="{cell_position}" t="str"><v>{cell.value}</v></c>'
                )

        # Генерируем строки XML с каждой строкой, содержащей свои ячейки
        rows_xml = [
            f'<row r="{row}">{" ".join(cells)}</row>'
            for row, cells in sorted(rows_data.items())
        ]
        xml_content = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheetData>
                {" ".join(rows_xml)}
            </sheetData>
        </worksheet>"""

        return xml_content.encode('utf-8')
