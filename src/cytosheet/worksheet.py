import io
from lxml import etree

from .cell import Cell

NS_MAIN = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'


class Worksheet:
    def __init__(self, shared_strings=None, title="Sheet"):
        if shared_strings is None:
            shared_strings = []
        self.title = title
        self._cells = {}
        self._shared_strings = shared_strings
        self._data = {}

    def __getitem__(self, key: str):
        if key not in self._cells:
            self._cells[key] = Cell(position=key)
        return self._cells[key]

    def __setitem__(self, cell: str, value):
        self[cell].value = value

    def cell(self, row: int, column: int, value=None):
        col_letter = chr(ord('A') + column - 1)
        position = f"{col_letter}{row}"
        c = self[position]
        if value is not None:
            c.value = value
        return c

    def _parse_sheet(self, xml_data: bytes):
        xml_stream = io.BytesIO(xml_data)
        context = etree.iterparse(xml_stream, events=('end',), tag=NS_MAIN + 'c')
        for _, cell in context:
            col_ref = cell.attrib['r']
            cell_type = cell.attrib.get('t')
            value_elem = cell.find(NS_MAIN + 'v')
            value = None
            if cell_type == 's' and value_elem is not None:
                idx = int(value_elem.text)
                if idx < len(self._shared_strings):
                    value = self._shared_strings[idx]
            elif value_elem is not None:
                value = str(value_elem.text)
            self._cells[col_ref] = Cell(position=col_ref, value=value)
            cell.clear()

    def get_xml_data(self) -> bytes:
        rows_data = {}
        for cell_position, cell in self._cells.items():
            if cell.value is not None:
                column, row = cell_position[0], int(cell_position[1:])
                rows_data.setdefault(row, []).append(
                    f'<c r="{cell_position}" t="str"><v>{cell.value}</v></c>'
                )
        rows_xml = [
            f'<row r="{row}">{" ".join(cells)}</row>'
            for row, cells in sorted(rows_data.items())
        ]
        xml_content = (
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
            "<sheetData>"
            f"{' '.join(rows_xml)}"
            "</sheetData></worksheet>"
        )
        return xml_content.encode('utf-8')
