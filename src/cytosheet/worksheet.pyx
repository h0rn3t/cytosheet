from .cell import Cell
from lxml import etree
import io

cdef class Worksheet:
    cdef public dict _cells
    cdef public str title
    cdef public list _shared_strings

    def __init__(self, str title, list shared_strings):
        self.title = title
        self._cells = {}
        self._shared_strings = shared_strings

    def __getitem__(self, key: str):
        if key in self._cells:
            return self._cells[key]
        else:
            self._cells[key] = Cell()
            return self._cells[key]

    def _parse_sheet(self, bytes xml_data):
        ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        xml_stream = io.BytesIO(xml_data)
        context = etree.iterparse(xml_stream, events=('start', 'end'),
                                  tag='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row')
        for event, row in context:
            if event == 'end':
                row_num = row.attrib['r']
                cells = row.findall('.//main:c', namespaces=ns)

                for cell in cells:
                    col_ref = cell.attrib['r']
                    cell_type = cell.attrib.get('t', None)
                    value = None

                    if cell_type == 's':
                        shared_string_index = int(cell.find('main:v', namespaces=ns).text)
                        if shared_string_index < len(self._shared_strings):
                            value = self._shared_strings[shared_string_index]
                    elif cell.find('main:v', namespaces=ns) is not None:
                        value = str(cell.find('main:v', namespaces=ns).text)

                    self._cells[col_ref] = Cell(value=value)

                row.clear()
