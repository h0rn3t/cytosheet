from zipfile import ZipFile

from .cell import Cell
from lxml import etree
import io


class Worksheet:
    def __init__(
        self,
        shared_strings: list[str],
        title: str = "Sheet",
        archive: ZipFile | None = None,
        _sheet_path: str | None = None,
        _preload: bool = True
    ):

        self.title = title
        self._shared_strings = shared_strings
        self._cells: dict[str, object] = {}

        self._archive = archive
        self._sheet_path = _sheet_path

        if archive and _preload:
            xml = archive.read(_sheet_path)
            self._parse_sheet(xml)

    def iter_rows(self, values_only: bool = True):
        """
        Потоковое чтение строк. Работает только для листов,
        открытых из файла (self._archive != None).
        """
        if not self._archive:
            raise RuntimeError("iter_rows доступен только в режиме lazy")

        ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
        row_tag = f"{ns}row"
        c_tag   = f"{ns}c"
        v_tag   = f"{ns}v"

        with self._archive.open(self._sheet_path) as fh:
            for _event, row in etree.iterparse(fh, tag=row_tag, events=("end",)):
                out = []
                for c in row.iter(c_tag):
                    t = c.get("t")          # тип ячейки
                    v = c.find(v_tag)
                    if v is None:
                        continue
                    raw = v.text
                    if t == "s":            # shared-string
                        idx = int(raw)
                        if 0 <= idx < len(self._shared_strings):
                            value = self._shared_strings[idx]
                        else:  # резервный вариант — вернуть исходный индекс
                            value = raw
                    else:                   # число / строка
                        try:
                            value = int(raw)
                        except ValueError:
                            try:
                                value = float(raw)
                            except ValueError:
                                value = raw
                    out.append(value if values_only else raw)
                yield out
                row.clear()    # важное освобождение памяти


# cdef class Worksheet:
#     cdef public dict _cells
#     cdef public str title
#     cdef public list _shared_strings
#     cdef public dict _data
#     cdef public object _archive  # ZipFile | None
#     cdef public str    _sheet_path
#
#
#     def __init__(self, list shared_strings=None, str title="Sheet"):
#         if shared_strings is None:
#             shared_strings = []
#         self.title = title
#         self._cells = {}
#         self._shared_strings = shared_strings
#         self._data = {}

    def __getitem__(self, key: str):
        if key in self._cells:
            return self._cells[key]
        else:
            self._cells[key] = Cell(position=key)
            return self._cells[key]

    # def __setitem__(self, cell: str, value):
    #     # Создаем новую ячейку, если её ещё нет
    #     if cell not in self._cells:
    #         self._cells[cell] = Cell(position=cell)
    #     # Устанавливаем значение ячейки
    #     self._cells[cell] = value

    def __setitem__(self, key: str, value):
        if key not in self._cells:
            self._cells[key] = Cell(position=key, value=value)
        else:
            self._cells[key].value = value

    def _parse_sheet(self, bytes xml_data):
        ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        xml_stream = io.BytesIO(xml_data)
        context = etree.iterparse(
            xml_stream,
            events=('start', 'end'),
            tag='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row'
        )
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
                    # elif cell.find('main:v', namespaces=ns) is not None:
                    #     value = str(cell.find('main:v', namespaces=ns).text)
                    elif cell.find('main:v', namespaces=ns) is not None:
                        raw = cell.find('main:v', namespaces=ns).text
                        try:
                            value = int(raw)
                        except ValueError:
                            try:
                                value = float(raw)
                            except ValueError:
                                value = raw

                    self._cells[col_ref] = Cell(value=value)

                row.clear()

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
                # rows_data[row].append(
                #     f'<c r="{cell_position}" t="str"><v>{cell.value}</v></c>'
                # )
                # Определяем тип данных
                if isinstance(cell.value, (int, float)):
                    tag = f'<c r="{cell_position}" t="n"><v>{cell.value}</v></c>'
                else:
                    tag = f'<c r="{cell_position}" t="str"><v>{cell.value}</v></c>'
                rows_data[row].append(tag)

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
