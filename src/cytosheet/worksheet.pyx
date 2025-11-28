from lxml import etree
import io

from .cell import Cell


cdef class Worksheet:
    cdef public dict _cells
    cdef public int _max_row
    cdef public int _max_column
    cdef public str title
    cdef public list _shared_strings
    cdef public object _archive  # ZipFile | None
    cdef public str _sheet_path
    cdef public bint _preloaded
    cdef public bint _is_small_file
    cdef public list _merged_cells  # List[str]

    def __init__(
            self,
            list shared_strings,
            str title = "Sheet",
            object archive = None,
            str _sheet_path = None,
            bint _preload = True
    ):
        self.title = title
        self._shared_strings = shared_strings if shared_strings is not None else []
        self._cells = {}
        self._archive = archive
        self._sheet_path = _sheet_path
        self._preloaded = False
        self._is_small_file = False
        self._merged_cells = []
        self._max_row = 0
        self._max_column = 0

        if archive is not None and _preload and _sheet_path is not None:
            xml = archive.read(_sheet_path)
            xml_size = len(xml)
            self._is_small_file = xml_size < 50000  # ~500 строк
            if self._is_small_file:
                self._parse_sheet_simple(xml)
            else:
                self._parse_sheet(xml)

            # Після повного парсингу один раз обчислюємо max_row/max_column
            self._recalculate_bounds()
            self._preloaded = True

    cdef void _update_bounds_for_cell(self, str coord):
        """Оновити _max_row/_max_column для однієї координати типу 'A1'."""
        cdef int i = 0
        cdef int row, col_num
        while i < len(coord) and coord[i].isalpha():
            i += 1
        if i == 0 or i == len(coord):
            return
        row = int(coord[i:])
        if row > self._max_row:
            self._max_row = row
        col_num = self._col_to_num(coord[:i])
        if col_num > self._max_column:
            self._max_column = col_num

    cdef void _recalculate_bounds(self):
        """Повне переобчислення max_row/max_column з _cells."""
        cdef str coord
        cdef int i, row, col_num
        self._max_row = 0
        self._max_column = 0
        for coord in self._cells.keys():
            i = 0
            while i < len(coord) and coord[i].isalpha():
                i += 1
            if i == 0 or i == len(coord):
                continue
            row = int(coord[i:])
            if row > self._max_row:
                self._max_row = row
            col_num = self._col_to_num(coord[:i])
            if col_num > self._max_column:
                self._max_column = col_num

    # ------------------------------------------------------------------
    # Парсинг листа з XML
    # ------------------------------------------------------------------

    cdef void _parse_sheet(self, bytes xml_data):
        """Выбор между стандартным и чанковым парсером для средних/больших файлов."""
        cdef int xml_size = len(xml_data)
        cdef bint use_chunked_parsing = xml_size > 100000
        if use_chunked_parsing:
            self._parse_sheet_chunked(xml_data)
        else:
            self._parse_sheet_standard(xml_data)

    cdef void _parse_sheet_simple(self, bytes xml_data):
        """Упрощённый string-based парсер для малых файлов с поддержкой формул.<f>."""
        cdef str xml_str = xml_data.decode('utf-8')
        cdef int start_pos = 0
        cdef int cell_start, cell_end
        cdef int r_start, r_end, t_start, t_end, v_start, v_end, f_start, f_end
        cdef str cell_ref, cell_type, raw_value, formula_text
        cdef object value
        cdef int shared_index

        while True:
            cell_start = xml_str.find('<c ', start_pos)
            if cell_start == -1:
                break
            cell_end = xml_str.find('</c>', cell_start)
            if cell_end == -1:
                break

            # r="A1"
            r_start = xml_str.find('r="', cell_start, cell_end)
            if r_start == -1:
                start_pos = cell_end + 4
                continue
            r_start += 3
            r_end = xml_str.find('"', r_start, cell_end)
            if r_end == -1:
                start_pos = cell_end + 4
                continue
            cell_ref = xml_str[r_start:r_end]

            # t="s" / t="str" / t="n" / ...
            cell_type = None
            t_start = xml_str.find('t="', cell_start, cell_end)
            if t_start != -1:
                t_start += 3
                t_end = xml_str.find('"', t_start, cell_end)
                if t_end != -1:
                    cell_type = xml_str[t_start:t_end]

            # Формула имеет приоритет над <v>
            f_start = xml_str.find('<f>', cell_start, cell_end)
            if f_start != -1:
                f_start += 3
                f_end = xml_str.find('</f>', f_start, cell_end)
                if f_end != -1:
                    formula_text = xml_str[f_start:f_end]
                    value = '=' + formula_text
                    cell = Cell(position=cell_ref, value=value)
                    cell.data_type = 'f'
                    self._cells[cell_ref] = cell
                    start_pos = cell_end + 4
                    continue

            v_start = xml_str.find('<v>', cell_start, cell_end)
            if v_start == -1:
                start_pos = cell_end + 4
                continue
            v_start += 3
            v_end = xml_str.find('</v>', v_start, cell_end)
            if v_end == -1:
                start_pos = cell_end + 4
                continue
            raw_value = xml_str[v_start:v_end]

            if cell_type == 's':
                try:
                    shared_index = int(raw_value)
                    if 0 <= shared_index < len(self._shared_strings):
                        value = self._shared_strings[shared_index]
                    else:
                        value = raw_value
                except (ValueError, TypeError):
                    value = raw_value
            else:
                value = self._convert_cell_value_fast(raw_value)

            if value is not None:
                self._cells[cell_ref] = Cell(position=cell_ref, value=value)

            start_pos = cell_end + 4

    cdef void _parse_sheet_standard(self, bytes xml_data):
        """Стандартный iterparse-парсер для средних файлов с поддержкой <f>."""
        cdef dict ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        cdef object xml_stream = io.BytesIO(xml_data)
        cdef object context = etree.iterparse(
            xml_stream,
            events=('end',),
            tag='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'
        )

        cdef str col_ref, cell_type, raw
        cdef int shared_string_index
        cdef object value, v_elem, f_elem, cell
        cdef dict temp_cells = {}

        try:
            for event, cell_elem in context:
                col_ref = cell_elem.get('r')
                if col_ref is None:
                    cell_elem.clear()
                    continue

                # Формула
                f_elem = cell_elem.find('main:f', namespaces=ns)
                if f_elem is not None and f_elem.text is not None:
                    value = '=' + f_elem.text
                    cell = Cell(position=col_ref, value=value)
                    cell.data_type = 'f'
                    temp_cells[col_ref] = cell
                    cell_elem.clear()
                    continue

                cell_type = cell_elem.get('t')
                v_elem = cell_elem.find('main:v', namespaces=ns)
                if v_elem is None:
                    cell_elem.clear()
                    continue

                raw = v_elem.text
                if raw is None:
                    cell_elem.clear()
                    continue

                if cell_type == 's':
                    try:
                        shared_string_index = int(raw)
                        if 0 <= shared_string_index < len(self._shared_strings):
                            value = self._shared_strings[shared_string_index]
                        else:
                            value = raw
                    except (ValueError, TypeError):
                        value = raw
                else:
                    value = self._convert_cell_value_fast(raw)

                if value is not None:
                    temp_cells[col_ref] = Cell(position=col_ref, value=value)

                cell_elem.clear()
        except Exception as e:
            print(f"Ошибка стандартного парсинга: {e}")
        finally:
            self._cells.update(temp_cells)

    cdef void _parse_sheet_chunked(self, bytes xml_data):
        """Чанковый парсер для больших файлов с поддержкой <f>."""
        cdef dict ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        cdef object xml_stream = io.BytesIO(xml_data)
        cdef object context = etree.iterparse(
            xml_stream,
            events=('end',),
            tag='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c',
            recover=True
        )

        cdef str col_ref, cell_type, raw
        cdef int shared_string_index, batch_size, processed_count
        cdef object value, v_elem, f_elem, cell
        cdef dict temp_batch = {}
        batch_size = 1000
        processed_count = 0

        try:
            for event, cell_elem in context:
                col_ref = cell_elem.get('r')
                if col_ref is None:
                    cell_elem.clear()
                    continue

                # Формула
                f_elem = cell_elem.find('main:f', namespaces=ns)
                if f_elem is not None and f_elem.text is not None:
                    value = '=' + f_elem.text
                    cell = Cell(position=col_ref, value=value)
                    cell.data_type = 'f'
                    temp_batch[col_ref] = cell
                    processed_count += 1
                else:
                    cell_type = cell_elem.get('t')
                    v_elem = cell_elem.find('main:v', namespaces=ns)
                    if v_elem is None:
                        cell_elem.clear()
                        continue

                    raw = v_elem.text
                    if raw is None:
                        cell_elem.clear()
                        continue

                    if cell_type == 's':
                        try:
                            shared_string_index = int(raw)
                            if 0 <= shared_string_index < len(self._shared_strings):
                                value = self._shared_strings[shared_string_index]
                            else:
                                value = raw
                        except (ValueError, TypeError):
                            value = raw
                    else:
                        value = self._convert_cell_value_fast(raw)

                    if value is not None:
                        temp_batch[col_ref] = Cell(position=col_ref, value=value)
                        processed_count += 1

                if processed_count >= batch_size:
                    self._cells.update(temp_batch)
                    temp_batch.clear()
                    processed_count = 0

                cell_elem.clear()
        except Exception as e:
            print(f"Ошибка чанкового парсинга: {e}")
        finally:
            if temp_batch:
                self._cells.update(temp_batch)

    # ------------------------------------------------------------------
    # Доступ і запис ячеек
    # ------------------------------------------------------------------

    def __getitem__(self, str key):
        cdef object cell
        if key in self._cells:
            return self._cells[key]
        cell = Cell(position=key, parent=self)
        self._cells[key] = cell
        self._update_bounds_for_cell(key)
        return cell

    def __setitem__(self, str key, object value):
        cdef object cell
        if key not in self._cells:
            cell = Cell(position=key, value=value, parent=self)
            # простая эвристика для формул
            if isinstance(value, str) and value.startswith('='):
                cell.data_type = 'f'
            self._cells[key] = cell
        else:
            cell = self._cells[key]
            cell.value = value
            if isinstance(value, str) and value.startswith('='):
                cell.data_type = 'f'

        self._update_bounds_for_cell(key)

    # ------------------------------------------------------------------
    # openpyxl-совместный доступ к ячейкам по row/column + append/max_*.
    # ------------------------------------------------------------------

    def cell(self, int row, int column, object value=None):
        """Доступ до ячейки по координатах (1-based), як в openpyxl.

        Якщо value не None – одразу встановлюємо його.
        """
        if row < 1 or column < 1:
            raise ValueError("row and column must be >= 1")

        cdef str col_letters = self._num_to_col(column)
        cdef str coord = f"{col_letters}{row}"
        cdef object cell = self._cells.get(coord)

        if cell is None:
            cell = Cell(position=coord, parent=self)
            self._cells[coord] = cell

        if value is not None:
            cell.value = value
            if isinstance(value, str) and value.startswith('='):
                cell.data_type = 'f'

        # Оновлюємо границі
        if row > self._max_row:
            self._max_row = row
        if column > self._max_column:
            self._max_column = column

        return cell

    @property
    def max_row(self):
        """Номер останнього рядка з даними (як у openpyxl)."""
        if self._max_row == 0 and self._cells:
            self._recalculate_bounds()
        return self._max_row

    @property
    def max_column(self):
        """Номер останньої колонки з даними (як у openpyxl)."""
        if self._max_column == 0 and self._cells:
            self._recalculate_bounds()
        return self._max_column

    def append(self, object iterable):
        """Додати рядок значень в кінець аркуша (openpyxl-сумісний append)."""
        if iterable is None:
            return

        cdef int row = self.max_row + 1
        cdef int col

        # Строки и bytes считаем скалярами, як в openpyxl: кладём целиком в первый столбец.
        if isinstance(iterable, (str, bytes)):
            self.cell(row=row, column=1, value=iterable)
            return

        # Підтримка словників поки не потрібна для наших сценаріїв,
        # тому працюємо з послідовностями.
        try:
            for col, value in enumerate(iterable, 1):
                self.cell(row=row, column=col, value=value)
        except TypeError:
            # Неітерований об'єкт – просто кладемо його в першу колонку
            self.cell(row=row, column=1, value=iterable)

    # ------------------------------------------------------------------
    # Lazy iter_rows (используется в бенчмарках и тестах)
    # ------------------------------------------------------------------

    def iter_rows(self, bint values_only = True):
        """Потоковое чтение строк (lazy режим)."""
        if self._archive is None:
            raise RuntimeError("iter_rows доступен только в режиме lazy")

        cdef str ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
        cdef str row_tag = f"{ns}row"
        cdef str c_tag = f"{ns}c"
        cdef str v_tag = f"{ns}v"
        cdef str f_tag = f"{ns}f"
        cdef list out
        cdef str t, raw, ftext
        cdef object value, v, f
        cdef int idx

        with self._archive.open(self._sheet_path) as fh:
            for _event, row in etree.iterparse(fh, tag=row_tag, events=("end",)):
                out = []
                for c in row.iter(c_tag):
                    # Формула має приоритет над <v>
                    f = c.find(f_tag)
                    if f is not None and f.text is not None:
                        ftext = f.text
                        value = '=' + ftext
                        out.append(value if values_only else value)
                        continue

                    t = c.get("t")
                    v = c.find(v_tag)
                    if v is None:
                        out.append(None)
                        continue
                    raw = v.text
                    if raw is None:
                        out.append(None)
                        continue

                    if t == "s":
                        try:
                            idx = int(raw)
                            if 0 <= idx < len(self._shared_strings):
                                value = self._shared_strings[idx]
                            else:
                                value = raw
                        except (ValueError, TypeError):
                            value = raw
                    else:
                        value = self._convert_cell_value(raw)

                    out.append(value if values_only else raw)
                yield tuple(out)
                row.clear()

    cdef object _convert_cell_value(self, str raw_value):
        """Быстрое преобразование значения ячейки (используется в lazy-режиме)."""
        if not raw_value:
            return raw_value
        try:
            if '.' in raw_value or 'e' in raw_value.lower():
                return float(raw_value)
            return int(raw_value)
        except (ValueError, OverflowError, TypeError):
            return raw_value

    cdef object _convert_cell_value_fast(self, str raw_value):
        """Максимально быстрый конвертер для парсинга всього листа."""
        if not raw_value:
            return raw_value
        # Используем чисто Python-логику без cdef char для совместимости с Unicode
        first_char = raw_value[0]
        if first_char == '-' or ('0' <= first_char <= '9'):
            if '.' in raw_value or 'e' in raw_value.lower():
                try:
                    return float(raw_value)
                except (ValueError, OverflowError):
                    return raw_value
            else:
                try:
                    return int(raw_value)
                except (ValueError, OverflowError):
                    return raw_value
        return raw_value

    # ------------------------------------------------------------------
    # Вспомогательные функции для координат и диапазонов
    # ------------------------------------------------------------------

    cdef int _col_to_num(self, str col):
        cdef int num = 0
        cdef int c_val
        for ch in col:
            c_val = ord(ch)
            if 97 <= c_val <= 122:  # a-z
                c_val -= 32
            num = num * 26 + (c_val - ord('A') + 1)
        return num

    cdef str _num_to_col(self, int num):
        cdef str col = ""
        cdef int remainder
        while num > 0:
            num, remainder = divmod(num - 1, 26)
            col = chr(ord('A') + remainder) + col
        return col

    cdef tuple _parse_range(self, str range_string):
        cdef str start_ref, end_ref, start_col, end_col
        cdef int start_row, end_row, start_col_num, end_col_num, i

        start_ref, end_ref = range_string.split(':')

        i = 0
        while i < len(start_ref) and start_ref[i].isalpha():
            i += 1
        start_col = start_ref[:i]
        start_row = int(start_ref[i:])

        i = 0
        while i < len(end_ref) and end_ref[i].isalpha():
            i += 1
        end_col = end_ref[:i]
        end_row = int(end_ref[i:])

        start_col_num = self._col_to_num(start_col)
        end_col_num = self._col_to_num(end_col)

        return (start_ref, end_ref, start_col, start_row, end_col, end_row, start_col_num, end_col_num)

    # ------------------------------------------------------------------
    # Merge / unmerge
    # ------------------------------------------------------------------

    cpdef object merge_cells(self, str range_string):
        cdef str start_ref, end_ref, start_col, end_col, col, cell_ref
        cdef int start_row, end_row, start_col_num, end_col_num, row, col_num
        cdef object cell

        if ':' not in range_string:
            raise ValueError(f"Неверный формат диапазона: {range_string}. Ожидается формат 'A1:B2'")

        if range_string not in self._merged_cells:
            self._merged_cells.append(range_string)

        start_ref, end_ref, start_col, start_row, end_col, end_row, start_col_num, end_col_num = self._parse_range(range_string)

        for row in range(start_row, end_row + 1):
            for col_num in range(start_col_num, end_col_num + 1):
                col = self._num_to_col(col_num)
                cell_ref = f"{col}{row}"

                if cell_ref not in self._cells:
                    cell = Cell(position=cell_ref, parent=self)
                    self._cells[cell_ref] = cell
                else:
                    cell = self._cells[cell_ref]

                cell.is_merged_cell = True
                cell.merged_range = range_string

                if row != start_row or col_num != start_col_num:
                    cell.value = None

        return self._cells[start_ref]

    cpdef void unmerge_cells(self, str range_string):
        cdef str start_ref, end_ref, start_col, end_col, col, cell_ref
        cdef int start_row, end_row, start_col_num, end_col_num, row, col_num
        cdef object cell

        if range_string not in self._merged_cells:
            return

        self._merged_cells.remove(range_string)

        start_ref, end_ref, start_col, start_row, end_col, end_row, start_col_num, end_col_num = self._parse_range(range_string)

        for row in range(start_row, end_row + 1):
            for col_num in range(start_col_num, end_col_num + 1):
                col = self._num_to_col(col_num)
                cell_ref = f"{col}{row}"
                if cell_ref in self._cells:
                    cell = self._cells[cell_ref]
                    cell.is_merged_cell = False
                    cell.merged_range = None

    # ------------------------------------------------------------------
    # Генерация XML sheetData (значення + mergeCells)
    # ------------------------------------------------------------------

    cpdef bytes get_xml_data(self):
        cdef dict rows_data = {}
        cdef str cell_position, column, tag, escaped_value
        cdef int row, i
        cdef object cell_value
        cdef object cell
        cdef list cells_in_row, rows_xml, merged_cells_xml

        for cell_position, cell in self._cells.items():
            if cell.value is not None or cell.is_merged_cell:
                i = 0
                while i < len(cell_position) and cell_position[i].isalpha():
                    i += 1
                column = cell_position[:i]
                row = int(cell_position[i:])

                if row not in rows_data:
                    rows_data[row] = []

                if cell.is_merged_cell and cell_position != cell.merged_range.split(':')[0]:
                    continue

                cell_value = cell.value

                # Формулы
                if isinstance(cell_value, str) and cell_value.startswith('='):
                    tag = f'<c r="{cell_position}"><f>{cell_value[1:]}</f></c>'
                else:
                    style_attrs = ""
                    # TODO: подключить styles.xml и s="idx" позже

                    if isinstance(cell_value, (int, float)):
                        tag = f'<c r="{cell_position}" t="n"{style_attrs}><v>{cell_value}</v></c>'
                    elif cell_value is not None:
                        # строка
                        escaped_value = str(cell_value).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                        tag = f'<c r="{cell_position}" t="str"{style_attrs}><v>{escaped_value}</v></c>'
                    else:
                        tag = f'<c r="{cell_position}"{style_attrs}></c>'

                rows_data[row].append(tag)

        rows_xml = []
        for row in sorted(rows_data.keys()):
            cells_in_row = rows_data[row]
            rows_xml.append(f'<row r="{row}">{"".join(cells_in_row)}</row>')

        # Після генерації можна також оновити кеш max_row, max_column
        if rows_data:
            self._max_row = max(rows_data.keys())

        merged_cells_xml = []
        for merged_range in self._merged_cells:
            merged_cells_xml.append(f'<mergeCell ref="{merged_range}"/>')

        merged_cells_section = ""
        if merged_cells_xml:
            merged_cells_section = f"""
    <mergeCells count="{len(self._merged_cells)}">
        {"".join(merged_cells_xml)}
    </mergeCells>"""

        cdef str xml_content = f"""<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">
    <sheetData>
        {"".join(rows_xml)}
    </sheetData>{merged_cells_section}
</worksheet>"""

        return xml_content.encode('utf-8')
