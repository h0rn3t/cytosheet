from lxml import etree
import io

cdef class Worksheet:
    cdef public dict _cells
    cdef public str title
    cdef public list _shared_strings
    cdef public object _archive  # ZipFile | None
    cdef public str _sheet_path
    cdef public bint _preloaded
    cdef public bint _is_small_file

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

        if archive is not None and _preload and _sheet_path is not None:
            xml = archive.read(_sheet_path)

            # Быстрая оценка размера файла по размеру XML
            xml_size = len(xml)
            self._is_small_file = xml_size < 50000  # ~500 строк

            # Выбираем оптимальный метод парсинга
            if self._is_small_file:
                self._parse_sheet_simple(xml)
            else:
                self._parse_sheet(xml)

            self._preloaded = True

    cdef void _parse_sheet_simple(self, bytes xml_data):
        """Упрощенный парсинг для малых файлов"""
        from .cell import Cell

        # Для малых файлов используем более простой и быстрый подход
        cdef str xml_str = xml_data.decode('utf-8')

        # Используем простой string parsing вместо полноценного XML
        cdef int start_pos = 0
        cdef int cell_start, cell_end, r_start, r_end, v_start, v_end, t_start, t_end
        cdef str cell_ref, cell_type, raw_value
        cdef object value
        cdef int shared_index

        # Ищем все ячейки простым поиском строк
        while True:
            cell_start = xml_str.find('<c ', start_pos)
            if cell_start == -1:
                break

            cell_end = xml_str.find('</c>', cell_start)
            if cell_end == -1:
                break

            # Извлекаем атрибут r (адрес ячейки)
            r_start = xml_str.find('r="', cell_start)
            if r_start == -1 or r_start > cell_end:
                start_pos = cell_end + 4
                continue

            r_start += 3
            r_end = xml_str.find('"', r_start)
            if r_end == -1 or r_end > cell_end:
                start_pos = cell_end + 4
                continue

            cell_ref = xml_str[r_start:r_end]

            # Извлекаем атрибут t (тип ячейки)
            t_start = xml_str.find('t="', cell_start)
            if t_start != -1 and t_start < cell_end:
                t_start += 3
                t_end = xml_str.find('"', t_start)
                if t_end != -1 and t_end < cell_end:
                    cell_type = xml_str[t_start:t_end]
                else:
                    cell_type = None
            else:
                cell_type = None

            # Извлекаем значение ячейки
            v_start = xml_str.find('<v>', cell_start)
            if v_start == -1 or v_start > cell_end:
                start_pos = cell_end + 4
                continue

            v_start += 3
            v_end = xml_str.find('</v>', v_start)
            if v_end == -1 or v_end > cell_end:
                start_pos = cell_end + 4
                continue

            raw_value = xml_str[v_start:v_end]

            # Быстрая обработка значения
            if cell_type == 's':  # shared string
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

    def __getitem__(self, str key):
        from .cell import Cell
        cdef object cell
        if key in self._cells:
            return self._cells[key]
        else:
            cell = Cell(position=key)
            self._cells[key] = cell
            return cell

    def __setitem__(self, str key, object value):
        from .cell import Cell
        cdef object cell
        if key not in self._cells:
            cell = Cell(position=key, value=value)
            self._cells[key] = cell
        else:
            self._cells[key].value = value

    def iter_rows(self, bint values_only = True):
        """
        Потоковое чтение строк. Работает только для листов,
        открытых из файла (self._archive != None).
        """
        if self._archive is None:
            raise RuntimeError("iter_rows доступен только в режиме lazy")

        cdef str ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
        cdef str row_tag = f"{ns}row"
        cdef str c_tag = f"{ns}c"
        cdef str v_tag = f"{ns}v"
        cdef list out
        cdef str t, raw
        cdef object value, v
        cdef int idx

        with self._archive.open(self._sheet_path) as fh:
            for _event, row in etree.iterparse(fh, tag=row_tag, events=("end",)):
                out = []
                for c in row.iter(c_tag):
                    t = c.get("t")
                    v = c.find(v_tag)
                    if v is None:
                        continue
                    raw = v.text
                    if raw is None:
                        continue

                    if t == "s":  # shared-string
                        try:
                            idx = int(raw)
                            if 0 <= idx < len(self._shared_strings):
                                value = self._shared_strings[idx]
                            else:
                                value = raw
                        except (ValueError, TypeError):
                            value = raw
                    else:  # число / строка
                        value = self._convert_cell_value(raw)

                    out.append(value if values_only else raw)
                yield out
                row.clear()

    cdef object _convert_cell_value(self, str raw_value):
        """Быстрое преобразование значения ячейки"""
        cdef long long_val
        cdef double float_val

        if not raw_value:
            return raw_value

        try:
            # Сначала пробуем int, но с проверкой размера
            if '.' not in raw_value and 'e' not in raw_value.lower() and 'E' not in raw_value:
                # Проверяем, что число не слишком большое для int
                if len(raw_value) <= 10:  # Примерно 2^31 = 2,147,483,647
                    long_val = int(raw_value)
                    # Проверяем диапазон для безопасности
                    if -2147483648 <= long_val <= 2147483647:
                        return int(long_val)
                    else:
                        return long_val  # Возвращаем как long
                else:
                    # Для очень больших чисел пробуем float
                    float_val = float(raw_value)
                    return float_val
            else:
                # Затем float
                float_val = float(raw_value)
                return float_val
        except (ValueError, TypeError, OverflowError):
            return raw_value

    cdef void _parse_sheet(self, bytes xml_data):
        """Максимально оптимизированный парсинг листа для больших файлов"""
        from .cell import Cell

        cdef dict ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        cdef object xml_stream = io.BytesIO(xml_data)

        # Для больших файлов используем более эффективную стратегию
        cdef int xml_size = len(xml_data)
        cdef bint use_chunked_parsing = xml_size > 100000  # >~1000 строк

        if use_chunked_parsing:
            self._parse_sheet_chunked(xml_data)
        else:
            self._parse_sheet_standard(xml_data)

    cdef void _parse_sheet_chunked(self, bytes xml_data):
        """Чанковый парсинг для больших файлов"""
        from .cell import Cell

        cdef dict ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        cdef object xml_stream = io.BytesIO(xml_data)

        # Используем более эффективный iterparse с очисткой памяти
        cdef object context = etree.iterparse(
            xml_stream,
            events=('end',),
            tag='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c',
            recover=True  # Более быстрый парсинг
        )

        cdef str col_ref, cell_type, raw
        cdef int shared_string_index, batch_size, processed_count
        cdef object value, v_elem, cell

        # Батчевая обработка для больших файлов
        cdef dict temp_batch = {}
        batch_size = 1000  # Обрабатываем по 1000 ячеек за раз
        processed_count = 0

        try:
            for event, cell_elem in context:
                # Быстрая проверка атрибутов
                col_ref = cell_elem.get('r')
                if col_ref is None:
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

                # Быстрое преобразование значения
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

                # Периодически сбрасываем батч в основной словарь
                if processed_count >= batch_size:
                    self._cells.update(temp_batch)
                    temp_batch.clear()
                    processed_count = 0

                # Важно: немедленно очищаем элемент
                cell_elem.clear()

        except Exception as e:
            print(f"Ошибка чанкового парсинга: {e}")
        finally:
            # Обрабатываем оставшиеся элементы
            if temp_batch:
                self._cells.update(temp_batch)

    cdef void _parse_sheet_standard(self, bytes xml_data):
        """Стандартный парсинг для средних файлов"""
        from .cell import Cell

        cdef dict ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        cdef object xml_stream = io.BytesIO(xml_data)
        cdef object context = etree.iterparse(
            xml_stream,
            events=('end',),
            tag='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'
        )

        cdef str col_ref, cell_type, raw
        cdef int shared_string_index
        cdef object value, v_elem, cell
        cdef dict temp_cells = {}

        try:
            for event, cell_elem in context:
                col_ref = cell_elem.get('r')
                if col_ref is None:
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

    cdef object _convert_cell_value_fast(self, str raw_value):
        """Максимально быстрое преобразование значения ячейки"""
        if not raw_value:
            return raw_value

        cdef char first_char = ord(raw_value[0])
        cdef int length = len(raw_value)

        # Быстрая проверка: если начинается с цифры или минуса - вероятно число
        if (48 <= first_char <= 57) or first_char == 45:  # '0'-'9' или '-'
            # Проверяем наличие точки для float
            if '.' in raw_value:
                try:
                    return float(raw_value)
                except (ValueError, OverflowError):
                    return raw_value
            else:
                # Для целых чисел проверяем длину
                if length <= 9:  # Безопасно для int32
                    try:
                        return int(raw_value)
                    except (ValueError, OverflowError):
                        return raw_value
                elif length <= 18:  # Может быть long
                    try:
                        return int(raw_value)
                    except (ValueError, OverflowError):
                        try:
                            return float(raw_value)
                        except (ValueError, OverflowError):
                            return raw_value
                else:
                    # Очень длинное число - делаем float
                    try:
                        return float(raw_value)
                    except (ValueError, OverflowError):
                        return raw_value

        # Если не число - возвращаем строку
        return raw_value

    cpdef bytes get_xml_data(self):
        """
        Оптимизированная генерация XML данных для листа.
        """
        cdef dict rows_data = {}
        cdef str cell_position, column, tag
        cdef int row, i
        cdef object cell_value
        cdef object cell
        cdef list cells_in_row, rows_xml

        # Группируем ячейки по строкам
        for cell_position, cell in self._cells.items():
            if cell.value is not None:
                # Более эффективное разделение позиции
                i = 0
                while i < len(cell_position) and cell_position[i].isalpha():
                    i += 1
                column = cell_position[:i]
                row = int(cell_position[i:])

                if row not in rows_data:
                    rows_data[row] = []

                cell_value = cell.value
                # Быстрое определение типа данных
                if isinstance(cell_value, (int, float)):
                    tag = f'<c r="{cell_position}" t="n"><v>{cell_value}</v></c>'
                else:
                    # Экранируем XML символы для строк
                    escaped_value = str(cell_value).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                    tag = f'<c r="{cell_position}" t="str"><v>{escaped_value}</v></c>'

                rows_data[row].append(tag)

        # Генерируем строки XML
        rows_xml = []
        for row in sorted(rows_data.keys()):
            cells_in_row = rows_data[row]
            rows_xml.append(f'<row r="{row}">{"".join(cells_in_row)}</row>')

        cdef str xml_content = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        {"".join(rows_xml)}
    </sheetData>
</worksheet>"""

        return xml_content.encode('utf-8')