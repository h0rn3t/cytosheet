from lxml import etree
import io
from collections.abc import MutableMapping

from .cell import Cell


cdef class RowDimension:
    cdef public int index
    cdef public double height
    cdef public bint hidden

    def __init__(self, int index):
        self.index = index
        self.height = 0.0
        self.hidden = False


cdef class ColumnDimension:
    cdef public str index
    cdef public double width
    cdef public bint hidden

    def __init__(self, str index):
        self.index = index
        self.width = 0.0
        self.hidden = False


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
    cdef public dict _xf_style_map  # xfId -> Style (повні стилі з styles.xml)
    cdef public dict _row_dimensions   # int -> RowDimension (внутреннее хранилище)
    cdef public dict _column_dimensions  # str -> ColumnDimension (внутреннее хранилище)
    cdef object _row_dim_container
    cdef object _col_dim_container
    cdef bytes _original_xml  # оригинальный XML листа
    cdef public bint _modified  # флаг изменения листа

    def __init__(
            self,
            list shared_strings,
            str title = "Sheet",
            object archive = None,
            str _sheet_path = None,
            bint _preload = True,
            dict xf_style_map = None
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
        self._xf_style_map = xf_style_map if xf_style_map is not None else {}
        self._row_dimensions = {}
        self._column_dimensions = {}
        self._row_dim_container = RowDimensionContainer(self)
        self._col_dim_container = ColumnDimensionContainer(self)
        self._original_xml = None
        self._modified = False
        
        # Если загружаем из архива - сохраняем оригинальный XML
        if self._archive is not None and self._sheet_path is not None:
            try:
                self._original_xml = self._archive.read(self._sheet_path)
            except Exception:
                pass

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

    cdef object _style_for_xf(self, int style_idx):
        """Повертає КОПІЮ повного Style за xfId (або None, якщо стилю немає).

        Копія потрібна, щоб правка стилю однієї комірки не "протікала" на інші
        комірки з тим самим xf (D-6). Стилі будує Workbook у _xf_style_map (D-1).
        """
        cdef object st
        if self._xf_style_map:
            st = self._xf_style_map.get(style_idx)
            if st is not None:
                return st.copy()
        return None

    cdef bint _apply_style_simple(self, object cell, str xml_str, int cell_start, int cell_end):
        """Парсить s="..." у діапазоні [cell_start,cell_end), проставляє cell._style_id
        і повний Style. Повертає True, якщо атрибут s= знайдено (для string-парсера)."""
        cdef int s_start = xml_str.find('s="', cell_start, cell_end)
        cdef int s_end, style_idx
        cdef object st
        if s_start == -1:
            # Немає s= → комірка успадковує дефолтний стиль книги (xf=0), як openpyxl
            st = self._style_for_xf(0)
            if st is not None:
                cell.style = st
            return False
        s_start += 3
        s_end = xml_str.find('"', s_start, cell_end)
        if s_end == -1:
            return False
        try:
            style_idx = int(xml_str[s_start:s_end])
        except ValueError:
            return False
        cell._style_id = style_idx
        st = self._style_for_xf(style_idx)
        if st is not None:
            cell.style = st
        return True

    # ------------------------------------------------------------------
    # Парсинг листа з XML
    # ------------------------------------------------------------------

    cdef void _parse_sheet(self, bytes xml_data):
        """Выбор между стандартным и чанковым парсером для средних/больших файлов."""
        cdef int xml_size = len(xml_data)
        cdef bint use_chunked_parsing = xml_size > 100000
        # очищаем внутренние словари размеров
        self._row_dimensions.clear()
        self._column_dimensions.clear()
        if use_chunked_parsing:
            self._parse_sheet_chunked(xml_data)
        else:
            self._parse_sheet_standard(xml_data)

    cdef void _parse_row_col_dimensions_simple(self, str xml_str):
        """Парсинг <row> и <col> для _parse_sheet_simple (DM-005, DM-006)."""
        cdef int pos, col_start, col_end, attr_start, attr_end
        cdef str col_tag
        cdef int min_idx, max_idx, c
        cdef double width
        cdef bint hidden
        cdef int row_tag_start, row_tag_end, row_idx
        cdef str row_tag
        cdef double ht
        cdef bint row_hidden
        cdef RowDimension rdim
        cdef ColumnDimension cdim
        cdef str col_letter

        self._column_dimensions.clear()
        pos = 0
        while True:
            col_start = xml_str.find('<col ', pos)
            if col_start == -1:
                break
            col_end = xml_str.find('/>', col_start)
            if col_end == -1:
                break
            col_tag = xml_str[col_start:col_end]

            min_idx = 0
            max_idx = 0
            width = 0.0
            hidden = False

            attr_start = col_tag.find('min="')
            if attr_start != -1:
                attr_start += 5
                attr_end = col_tag.find('"', attr_start)
                if attr_end != -1:
                    try:
                        min_idx = int(col_tag[attr_start:attr_end])
                    except ValueError:
                        min_idx = 0

            attr_start = col_tag.find('max="')
            if attr_start != -1:
                attr_start += 5
                attr_end = col_tag.find('"', attr_start)
                if attr_end != -1:
                    try:
                        max_idx = int(col_tag[attr_start:attr_end])
                    except ValueError:
                        max_idx = min_idx

            attr_start = col_tag.find('width="')
            if attr_start != -1:
                attr_start += 7
                attr_end = col_tag.find('"', attr_start)
                if attr_end != -1:
                    try:
                        width = float(col_tag[attr_start:attr_end])
                    except ValueError:
                        width = 0.0

            hidden = 'hidden="1"' in col_tag or 'hidden="true"' in col_tag

            if min_idx <= 0:
                min_idx = 1
            if max_idx <= 0:
                max_idx = min_idx

            for c in range(min_idx, max_idx + 1):
                col_letter = self._num_to_col(c)
                cdim = self._column_dimensions.get(col_letter)
                if cdim is None:
                    cdim = ColumnDimension(col_letter)
                    self._column_dimensions[col_letter] = cdim
                if width > 0.0:
                    cdim.width = width
                if hidden:
                    cdim.hidden = True

            pos = col_end + 2

        self._row_dimensions.clear()
        pos = 0
        while True:
            row_tag_start = xml_str.find('<row ', pos)
            if row_tag_start == -1:
                break
            row_tag_end = xml_str.find('>', row_tag_start)
            if row_tag_end == -1:
                break
            row_tag = xml_str[row_tag_start:row_tag_end]

            row_idx = 0
            ht = 0.0
            row_hidden = False

            attr_start = row_tag.find('r="')
            if attr_start != -1:
                attr_start += 3
                attr_end = row_tag.find('"', attr_start)
                if attr_end != -1:
                    try:
                        row_idx = int(row_tag[attr_start:attr_end])
                    except ValueError:
                        row_idx = 0

            attr_start = row_tag.find('ht="')
            if attr_start != -1:
                attr_start += 4
                attr_end = row_tag.find('"', attr_start)
                if attr_end != -1:
                    try:
                        ht = float(row_tag[attr_start:attr_end])
                    except ValueError:
                        ht = 0.0

            row_hidden = 'hidden="1"' in row_tag or 'hidden="true"' in row_tag

            if row_idx > 0 and (ht > 0.0 or row_hidden):
                rdim = self._row_dimensions.get(row_idx)
                if rdim is None:
                    rdim = RowDimension(row_idx)
                    self._row_dimensions[row_idx] = rdim
                if ht > 0.0:
                    rdim.height = ht
                if row_hidden:
                    rdim.hidden = True

            pos = row_tag_end + 1

    cdef void _parse_sheet_simple(self, bytes xml_data):
        """Упрощённый string-based парсер для малых файлов с поддержкой формул.<f> и полного Style."""
        cdef str xml_str = xml_data.decode('utf-8')
        self._parse_row_col_dimensions_simple(xml_str)

        # затем парсим <row> для row_dimensions и сами ячейки как раньше
        cdef int start_pos = 0
        cdef int cell_start, cell_end
        cdef int r_start, r_end, t_start, t_end, v_start, v_end, f_start, f_end
        cdef str cell_ref, cell_type, raw_value, formula_text
        cdef object value, cell
        cdef int shared_index
        cdef int is_start, t2_start, t2_end
        cdef int row_tag_start, row_tag_end, row_idx
        cdef str row_tag
        cdef double ht
        cdef bint row_hidden, is_self_closing
        cdef int closing_tag_end

        # парсим <row> атрибуты
        pos = 0
        while True:
            row_tag_start = xml_str.find('<row ', pos)
            if row_tag_start == -1:
                break
            row_tag_end = xml_str.find('>', row_tag_start)
            if row_tag_end == -1:
                break
            row_tag = xml_str[row_tag_start:row_tag_end]
            row_idx = 0
            ht = 0.0
            row_hidden = False

            attr_start = row_tag.find('r="')
            if attr_start != -1:
                attr_start += 3
                attr_end = row_tag.find('"', attr_start)
                if attr_end != -1:
                    try:
                        row_idx = int(row_tag[attr_start:attr_end])
                    except ValueError:
                        row_idx = 0

            attr_start = row_tag.find('ht="')
            if attr_start != -1:
                attr_start += 4
                attr_end = row_tag.find('"', attr_start)
                if attr_end != -1:
                    try:
                        ht = float(row_tag[attr_start:attr_end])
                    except ValueError:
                        ht = 0.0

            row_hidden = 'hidden="1"' in row_tag or 'hidden="true"' in row_tag

            if row_idx > 0 and (ht > 0.0 or row_hidden):
                rdim = self._row_dimensions.get(row_idx)
                if rdim is None:
                    rdim = RowDimension(row_idx)
                    self._row_dimensions[row_idx] = rdim
                if ht > 0.0:
                    rdim.height = ht
                if row_hidden:
                    rdim.hidden = True

            pos = row_tag_end + 1

        # ниже — существующий цикл парсинга <c> ячеек
        while True:
            # openpyxl генерирует теги вида '<c r="A1" s="1" t="n">', поэтому ищем по сигнатуре '<c r="'
            cell_start = xml_str.find('<c r="', start_pos)
            if cell_start == -1:
                break
            
            # Проверяем самозакрывающийся тег или обычный
            cell_end = xml_str.find('/>', cell_start)
            closing_tag_end = xml_str.find('</c>', cell_start)
            
            # Выбираем ближайший конец
            if cell_end != -1 and (closing_tag_end == -1 or cell_end < closing_tag_end):
                # Самозакрывающийся тег <c ... />
                is_self_closing = True
            else:
                # Обычный тег <c ...>...</c>
                cell_end = closing_tag_end
                is_self_closing = False
                
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
            
            # Если это самозакрывающийся тег - создаем пустую ячейку со стилем
            if is_self_closing:
                cell = Cell(position=cell_ref, value=None, parent=self)
                self._apply_style_simple(cell, xml_str, cell_start, cell_end)
                self._cells[cell_ref] = cell
                start_pos = cell_end + 2  # +2 для />
                continue

            # t="s" / t="str" / t="n" / t="inlineStr" / ...
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
                    cell = Cell(position=cell_ref, value=value, parent=self)
                    cell.data_type = 'f'
                    self._apply_style_simple(cell, xml_str, cell_start, cell_end)
                    self._cells[cell_ref] = cell
                    start_pos = cell_end + 4
                    continue

            # inlineStr: <c t="inlineStr"><is><t>Text</t></is></c>
            if cell_type == 'inlineStr':
                is_start = xml_str.find('<is>', cell_start, cell_end)
                if is_start != -1:
                    t2_start = xml_str.find('<t>', is_start, cell_end)
                    if t2_start != -1:
                        t2_start += 3
                        t2_end = xml_str.find('</t>', t2_start, cell_end)
                        if t2_end != -1:
                            raw_value = xml_str[t2_start:t2_end]
                            value = raw_value
                            cell = Cell(position=cell_ref, value=value, parent=self)
                            self._apply_style_simple(cell, xml_str, cell_start, cell_end)
                            self._cells[cell_ref] = cell
                            start_pos = cell_end + 4
                            continue
                start_pos = cell_end + 4
                continue

            # Обычный путь через <v>
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
                        # Сохраняем индекс для последующей генерации XML
                        cell = Cell(position=cell_ref, value=value, parent=self)
                        cell._shared_string_index = shared_index
                        cell.data_type = 's'
                    else:
                        value = raw_value
                        cell = Cell(position=cell_ref, value=value, parent=self)
                except (ValueError, TypeError):
                    value = raw_value
                    cell = Cell(position=cell_ref, value=value, parent=self)
            else:
                value = self._convert_cell_value_fast(raw_value)
                cell = Cell(position=cell_ref, value=value, parent=self)
                if cell_type:
                    cell.data_type = cell_type

            # Сохраняем ячейку если есть значение ИЛИ стиль
            has_style = self._apply_style_simple(cell, xml_str, cell_start, cell_end)

            if value is not None or has_style:
                self._cells[cell_ref] = cell

            start_pos = cell_end + 4
        
        # Парсим mergeCells
        merge_start = xml_str.find('<mergeCells')
        if merge_start != -1:
            merge_end = xml_str.find('</mergeCells>', merge_start)
            if merge_end != -1:
                merge_section = xml_str[merge_start:merge_end + 13]
                # Ищем все mergeCell ref="..."
                pos = 0
                while True:
                    ref_start = merge_section.find('ref="', pos)
                    if ref_start == -1:
                        break
                    ref_start += 5
                    ref_end = merge_section.find('"', ref_start)
                    if ref_end == -1:
                        break
                    merge_ref = merge_section[ref_start:ref_end]
                    if merge_ref not in self._merged_cells:
                        self._merged_cells.append(merge_ref)
                    pos = ref_end + 1

    cdef void _parse_row_col_dimensions_iterparse(self, bytes xml_data, bint recover=False):
        """Парсинг <col>/<row> для iterparse-парсеров (DM-005, DM-006)."""
        cdef object xml_stream = io.BytesIO(xml_data)
        cdef object context
        cdef object col_elem, row_elem
        cdef int min_idx, max_idx, c, r_idx
        cdef double width, ht
        cdef bint hidden, row_hidden
        cdef ColumnDimension cdim
        cdef RowDimension rdim
        cdef str col_letter

        self._column_dimensions.clear()
        context = etree.iterparse(
            xml_stream,
            events=('end',),
            tag='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}col',
            recover=recover
        )
        try:
            for event, col_elem in context:
                try:
                    min_idx = int(col_elem.get('min', '0'))
                except Exception:
                    min_idx = 0
                try:
                    max_idx = int(col_elem.get('max', str(min_idx)))
                except Exception:
                    max_idx = min_idx
                try:
                    width = float(col_elem.get('width', '0') or '0')
                except Exception:
                    width = 0.0
                hidden = col_elem.get('hidden') in ('1', 'true')

                if min_idx <= 0:
                    min_idx = 1
                if max_idx <= 0:
                    max_idx = min_idx

                for c in range(min_idx, max_idx + 1):
                    col_letter = self._num_to_col(c)
                    cdim = self._column_dimensions.get(col_letter)
                    if cdim is None:
                        cdim = ColumnDimension(col_letter)
                        self._column_dimensions[col_letter] = cdim
                    if width > 0.0:
                        cdim.width = width
                    if hidden:
                        cdim.hidden = True
                col_elem.clear()
        except Exception:
            pass

        xml_stream.seek(0)
        self._row_dimensions.clear()
        context = etree.iterparse(
            xml_stream,
            events=('end',),
            tag='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row',
            recover=recover
        )
        try:
            for event, row_elem in context:
                try:
                    r_idx = int(row_elem.get('r', '0'))
                except Exception:
                    r_idx = 0
                try:
                    ht = float(row_elem.get('ht', '0') or '0')
                except Exception:
                    ht = 0.0
                row_hidden = row_elem.get('hidden') in ('1', 'true')

                if r_idx > 0 and (ht > 0.0 or row_hidden):
                    rdim = self._row_dimensions.get(r_idx)
                    if rdim is None:
                        rdim = RowDimension(r_idx)
                        self._row_dimensions[r_idx] = rdim
                    if ht > 0.0:
                        rdim.height = ht
                    if row_hidden:
                        rdim.hidden = True
                row_elem.clear()
        except Exception:
            pass

    cdef void _parse_sheet_standard(self, bytes xml_data):
        """Стандартный iterparse-парсер для средних файлов с поддержкой <f>."""
        self._parse_row_col_dimensions_iterparse(xml_data, recover=False)
        cdef dict ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        cdef object xml_stream = io.BytesIO(xml_data)

        # повторный проход по <col> для совместимости (заполняем _column_dimensions тем же форматом)
        cdef object context_cols = etree.iterparse(
            xml_stream,
            events=('end',),
            tag='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}col'
        )
        cdef object col_elem
        cdef int min_idx, max_idx, c
        cdef double width
        cdef bint hidden
        self._column_dimensions.clear()
        try:
            for event, col_elem in context_cols:
                try:
                    min_idx = int(col_elem.get('min', '0'))
                except Exception:
                    min_idx = 0
                try:
                    max_idx = int(col_elem.get('max', str(min_idx)))
                except Exception:
                    max_idx = min_idx
                try:
                    width = float(col_elem.get('width', '0') or '0')
                except Exception:
                    width = 0.0
                hidden = col_elem.get('hidden') in ('1', 'true')
                if min_idx <= 0:
                    min_idx = 1
                if max_idx <= 0:
                    max_idx = min_idx
                for c in range(min_idx, max_idx + 1):
                    col_letter = self._num_to_col(c)
                    dim = self._column_dimensions.get(col_letter)
                    if dim is None:
                        dim = ColumnDimension(col_letter)
                        self._column_dimensions[col_letter] = dim
                    if width > 0.0:
                        dim.width = width
                    if hidden:
                        dim.hidden = True
                col_elem.clear()
        except Exception:
            pass

        # перезапускаем поток для row/cell, так как iterparse уже прошёлся по нему
        xml_stream.seek(0)
        cdef object context = etree.iterparse(
            xml_stream,
            events=('end',),
            tag='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'
        )

        cdef str col_ref, cell_type, raw
        cdef int shared_string_index
        cdef object value, v_elem, f_elem, cell
        cdef dict temp_cells = {}
        cdef str s_attr
        cdef int style_idx
        cdef object style

        # также отдельным проходом по <row> считаем размеры строк
        xml_stream.seek(0)
        cdef object context_rows = etree.iterparse(
            xml_stream,
            events=('end',),
            tag='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row'
        )
        cdef object row_elem
        cdef int r_idx
        cdef double ht
        cdef bint row_hidden
        self._row_dimensions.clear()
        try:
            for event, row_elem in context_rows:
                r_idx = int(row_elem.get('r', '0'))
                ht = float(row_elem.get('ht', '0') or '0')
                row_hidden = row_elem.get('hidden') in ('1', 'true')
                if r_idx > 0 and (ht > 0.0 or row_hidden):
                    dim = self._row_dimensions.get(r_idx)
                    if dim is None:
                        dim = RowDimension(r_idx)
                        self._row_dimensions[r_idx] = dim
                    if ht > 0.0:
                        dim.height = ht
                    if row_hidden:
                        dim.hidden = True
                row_elem.clear()
        except Exception:
            pass

        # и снова перезапускаем поток для ячеек
        xml_stream.seek(0)
        context = etree.iterparse(
            xml_stream,
            events=('end',),
            tag='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'
        )

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
                    cell = Cell(position=col_ref, value=value, parent=self)
                    cell.data_type = 'f'
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
                                # Создаем ячейку и сохраняем индекс
                                cell = Cell(position=col_ref, value=value, parent=self)
                                cell._shared_string_index = shared_string_index
                                cell.data_type = 's'
                            else:
                                value = raw
                                cell = Cell(position=col_ref, value=value, parent=self)
                        except (ValueError, TypeError):
                            value = raw
                            cell = Cell(position=col_ref, value=value, parent=self)
                    else:
                        value = self._convert_cell_value_fast(raw)
                        if value is None:
                            cell_elem.clear()
                            continue
                        cell = Cell(position=col_ref, value=value, parent=self)
                        if cell_type:
                            cell.data_type = cell_type

                # Применяем полный Style по xfId и сохраняем style_id
                s_attr = cell_elem.get('s')
                if s_attr is not None:
                    try:
                        style_idx = int(s_attr)
                        cell._style_id = style_idx  # Сохраняем для генерации XML
                        style = self._style_for_xf(style_idx)
                        if style is not None:
                            cell.style = style
                    except ValueError:
                        pass
                else:
                    # Немає s= → дефолтний стиль книги (xf=0), как openpyxl
                    style = self._style_for_xf(0)
                    if style is not None:
                        cell.style = style

                temp_cells[col_ref] = cell
                cell_elem.clear()
        except Exception as e:
            print(f"Ошибка стандартного парсинга: {e}")
        finally:
            self._cells.update(temp_cells)

    cdef void _parse_sheet_chunked(self, bytes xml_data):
        """Чанковый парсер для больших файлов с поддержкой <f>."""
        # сначала проинициализируем размеры строк/колонок на отдельном BytesIO
        self._parse_row_col_dimensions_iterparse(xml_data, recover=True)
        cdef dict ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

        # работаем через BytesIO, чтобы iterparse не пытался интерпретировать bytes как имя файла
        cdef object xml_stream = io.BytesIO(xml_data)

        # <col>
        cdef object context_cols = etree.iterparse(
            xml_stream,
            events=('end',),
            tag='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}col',
            recover=True
        )
        cdef object col_elem
        cdef int min_idx, max_idx, c
        cdef double width
        cdef bint hidden
        self._column_dimensions.clear()
        try:
            for event, col_elem in context_cols:
                try:
                    min_idx = int(col_elem.get('min', '0'))
                except Exception:
                    min_idx = 0
                try:
                    max_idx = int(col_elem.get('max', str(min_idx)))
                except Exception:
                    max_idx = min_idx
                try:
                    width = float(col_elem.get('width', '0') or '0')
                except Exception:
                    width = 0.0
                hidden = col_elem.get('hidden') in ('1', 'true')
                if min_idx <= 0:
                    min_idx = 1
                if max_idx <= 0:
                    max_idx = min_idx
                for c in range(min_idx, max_idx + 1):
                    col_letter = self._num_to_col(c)
                    dim = self._column_dimensions.get(col_letter)
                    if dim is None:
                        dim = ColumnDimension(col_letter)
                        self._column_dimensions[col_letter] = dim
                    if width > 0.0:
                        dim.width = width
                    if hidden:
                        dim.hidden = True
                col_elem.clear()
        except Exception:
            pass

        # <row>
        xml_stream.seek(0)
        cdef object context_rows = etree.iterparse(
            xml_stream,
            events=('end',),
            tag='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row',
            recover=True
        )
        cdef object row_elem
        cdef int r_idx
        cdef double ht
        cdef bint row_hidden
        self._row_dimensions.clear()
        try:
            for event, row_elem in context_rows:
                try:
                    r_idx = int(row_elem.get('r', '0'))
                except Exception:
                    r_idx = 0
                try:
                    ht = float(row_elem.get('ht', '0') or '0')
                except Exception:
                    ht = 0.0
                row_hidden = row_elem.get('hidden') in ('1', 'true')
                if r_idx > 0 and (ht > 0.0 or row_hidden):
                    dim = self._row_dimensions.get(r_idx)
                    if dim is None:
                        dim = RowDimension(r_idx)
                        self._row_dimensions[r_idx] = dim
                    if ht > 0.0:
                        dim.height = ht
                    if row_hidden:
                        dim.hidden = True
                row_elem.clear()
        except Exception:
            pass

        # ячейки с батчами
        xml_stream.seek(0)
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
        cdef str s_attr
        cdef int style_idx
        cdef object style

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
                    cell = Cell(position=col_ref, value=value, parent=self)
                    cell.data_type = 'f'
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
                                cell = Cell(position=col_ref, value=value, parent=self)
                                cell._shared_string_index = shared_string_index
                                cell.data_type = 's'
                            else:
                                value = raw
                                cell = Cell(position=col_ref, value=value, parent=self)
                        except (ValueError, TypeError):
                            value = raw
                            cell = Cell(position=col_ref, value=value, parent=self)
                    else:
                        value = self._convert_cell_value_fast(raw)
                        if value is None:
                            cell_elem.clear()
                            continue
                        cell = Cell(position=col_ref, value=value, parent=self)
                        if cell_type:
                            cell.data_type = cell_type

                # Применяем полный Style по xfId и сохраняем style_id
                s_attr = cell_elem.get('s')
                if s_attr is not None:
                    try:
                        style_idx = int(s_attr)
                        cell._style_id = style_idx
                        style = self._style_for_xf(style_idx)
                        if style is not None:
                            cell.style = style
                    except ValueError:
                        pass
                else:
                    # Немає s= → дефолтний стиль книги (xf=0), как openpyxl
                    style = self._style_for_xf(0)
                    if style is not None:
                        cell.style = style

                temp_batch[col_ref] = cell
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

    cdef object _convert_cell_value_fast(self, str raw):
        """Быстрое преобразование строкового значения ячейки в число или None.

        Используется при парсинге XML, чтобы не тянуть openpyxl-логику.
        """
        if raw is None:
            return None
        if raw == "":
            return None
        try:
            if "." in raw:
                return float(raw)
            return int(raw)
        except ValueError:
            return raw

    # ------------------------------------------------------------------
    # Доступ і запис ячеек
    # ------------------------------------------------------------------

    def __getitem__(self, key):
        """Поддержка одиночной ячейки, диапазонов и доступа по номеру строки, как в openpyxl."""
        cdef object cell
        
        # Поддержка доступа по номеру строки: ws[1] возвращает генератор ячеек строки
        if isinstance(key, int):
            return next(self.iter_rows(min_row=key, max_row=key))
        
        # Диапазон вида "A1:C3"
        if ":" in key:
            return self._get_range(key)

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
            # используем set_value, чтобы централизованно обновлять data_type
            if hasattr(cell, 'set_value'):
                cell.set_value(value)
            else:
                cell.value = value
                cell._shared_string_index = -1
                if isinstance(value, str) and value.startswith('='):
                    cell.data_type = 'f'

        self._update_bounds_for_cell(key)
        # Присвоєння через індексатор змінює лист — інакше save завантаженої
        # книги повернув би оригінальний XML і втратив зміну (D-3).
        self._modified = True

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
            if hasattr(cell, 'set_value'):
                cell.set_value(value)
            else:
                cell.value = value
                if isinstance(value, str) and value.startswith('='):
                    cell.data_type = 'f'
            # Помечаем лист как измененный
            self._modified = True

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

    @property
    def dimensions(self):
        """Возвращает строку вида 'A1:D10' для використовуваного діапазону, як в openpyxl."""
        if not self._cells:
            return "A1:A1"

        # Убедимся, что границы актуальны
        if self._max_row == 0 or self._max_column == 0:
            self._recalculate_bounds()

        # Минимальные границы считаем по ключам _cells
        cdef int min_row = 0
        cdef int min_col = 0
        cdef str coord
        cdef int i, row, col

        for coord in self._cells.keys():
            if coord is None:
                continue
            i = 0
            while i < len(coord) and coord[i].isalpha():
                i += 1
            if i == 0 or i == len(coord):
                continue
            row = int(coord[i:])
            col = self._col_to_num(coord[:i])

            if min_row == 0 or row < min_row:
                min_row = row
            if min_col == 0 or col < min_col:
                min_col = col

        if min_row == 0 or min_col == 0:
            # fallback, якщо по якійсь причині не вдалося вирахувати
            min_row = 1
            min_col = 1

        cdef str min_col_letter = self._num_to_col(min_col)
        cdef str max_col_letter = self._num_to_col(self.max_column)
        return f"{min_col_letter}{min_row}:{max_col_letter}{self.max_row}"

    def append(self, object iterable):
        """Добавить строку значений в конец листа (openpyxl-совместный append)."""
        if iterable is None:
            return

        cdef int row = self.max_row + 1
        cdef int col

        # Строки и bytes считаем скалярами, как в openpyxl: кладём целиком в первый столбец.
        if isinstance(iterable, (str, bytes)):
            self.cell(row=row, column=1, value=iterable)
            return

        try:
            iterator = iter(iterable)
        except TypeError:
            # Неитерируемый объект -> в первый столбец
            self.cell(row=row, column=1, value=iterable)
            return

        col = 1
        for value in iterator:
            self.cell(row=row, column=col, value=value)
            col += 1

    @property
    def rows(self):
        """Генератор всіх строк (кортежі Cell), сумісний з openpyxl."""
        return self.iter_rows()

    @property
    def columns(self):
        """Генератор всіх колонок (кортежі Cell), сумісний з openpyxl."""
        return self.iter_cols()

    @property
    def values(self):
        """Генератор тільки значень по строках, як ws.iter_rows(values_only=True)."""
        return self.iter_rows(values_only=True)

    @property
    def row_dimensions(self):
        """Публичное API, совместимое с openpyxl: ws.row_dimensions[index]."""
        return self._row_dim_container

    @property
    def column_dimensions(self):
        """Публичное API, совместимое с openpyxl: ws.column_dimensions['A']."""
        return self._col_dim_container

    # ------------------------------------------------------------------
    # Итерация по строкам и колонкам
    # ------------------------------------------------------------------

    def iter_rows(self, min_row=1, max_row=None, min_col=1, max_col=None, bint values_only=False):
        """Итерация по строкам, совместимая с openpyxl.

        В lazy-режиме при первом вызове подгружает данные листа из архива.
        """
        # Ленивая подзагрузка для lazy=True
        if not self._preloaded and self._archive is not None and self._sheet_path is not None:
            xml = self._archive.read(self._sheet_path)
            if self._is_small_file:
                self._parse_sheet_simple(xml)
            else:
                self._parse_sheet(xml)
            self._recalculate_bounds()
            self._preloaded = True

        cdef int r_min, r_max, c_min, c_max
        cdef int row, col
        cdef object cell

        r_min = 1 if min_row is None else int(min_row)
        c_min = 1 if min_col is None else int(min_col)
        r_max = self.max_row if max_row is None else int(max_row)
        c_max = self.max_column if max_col is None else int(max_col)

        for row in range(r_min, r_max + 1):
            row_cells = []
            for col in range(c_min, c_max + 1):
                cell = self.cell(row=row, column=col)
                row_cells.append(cell.value if values_only else cell)
            yield tuple(row_cells)

    def iter_cols(self, min_col=1, max_col=None, min_row=1, max_row=None, bint values_only=False):
        """Итерация по колонкам, совместимая с openpyxl.

        В lazy-режиме при первом вызове подгружает данные листа из архива.
        """
        if not self._preloaded and self._archive is not None and self._sheet_path is not None:
            xml = self._archive.read(self._sheet_path)
            if self._is_small_file:
                self._parse_sheet_simple(xml)
            else:
                self._parse_sheet(xml)
            self._recalculate_bounds()
            self._preloaded = True

        cdef int r_min, r_max, c_min, c_max
        cdef int row, col
        cdef object cell

        c_min = 1 if min_col is None else int(min_col)
        r_min = 1 if min_row is None else int(min_row)
        c_max = self.max_column if max_col is None else int(max_col)
        r_max = self.max_row if max_row is None else int(max_row)

        for col in range(c_min, c_max + 1):
            col_cells = []
            for row in range(r_min, r_max + 1):
                cell = self.cell(row=row, column=col)
                col_cells.append(cell.value if values_only else cell)
            yield tuple(col_cells)

    cdef int _col_to_num(self, str col):
        """Преобразовать 'A' -> 1, 'Z' -> 26, 'AA' -> 27 и т.п.

        Избегаем использования cdef char, чтобы не ловить артефакты при итерации по Python-строке.
        """
        cdef int result = 0
        cdef object ch
        for ch in col:
            result = result * 26 + (ord((<str>ch).upper()) - ord('A') + 1)
        return result

    # ------------------------------------------------------------------
    # Диапазоны и утилиты конвертации колонок
    # ------------------------------------------------------------------

    cdef tuple _split_coord(self, str coord):
        """Разбить 'A10' на (col_letters, row_int)."""
        cdef int i = 0
        cdef int n = len(coord)
        while i < n and coord[i].isalpha():
            i += 1
        return coord[:i], int(coord[i:])

    cdef str _num_to_col(self, int col):
        """Обратное преобразование: 1 -> 'A', 27 -> 'AA'."""
        if col <= 0:
            raise ValueError("Column index must be >= 1")
        cdef list letters = []
        cdef int n = col
        cdef int rem
        while n > 0:
            n, rem = divmod(n - 1, 26)
            letters.append(chr(ord('A') + rem))
        letters.reverse()
        return "".join(letters)

    cdef object _get_range(self, str range_string):
        """Вернуть диапазон ws['A1:C3'] как tuple(tuple(Cell))."""
        cdef str start_ref, end_ref
        cdef str start_col_letters, end_col_letters
        cdef int start_row, end_row, start_col, end_col
        start_ref, end_ref = range_string.split(":", 1)

        start_col_letters, start_row = self._split_coord(start_ref)
        end_col_letters, end_row = self._split_coord(end_ref)

        start_col = self._col_to_num(start_col_letters)
        end_col = self._col_to_num(end_col_letters)

        cdef list rows = []
        cdef int row, col
        cdef list row_cells
        for row in range(start_row, end_row + 1):
            row_cells = []
            for col in range(start_col, end_col + 1):
                row_cells.append(self.cell(row=row, column=col))
            rows.append(tuple(row_cells))
        return tuple(rows)

    cdef tuple _parse_range(self, str range_string):
        """Разбор диапазона вида 'A1:C3' в составные части.

        Возвращает кортеж:
        (start_ref, end_ref, start_col, start_row, end_col, end_row,
         start_col_num, end_col_num)
        """
        cdef str start_ref, end_ref
        cdef str start_col, end_col
        cdef int start_row, end_row
        cdef int start_col_num, end_col_num

        if ':' not in range_string:
            raise ValueError(f"Неверный формат диапазона: {range_string}. Ожидается формат 'A1:B2'")

        start_ref, end_ref = range_string.split(':', 1)

        start_col, start_row = self._split_coord(start_ref)
        end_col, end_row = self._split_coord(end_ref)

        start_col_num = self._col_to_num(start_col)
        end_col_num = self._col_to_num(end_col)

        return (start_ref, end_ref,
                start_col, start_row,
                end_col, end_row,
                start_col_num, end_col_num)

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

        self._modified = True

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
        self._modified = True

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
        """Генерация XML содержимого листа + <row>/<col> для dimensions.
        
        Если лист не был изменен и есть оригинальный XML - возвращаем его.
        """
        # Если лист не изменялся и есть оригинальный XML - возвращаем его
        if not self._modified and self._original_xml is not None:
            return self._original_xml
            
        cdef dict rows_data = {}
        cdef str cell_position, column
        cdef int row, i
        cdef object cell_value, cell
        cdef list cells_in_row, rows_xml, merged_cells_xml
        cdef str tag, escaped_value, style_attr
        cdef RowDimension rdim
        cdef list cols_xml
        cdef str col_letter

        # Создаем список (row, col_num, cell_position) для правильной сортировки
        cdef list cell_sort_list = []
        cdef int col_num
        for cell_position in self._cells.keys():
            i = 0
            while i < len(cell_position) and cell_position[i].isalpha():
                i += 1
            column = cell_position[:i]
            row = int(cell_position[i:])
            col_num = self._col_to_num(column)
            cell_sort_list.append((row, col_num, cell_position))
        
        # Сортируем по строке, затем по номеру колонки
        cell_sort_list.sort()

        # собираем данные ячеек
        for row, col_num, cell_position in cell_sort_list:
            cell = self._cells[cell_position]
            # Серіалізуємо комірку зі значенням, merge-частиною АБО власним стилем
            # (порожня стильована комірка має зберегтись як <c r=".." s="N"/>, як в openpyxl).
            if cell.value is not None or getattr(cell, 'is_merged_cell', False) or cell._style_id >= 0:
                i = 0
                while i < len(cell_position) and cell_position[i].isalpha():
                    i += 1
                column = cell_position[:i]
                row = int(cell_position[i:])

                if row not in rows_data:
                    rows_data[row] = []

                # пропускаем вторичные ячейки merge-діапазона
                if getattr(cell, 'is_merged_cell', False) and cell_position != cell.merged_range.split(':')[0]:
                    continue

                cell_value = cell.value

                style_attr = ""
                if hasattr(cell, '_style_id') and getattr(cell, '_style_id') is not None and getattr(cell, '_style_id') >= 0:
                    style_attr = f' s="{cell._style_id}"'

                # Если ячейка была загружена из файла и имеет ссылку на sharedString - используем её
                try:
                    if cell._shared_string_index >= 0:
                        tag = f'<c r="{cell_position}" s="{cell._style_id if cell._style_id >= 0 else 0}" t="s"><v>{cell._shared_string_index}</v></c>'
                        rows_data[row].append(tag)
                        continue
                except (AttributeError, TypeError):
                    pass
                
                # Булевые значения должны сериализоваться как t="b" c 1/0,
                # иначе openpyxl пытается парсить "True"/"False" как число и падает.
                if isinstance(cell_value, bool):
                    tag = f'<c r="{cell_position}" t="b"{style_attr}><v>{"1" if cell_value else "0"}</v></c>'
                elif isinstance(cell_value, str) and cell_value.startswith('='):
                    tag = f'<c r="{cell_position}"{style_attr}><f>{cell_value[1:]}</f></c>'
                else:
                    if isinstance(cell_value, (int, float)):
                        tag = f'<c r="{cell_position}" t="n"{style_attr}><v>{cell_value}</v></c>'
                    elif cell_value is not None:
                        escaped_value = str(cell_value).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                        tag = f'<c r="{cell_position}" t="str"{style_attr}><v>{escaped_value}</v></c>'
                    else:
                        # Порожня комірка лише зі стилем — self-closing, як в openpyxl
                        tag = f'<c r="{cell_position}"{style_attr}/>'

                rows_data[row].append(tag)

        rows_xml = []
        for row in sorted(rows_data.keys()):
            cells_in_row = rows_data[row]
            # Берём RowDimension напрямую из внутреннего словаря
            rdim = self._row_dimensions.get(row)
            if rdim is not None and (rdim.height > 0.0 or rdim.hidden):
                if rdim.height > 0.0 and rdim.hidden:
                    rows_xml.append(f'<row r="{row}" ht="{rdim.height}" customHeight="1" hidden="1">' + "".join(cells_in_row) + '</row>')
                elif rdim.height > 0.0:
                    rows_xml.append(f'<row r="{row}" ht="{rdim.height}" customHeight="1">' + "".join(cells_in_row) + '</row>')
                elif rdim.hidden:
                    rows_xml.append(f'<row r="{row}" hidden="1">' + "".join(cells_in_row) + '</row>')
            else:
                rows_xml.append(f'<row r="{row}">{"".join(cells_in_row)}</row>')

        if rows_data:
            self._max_row = max(rows_data.keys())

        cols_xml = []
        for col_letter, cdim in sorted(self._column_dimensions.items()):
            i = self._col_to_num(col_letter)
            cols_xml.append(
                f'<col min="{i}" max="{i}"'
                + (f' width="{cdim.width}" customWidth="1"' if cdim.width > 0.0 else '')
                + (' hidden="1"' if cdim.hidden else '')
                + '/>'
            )

        merged_cells_xml = []
        for merged_range in self._merged_cells:
            merged_cells_xml.append(f'<mergeCell ref="{merged_range}"/>')

        cdef str merged_cells_section = ""
        if merged_cells_xml:
            merged_cells_section = f"""\n    <mergeCells count=\"{len(self._merged_cells)}\">\n        {"".join(merged_cells_xml)}\n    </mergeCells>"""

        cdef str cols_section = ""
        if cols_xml:
            cols_section = "\n    <cols>" + "".join(cols_xml) + "</cols>"

        cdef str xml_content = f"""<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">{cols_section}
    <sheetData>
        {"".join(rows_xml)}
    </sheetData>{merged_cells_section}
</worksheet>"""

        return xml_content.encode('utf-8')


# ------------------------------------------------------------------
# Контейнеры размеров в стиле openpyxl
# ------------------------------------------------------------------

class RowDimensionContainer(MutableMapping):
    """Контейнер-доступ к размерам строк: ws.row_dimensions[index]."""

    def __init__(self, ws):
        self._ws = ws

    def __getitem__(self, key):
        idx = int(key)
        rdim = self._ws._row_dimensions.get(idx)
        if rdim is None:
            rdim = RowDimension(idx)
            self._ws._row_dimensions[idx] = rdim
        # Доступ до розмірів зазвичай передує їх зміні — позначаємо лист зміненим,
        # щоб save завантаженої книги відобразив нову висоту/ширину (D-3).
        self._ws._modified = True
        return rdim

    def __setitem__(self, key, value):
        idx = int(key)
        if not isinstance(value, RowDimension):
            raise TypeError("value must be RowDimension")
        self._ws._row_dimensions[idx] = value

    def __delitem__(self, key):
        idx = int(key)
        if idx in self._ws._row_dimensions:
            del self._ws._row_dimensions[idx]

    def __iter__(self):
        return iter(self._ws._row_dimensions)

    def __len__(self):
        return len(self._ws._row_dimensions)


class ColumnDimensionContainer(MutableMapping):
    """Контейнер-доступ к размерам колонок: ws.column_dimensions['A']."""

    def __init__(self, ws):
        self._ws = ws

    def __getitem__(self, key):
        col = str(key)
        cdim = self._ws._column_dimensions.get(col)
        if cdim is None:
            cdim = ColumnDimension(col)
            self._ws._column_dimensions[col] = cdim
        self._ws._modified = True
        return cdim

    def __setitem__(self, key, value):
        col = str(key)
        if not isinstance(value, ColumnDimension):
            raise TypeError("value must be ColumnDimension")
        self._ws._column_dimensions[col] = value

    def __delitem__(self, key):
        col = str(key)
        if col in self._ws._column_dimensions:
            del self._ws._column_dimensions[col]

    def __iter__(self):
        return iter(self._ws._column_dimensions)

    def __len__(self):
        return len(self._ws._column_dimensions)


# Публичные свойства Worksheet для совместимости с openpyxl
