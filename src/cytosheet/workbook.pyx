import os
from zipfile import ZipFile
from lxml import etree
from io import BytesIO
from .worksheet import Worksheet
from .styles import Style

# XML templates - делаем их константами на уровне модуля для производительности
cdef str WORKBOOK_XML_TEMPLATE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>{sheets}</sheets>
</workbook>"""

cdef str CONTENT_TYPES_XML_TEMPLATE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    {sheet_overrides}
</Types>"""

cdef str RELATIONSHIPS_XML_TEMPLATE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    {relationships}
</Relationships>"""

cdef str MAIN_RELATIONSHIPS_XML_TEMPLATE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

cdef str STYLES_XML_TEMPLATE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <numFmts count="{numFmts_count}">{numFmts}</numFmts>
    <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
    <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
    <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
    <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
    <cellXfs count="{cellXfs_count}">{cellXfs}</cellXfs>
    <!-- Минимальный набор для named style "Normal", как ожидает openpyxl -->
    <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""

cdef dict BUILTIN_NUMFMTS = {
    0: 'General',
    1: '0',
    2: '0.00',
    3: '#,##0',
    4: '#,##0.00',
    9: '0%',
    10: '0.00%',
    14: 'm/d/yy',
    22: 'm/d/yy h:mm',
}

cdef class Workbook:
    cdef public dict _sheets
    cdef public list _shared_strings
    cdef public int _active_sheet_index
    cdef public bint _lazy
    cdef public object _archive  # ZipFile | None
    cdef dict _num_format_map      # map numberFormat string -> numFmtId
    cdef dict _style_xf_map        # map Style id() -> xfId
    cdef dict _xf_numfmt_map       # map xfId -> numberFormat (при чтении)

    def __init__(self, str sheet_name=None, object _archive=None, bint lazy=False):
        """
        Initializes a new Workbook instance with an empty sheet dictionary,
        shared strings list, and default active sheet.
        """
        self._sheets = {}
        self._shared_strings = []
        self._active_sheet_index = 0
        self._archive = _archive
        self._lazy = lazy
        self._num_format_map = {}
        self._style_xf_map = {}
        self._xf_numfmt_map = {}

        # ВАЖНО: если архив передан, навешиваем на него ссылку на Workbook,
        # чтобы Worksheet через self._archive._workbook_ref мог получить доступ
        # к _xf_numfmt_map (NF-003).
        if self._archive is not None:
            try:
                setattr(self._archive, '_workbook_ref', self)
            except Exception:
                pass

        if self._archive is not None:
            self._load_from_archive(lazy)
        else:
            self._add_sheet(sheet_name)

    cdef void _load_from_archive(self, bint lazy):
        """Оптимизированная загрузка из архива"""
        cdef list sheet_paths
        cdef str sheet_path, name
        cdef object ws

        # На всякий случай ещё раз линкуем архив к книге
        if self._archive is not None:
            try:
                setattr(self._archive, '_workbook_ref', self)
            except Exception:
                pass

        # Загружаем shared strings если есть
        if "xl/sharedStrings.xml" in self._archive.namelist():
            xml = self._archive.read("xl/sharedStrings.xml")
            self._parse_shared_strings(xml)

        # Загружаем стили, если есть
        if "xl/styles.xml" in self._archive.namelist():
            xml = self._archive.read("xl/styles.xml")
            self._parse_styles(xml)

        # --- Новое: читаем реальные имена листов из workbook.xml, если он есть ---
        cdef dict sheetId_to_title = {}
        cdef object root
        cdef object sheets_elem
        cdef object sheet_elem
        cdef str sheet_id_str, title
        if "xl/workbook.xml" in self._archive.namelist():
            try:
                xml = self._archive.read("xl/workbook.xml")
                root = etree.fromstring(xml)
                sheets_elem = root.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheets')
                if sheets_elem is not None:
                    for sheet_elem in sheets_elem.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet'):
                        title = sheet_elem.get('name')
                        sheet_id_str = sheet_elem.get('sheetId')
                        if title is not None and sheet_id_str is not None:
                            sheetId_to_title[int(sheet_id_str)] = title
            except Exception:
                # В случае любой ошибки просто откатимся к именам по умолчанию
                sheetId_to_title = {}

        # Получаем и сортируем пути к листам
        sheet_paths = [
            p for p in self._archive.namelist()
            if p.startswith("xl/worksheets/") and p.endswith(".xml")
        ]
        sheet_paths.sort()

        # Создаем листы
        for idx, sheet_path in enumerate(sheet_paths):
            # sheetId по спецификации начинается с 1
            sheet_id = idx + 1
            name = sheetId_to_title.get(sheet_id, f"sheet{sheet_id}")
            ws = Worksheet(
                self._shared_strings,
                name,
                self._archive,
                sheet_path,
                not lazy,
                xf_numfmt_map=self._xf_numfmt_map,
            )
            self._sheets[name] = ws

    cdef void _add_sheet(self, str sheet_name):
        """Быстрое добавление листа"""
        if sheet_name is None:
            sheet_name = "Sheet"
        cdef object new_sheet = Worksheet(shared_strings=self._shared_strings, title=sheet_name)
        self._sheets[sheet_name] = new_sheet

    def get_sheet_by_name(self, str sheet_name):
        """
        Оптимизированный поиск листа по имени с поддержкой case-insensitive поиска.
        """
        cdef object sheet
        cdef dict lower_sheets
        cdef str sheet_name_lower

        # Сначала прямой поиск
        if sheet_name in self._sheets:
            return self._sheets[sheet_name]

        # Затем case-insensitive
        sheet_name_lower = sheet_name.lower()
        for name, sheet in self._sheets.items():
            if name.lower() == sheet_name_lower:
                return sheet

        raise KeyError(f"No sheet named '{sheet_name}' exists.")

    cpdef void _parse_shared_strings(self, bytes xml_data):
        """
        Оптимизированный парсинг shared strings
        """
        cdef object root = etree.fromstring(xml_data)
        cdef list strings = root.xpath('//si')
        cdef object s
        cdef str text

        # Предварительно выделяем память для списка
        self._shared_strings = [None] * len(strings)

        for i, s in enumerate(strings):
            text = s.xpath('string(.)')
            self._shared_strings[i] = text[0] if text else ""

    cdef void _parse_styles(self, bytes xml_data):
        """Простейший парсинг styles.xml для number_format.

        Строим маппинг xfId -> number_format.
        Поддерживаем только пользовательские numFmts и ссылку numFmtId в cellXfs.
        """
        cdef object root = etree.fromstring(xml_data)
        cdef dict num_fmt_by_id = {}
        cdef object numFmts_elem
        cdef object numFmt
        cdef str numFmtId_str, fmt_code
        cdef int numFmtId

        # Добавляем builtin форматы по умолчанию (включая General)
        num_fmt_by_id.update(BUILTIN_NUMFMTS)

        # Собираем numFmts: numFmtId -> formatCode
        numFmts_elem = root.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}numFmts')
        if numFmts_elem is not None:
            for numFmt in numFmts_elem.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}numFmt'):
                numFmtId_str = numFmt.get('numFmtId')
                fmt_code = numFmt.get('formatCode')
                if numFmtId_str is None or fmt_code is None:
                    continue
                try:
                    numFmtId = int(numFmtId_str)
                except ValueError:
                    continue
                num_fmt_by_id[numFmtId] = fmt_code

        # Теперь cellXfs: индекс xf в списке -> numFmtId -> number_format
        cdef object cellXfs_elem = root.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}cellXfs')
        cdef object xf
        cdef list xfs
        cdef int idx
        cdef str xf_numFmtId_str

        self._xf_numfmt_map.clear()

        if cellXfs_elem is not None:
            xfs = cellXfs_elem.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}xf')
            for idx, xf in enumerate(xfs):
                xf_numFmtId_str = xf.get('numFmtId')
                if xf_numFmtId_str is None:
                    continue
                try:
                    numFmtId = int(xf_numFmtId_str)
                except ValueError:
                    continue
                fmt_code = num_fmt_by_id.get(numFmtId)
                if fmt_code is None and numFmtId in BUILTIN_NUMFMTS:
                    fmt_code = BUILTIN_NUMFMTS[numFmtId]
                if fmt_code is not None:
                    self._xf_numfmt_map[idx] = fmt_code

        print("[DEBUG] _xf_numfmt_map:", self._xf_numfmt_map)

    def create_sheet(self, str title = None):
        """
        Creates a new worksheet with an optional title.
        """
        if title is None:
            title = f"Sheet{len(self._sheets) + 1}"
        cdef object new_sheet = Worksheet(self._shared_strings, title)
        self._sheets[title] = new_sheet
        return new_sheet

    def remove(self, object worksheet):
        """Удалить лист из книги (API, совместимый с openpyxl).

        Принимает объект Worksheet. Если лист не принадлежит этой книге,
        выбрасывает ValueError.
        """
        cdef str key
        cdef int idx
        cdef list names

        # Находим ключ по объекту листа
        for key, ws in self._sheets.items():
            if ws is worksheet:
                # Определяем индекс удаляемого листа
                names = list(self._sheets.keys())
                idx = names.index(key)
                del self._sheets[key]

                # Корректируем active_sheet_index по правилам openpyxl:
                # - если удалённый лист был активным и есть ещё листы,
                #   активным становится предыдущий (или 0, если удалён первый)
                # - если листов больше нет, индекс сбрасываем в 0
                if len(self._sheets) == 0:
                    self._active_sheet_index = 0
                else:
                    if self._active_sheet_index >= idx:
                        if idx == 0:
                            self._active_sheet_index = 0
                        else:
                            self._active_sheet_index = idx - 1
                return

        raise ValueError("The worksheet does not exist in this workbook.")

    @property
    def active(self):
        """
        Returns the active worksheet.
        """
        cdef list sheet_names
        cdef str active_name

        if 0 <= self._active_sheet_index < len(self._sheets):
            sheet_names = list(self._sheets.keys())
            active_name = sheet_names[self._active_sheet_index]
            return self._sheets[active_name]
        raise IndexError("No active sheet available")

    def remove_sheet(self, str title):
        """Удаляет лист по имени (устаревший API; оставлен для обратной совместимости).

        В openpyxl remove_sheet был заменён на remove(worksheet), поэтому тут
        просто ищем лист по имени и делегируем в remove().
        """
        cdef object ws
        if title in self._sheets:
            ws = self._sheets[title]
            self.remove(ws)
        else:
            raise ValueError(f"Worksheet '{title}' does not exist in workbook.")

    @property
    def sheetnames(self):
        """
        List of all sheet names in the workbook.
        """
        return list(self._sheets.keys())

    def set_active_sheet(self, str title):
        """
        Sets the active worksheet by its title.
        """
        cdef list sheet_names
        if title in self._sheets:
            sheet_names = list(self._sheets.keys())
            self._active_sheet_index = sheet_names.index(title)
        else:
            raise KeyError(f"No sheet named '{title}' exists.")

    cdef str _generate_sheet_elements(self):
        """
        Оптимизированная генерация XML <sheet> элементов
        """
        cdef list elements = []
        cdef int i
        cdef object sheet

        for i, sheet in enumerate(self._sheets.values()):
            elements.append(f'<sheet name="{sheet.title}" sheetId="{i + 1}" r:id="rId{i + 1}"/>')

        return "".join(elements)

    cdef str _generate_sheet_overrides(self):
        """
        Оптимизированная генерация XML <Override> элементов
        """
        cdef list overrides = []
        cdef int sheet_count = len(self._sheets)

        for i in range(sheet_count):
            overrides.append(
                f'<Override PartName="/xl/worksheets/sheet{i + 1}.xml" '
                f'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            )

        return "".join(overrides)

    cdef str _generate_relationships(self):
        """
        Оптимизированная генерация XML <Relationship> элементов
        """
        cdef list relationships = []
        cdef int i
        cdef object sheet

        for i, sheet in enumerate(self._sheets.values()):
            relationships.append(
                f'<Relationship Id="rId{i + 1}" '
                f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
                f'Target="worksheets/{sheet.title}.xml"/>'
            )

        return "".join(relationships)

    cdef bytes _get_workbook_xml(self):
        """
        Генерация XML содержимого для workbook.xml
        """
        return WORKBOOK_XML_TEMPLATE.format(sheets=self._generate_sheet_elements()).encode('utf-8')

    cdef bytes _get_content_types_xml(self):
        """
        Генерация XML содержимого для [Content_Types].xml
        """
        return CONTENT_TYPES_XML_TEMPLATE.format(sheet_overrides=self._generate_sheet_overrides()).encode('utf-8')

    cdef bytes _get_styles_xml(self):
        """Генерация minimal styles.xml, совместимого с openpyxl.

        - Если нет пользовательских number_format, numFmts пустой.
        - Всегда есть хотя бы один xf в cellXfs.
        - Всегда есть cellStyleXfs и cellStyles с style "Normal" (xfId=0),
          чтобы openpyxl не выдавал предупреждение
          "Workbook contains no default style, apply openpyxl's default".
        """
        cdef str numFmts_xml, cellXfs_xml
        cdef int numFmts_count, cellXfs_count

        numFmts_xml, cellXfs_xml = self._collect_styles()
        numFmts_count = 0 if not numFmts_xml else len(self._num_format_map)
        cellXfs_count = 1 if not cellXfs_xml else len(self._style_xf_map)

        if not cellXfs_xml:
            # хотя бы один xf по умолчанию (xfId=0)
            cellXfs_xml = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'

        return STYLES_XML_TEMPLATE.format(
            numFmts_count=numFmts_count,
            numFmts=numFmts_xml,
            cellXfs_count=cellXfs_count,
            cellXfs=cellXfs_xml,
        ).encode('utf-8')

    def save_virtual_workbook(self) -> bytes:
        """
        Оптимизированное создание ZIP файла в памяти
        """
        cdef object buffer = BytesIO()
        cdef object sheet

        with ZipFile(buffer, 'w') as zip_file:
            zip_file.writestr("xl/workbook.xml", self._get_workbook_xml())
            zip_file.writestr("[Content_Types].xml", self._get_content_types_xml())
            # styles.xml с number formats
            zip_file.writestr("xl/styles.xml", self._get_styles_xml())

            for sheet in self._sheets.values():
                zip_file.writestr(f"xl/worksheets/{sheet.title}.xml", sheet.get_xml_data())

        buffer.seek(0)
        return buffer.getvalue()

    def save(self, str file_path):
        """
        Оптимизированное сохранение в файл
        """
        if not file_path:
            raise ValueError("File path cannot be empty")

        cdef str dir_path = os.path.dirname(file_path)
        if dir_path:
            os.makedirs(dir_path, exist_ok=True)

        cdef object sheet
        with ZipFile(file_path, 'w') as zip_file:
            zip_file.writestr("xl/workbook.xml", self._get_workbook_xml())
            zip_file.writestr("[Content_Types].xml", self._get_content_types_xml())
            zip_file.writestr("xl/_rels/workbook.xml.rels",
                              RELATIONSHIPS_XML_TEMPLATE.format(relationships=self._generate_relationships()))
            zip_file.writestr("_rels/.rels", MAIN_RELATIONSHIPS_XML_TEMPLATE)
            # styles.xml с number formats
            zip_file.writestr("xl/styles.xml", self._get_styles_xml())

            for sheet in self._sheets.values():
                zip_file.writestr(f"xl/worksheets/{sheet.title}.xml", sheet.get_xml_data())

    @property
    def sheets(self):
        return self._sheets

    def close(self):
        if self._archive is not None:
            self._archive.close()
            self._archive = None

    cdef void _reset_style_caches(self):
        self._num_format_map.clear()
        self._style_xf_map.clear()

    cdef tuple _collect_styles(self):
        """Собирает используемые number_format из всех ячеек и строит таблицы numFmts и cellXfs.

        Возвращает кортеж (numFmts_xml, cellXfs_xml).
        """
        cdef dict num_format_map = {}
        cdef dict style_xf_map = {}
        cdef list numFmt_elems = []
        cdef list xf_elems = []
        cdef int next_numFmtId = 164  # начинаем с диапазона пользовательских форматов
        cdef int next_xfId = 0
        cdef object sheet, cell
        cdef str fmt
        cdef int numFmtId, xfId

        for sheet in self._sheets.values():
            for cell in getattr(sheet, '_cells', {}).values():
                if cell.style is None:
                    continue
                fmt = getattr(cell.style, 'numberFormat', None)
                if not fmt:
                    continue
                # Регистрируем numFmt
                if fmt in num_format_map:
                    numFmtId = num_format_map[fmt]
                else:
                    numFmtId = next_numFmtId
                    next_numFmtId += 1
                    num_format_map[fmt] = numFmtId
                    numFmt_elems.append(f'<numFmt numFmtId="{numFmtId}" formatCode="{fmt}"/>')
                # Регистрируем xf для стиля
                xf_key = id(cell.style)
                if xf_key in style_xf_map:
                    xfId = style_xf_map[xf_key]
                else:
                    xfId = next_xfId
                    next_xfId += 1
                    style_xf_map[xf_key] = xfId
                    xf_elems.append(f'<xf numFmtId="{numFmtId}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>')
                # Сохраняем xfId прямо в объект ячейки для использования при генерации sheet XML
                cell._style_id = xfId

        self._num_format_map = num_format_map
        self._style_xf_map = style_xf_map

        return "".join(numFmt_elems), "".join(xf_elems)

    def __getitem__(self, str key):
        """Доступ к листу по имени через синтаксис wb['Sheet1'], совместимый с openpyxl."""
        return self.get_sheet_by_name(key)
