import os
from zipfile import ZipFile
from lxml import etree
from io import BytesIO
from .worksheet import Worksheet

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

cdef class Workbook:
    cdef public dict _sheets
    cdef public list _shared_strings
    cdef public int _active_sheet_index
    cdef public bint _lazy
    cdef public object _archive  # ZipFile | None

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

        if self._archive is not None:
            self._load_from_archive(lazy)
        else:
            self._add_sheet(sheet_name)

    cdef void _load_from_archive(self, bint lazy):
        """Оптимизированная загрузка из архива"""
        cdef list sheet_paths
        cdef str sheet_path, name
        cdef object ws

        # Загружаем shared strings если есть
        if "xl/sharedStrings.xml" in self._archive.namelist():
            xml = self._archive.read("xl/sharedStrings.xml")
            self._parse_shared_strings(xml)

        # Получаем и сортируем пути к листам
        sheet_paths = [
            p for p in self._archive.namelist()
            if p.startswith("xl/worksheets/") and p.endswith(".xml")
        ]
        sheet_paths.sort()

        # Создаем листы
        for sheet_path in sheet_paths:
            name = sheet_path.rsplit("/", 1)[-1].replace(".xml", "")
            ws = Worksheet(
                self._shared_strings,
                name,
                self._archive,
                sheet_path,
                not lazy
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

    def create_sheet(self, str title = None):
        """
        Creates a new worksheet with an optional title.
        """
        if title is None:
            title = f"Sheet{len(self._sheets) + 1}"
        cdef object new_sheet = Worksheet(self._shared_strings, title)
        self._sheets[title] = new_sheet
        return new_sheet

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
        """
        Removes a worksheet from the workbook by title.
        """
        if title in self._sheets:
            del self._sheets[title]
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

    def save_virtual_workbook(self) -> bytes:
        """
        Оптимизированное создание ZIP файла в памяти
        """
        cdef object buffer = BytesIO()
        cdef object sheet

        with ZipFile(buffer, 'w') as zip_file:
            zip_file.writestr("xl/workbook.xml", self._get_workbook_xml())
            zip_file.writestr("[Content_Types].xml", self._get_content_types_xml())

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

            for sheet in self._sheets.values():
                zip_file.writestr(f"xl/worksheets/{sheet.title}.xml", sheet.get_xml_data())

    @property
    def sheets(self):
        return self._sheets

    def close(self):
        if self._archive is not None:
            self._archive.close()
            self._archive = None