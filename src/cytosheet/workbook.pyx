import os
import re
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree
from io import BytesIO
from .worksheet import Worksheet
from .styles import (
    Style,
    font_from_element, fill_from_element, border_from_element,
    alignment_from_element, protection_from_element,
)


def _worksheet_sort_key(path):
    """Числовий ключ для 'xl/worksheets/sheetN.xml', щоб sheet10 йшов після sheet2."""
    cdef str base = path.rsplit('/', 1)[-1]
    cdef str digits = ''.join([ch for ch in base if ch.isdigit()])
    return int(digits) if digits else 0


cdef str _esc_xml_attr(str s):
    """Екранує значення XML-атрибута (&, <, >, ") для генерації styles.xml."""
    return (s.replace('&', '&amp;')
             .replace('<', '&lt;')
             .replace('>', '&gt;')
             .replace('"', '&quot;'))


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
    <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
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
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">{numFmts_section}
    <fonts count="{fonts_count}">{fonts}</fonts>
    <fills count="{fills_count}">{fills}</fills>
    <borders count="{borders_count}">{borders}</borders>
    <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
    <cellXfs count="{cellXfs_count}">{cellXfs}</cellXfs>
    <!-- Минимальный набор для named style "Normal", как ожидает openpyxl -->
    <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""

# Вбудовані формати ECMA-376 (ті самі коди, що в openpyxl.styles.numbers).
# Таблиця має бути ПОВНОЮ: неповна означає number_format=None для файлів, які
# посилаються на пропущений id — а для дат/часу це ще й втрата типу, бо саме
# формат відрізняє дату від числа (напр. id 21 'h:mm:ss' у файлах openpyxl).
cdef dict BUILTIN_NUMFMTS = {
    0: 'General',
    1: '0',
    2: '0.00',
    3: '#,##0',
    4: '#,##0.00',
    5: '"$"#,##0_);("$"#,##0)',
    6: '"$"#,##0_);[Red]("$"#,##0)',
    7: '"$"#,##0.00_);("$"#,##0.00)',
    8: '"$"#,##0.00_);[Red]("$"#,##0.00)',
    9: '0%',
    10: '0.00%',
    11: '0.00E+00',
    12: '# ?/?',
    13: '# ??/??',
    14: 'mm-dd-yy',
    15: 'd-mmm-yy',
    16: 'd-mmm',
    17: 'mmm-yy',
    18: 'h:mm AM/PM',
    19: 'h:mm:ss AM/PM',
    20: 'h:mm',
    21: 'h:mm:ss',
    22: 'm/d/yy h:mm',
    37: '#,##0_);(#,##0)',
    38: '#,##0_);[Red](#,##0)',
    39: '#,##0.00_);(#,##0.00)',
    40: '#,##0.00_);[Red](#,##0.00)',
    41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
    42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
    43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
    44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)',
    45: 'mm:ss',
    46: '[h]:mm:ss',
    47: 'mmss.0',
    48: '##0.0E+0',
    49: '@',
}

cdef class Workbook:
    cdef public dict _sheets
    cdef public list _shared_strings
    cdef public int _active_sheet_index
    cdef public bint _lazy
    cdef public object _archive  # ZipFile | None
    cdef public bint _closed  # книга була завантажена з архіву і архів закрито
    cdef dict _xf_style_map        # map xfId -> повний Style (при чтении styles.xml)
    cdef bytes _original_styles_xml  # оригинальный styles.xml из загруженного файла
    cdef bytes _original_shared_strings_xml  # оригинальный sharedStrings.xml
    cdef bytes _original_content_types_xml  # оригинальный [Content_Types].xml
    cdef bytes _original_workbook_xml  # оригинальный xl/workbook.xml
    cdef bytes _original_workbook_rels_xml  # оригинальный xl/_rels/workbook.xml.rels

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
        self._closed = False
        self._xf_style_map = {}
        self._original_styles_xml = None
        self._original_shared_strings_xml = None
        self._original_content_types_xml = None
        self._original_workbook_xml = None
        self._original_workbook_rels_xml = None

        # ВАЖНО: если архив передан, навешиваем на него ссылку на Workbook,
        # чтобы Worksheet через self._archive._workbook_ref мог получить доступ
        # к _xf_style_map (стилям из styles.xml).
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

        # Сохраняем оригинальный [Content_Types].xml
        if "[Content_Types].xml" in self._archive.namelist():
            self._original_content_types_xml = self._archive.read("[Content_Types].xml")
        
        # Сохраняем оригинальный xl/workbook.xml
        if "xl/workbook.xml" in self._archive.namelist():
            self._original_workbook_xml = self._archive.read("xl/workbook.xml")
        
        # Сохраняем оригинальный xl/_rels/workbook.xml.rels
        if "xl/_rels/workbook.xml.rels" in self._archive.namelist():
            self._original_workbook_rels_xml = self._archive.read("xl/_rels/workbook.xml.rels")

        # Загружаем shared strings если есть
        if "xl/sharedStrings.xml" in self._archive.namelist():
            xml = self._archive.read("xl/sharedStrings.xml")
            self._original_shared_strings_xml = xml  # сохраняем оригинал
            self._parse_shared_strings(xml)

        # Загружаем стили, если есть
        if "xl/styles.xml" in self._archive.namelist():
            xml = self._archive.read("xl/styles.xml")
            self._original_styles_xml = xml  # сохраняем оригинал
            self._parse_styles(xml)

        # --- Листи: порядок беремо з workbook.xml, файли резолвимо через rels (D-5) ---
        # Раніше використовувалося лексикографічне сортування імен файлів +
        # позиційне присвоєння sheetId, через що для файлів з 10+ листами дані
        # "перемішувалися" (sheet10 йшов перед sheet2). Тепер резолвимо коректно:
        # workbook.xml дає порядок і r:id, а workbook.xml.rels — r:id -> файл.
        cdef list ordered_sheets = []      # [(title, rid)] у порядку workbook.xml
        cdef dict rid_to_target = {}        # rId -> Target (відносно xl/)
        cdef object root, sheets_elem, sheet_elem, rels_root, rel
        cdef str title, rid, target, full_path
        cdef str MAIN_NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
        cdef str R_NS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
        cdef bint used_rels = False
        cdef set available
        cdef bint date1904 = False
        cdef object pr_elem

        if self._original_workbook_xml is not None:
            try:
                root = etree.fromstring(self._original_workbook_xml)
                # Система дат книги: 1904 трапляється у файлах зі старих Mac-Excel.
                # Без неї serial->datetime зсувався б рівно на 1462 дні.
                pr_elem = root.find(MAIN_NS + 'workbookPr')
                if pr_elem is not None and pr_elem.get('date1904') in ('1', 'true'):
                    date1904 = True
                sheets_elem = root.find(MAIN_NS + 'sheets')
                if sheets_elem is not None:
                    for sheet_elem in sheets_elem.findall(MAIN_NS + 'sheet'):
                        title = sheet_elem.get('name')
                        rid = sheet_elem.get(R_NS + 'id')
                        if title is not None:
                            ordered_sheets.append((title, rid))
            except Exception:
                ordered_sheets = []

        if self._original_workbook_rels_xml is not None:
            try:
                rels_root = etree.fromstring(self._original_workbook_rels_xml)
                for rel in rels_root:
                    rid = rel.get('Id')
                    target = rel.get('Target')
                    if rid is not None and target is not None:
                        rid_to_target[rid] = target
            except Exception:
                rid_to_target = {}

        available = set(
            p for p in self._archive.namelist()
            if p.startswith("xl/worksheets/") and p.endswith(".xml")
        )

        # Основний шлях: порядок з workbook.xml + резолвинг r:id -> файл через rels
        if ordered_sheets and rid_to_target:
            for title, rid in ordered_sheets:
                target = rid_to_target.get(rid) if rid is not None else None
                if target is None:
                    continue
                if target.startswith('/'):
                    full_path = target.lstrip('/')
                else:
                    full_path = 'xl/' + target
                if full_path not in available:
                    continue
                ws = Worksheet(
                    self._shared_strings, title, self._archive, full_path,
                    not lazy, xf_style_map=self._xf_style_map,
                    date1904=date1904,
                )
                self._sheets[title] = ws
                used_rels = True

        # Fallback (нема/неповні rels): числове сортування файлів + імена за позицією
        if not used_rels:
            sheet_paths = sorted(available, key=_worksheet_sort_key)
            for idx, sheet_path in enumerate(sheet_paths):
                if idx < len(ordered_sheets):
                    name = ordered_sheets[idx][0]
                else:
                    name = f"sheet{idx + 1}"
                ws = Worksheet(
                    self._shared_strings, name, self._archive, sheet_path,
                    not lazy, xf_style_map=self._xf_style_map,
                    date1904=date1904,
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
        cdef list strings
        cdef object s
        cdef object text_result
        
        # Пробуем с namespace
        strings = root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si')
        if not strings:
            # Fallback без namespace
            strings = root.findall('.//si')

        # Предварительно выделяем память для списка
        self._shared_strings = [None] * len(strings)

        for i, s in enumerate(strings):
            text_result = s.xpath('string(.)')
            self._shared_strings[i] = str(text_result) if text_result else ""

    cdef void _parse_styles(self, bytes xml_data):
        """Повний парсинг styles.xml у мапу xfId -> Style (D-1).

        Раніше витягувався лише number_format; тепер для кожного xf збираємо
        повний Style (font/fill/border/alignment/protection/numberFormat) за
        індексами fontId/fillId/borderId/numFmtId, як це робить openpyxl. Це дає
        змогу читати реальні стилі завантаженого файлу через cell.font/fill/...
        """
        cdef str MAIN = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
        cdef object root = etree.fromstring(xml_data)
        cdef dict num_fmt_by_id = {}
        cdef object numFmts_elem, numFmt, fonts_elem, fills_elem, borders_elem
        cdef str numFmtId_str, fmt_code
        cdef int numFmtId

        # Builtin number formats (включаючи General)
        num_fmt_by_id.update(BUILTIN_NUMFMTS)

        # numFmts: numFmtId -> formatCode
        numFmts_elem = root.find(MAIN + 'numFmts')
        if numFmts_elem is not None:
            for numFmt in numFmts_elem.findall(MAIN + 'numFmt'):
                numFmtId_str = numFmt.get('numFmtId')
                fmt_code = numFmt.get('formatCode')
                if numFmtId_str is None or fmt_code is None:
                    continue
                try:
                    numFmtId = int(numFmtId_str)
                except ValueError:
                    continue
                num_fmt_by_id[numFmtId] = fmt_code

        # Індексовані таблиці компонентів: fonts / fills / borders
        cdef list fonts = []
        cdef list fills = []
        cdef list borders = []
        cdef object child

        fonts_elem = root.find(MAIN + 'fonts')
        if fonts_elem is not None:
            for child in fonts_elem.findall(MAIN + 'font'):
                fonts.append(font_from_element(child))

        fills_elem = root.find(MAIN + 'fills')
        if fills_elem is not None:
            for child in fills_elem.findall(MAIN + 'fill'):
                fills.append(fill_from_element(child))

        borders_elem = root.find(MAIN + 'borders')
        if borders_elem is not None:
            for child in borders_elem.findall(MAIN + 'border'):
                borders.append(border_from_element(child))

        # cellXfs: для кожного xf будуємо повний Style за його посиланнями
        cdef object cellXfs_elem = root.find(MAIN + 'cellXfs')
        cdef object xf, align_elem, prot_elem
        cdef list xfs
        cdef int idx, fontId, fillId, borderId
        cdef object style

        self._xf_style_map.clear()

        if cellXfs_elem is None:
            return

        xfs = cellXfs_elem.findall(MAIN + 'xf')
        for idx, xf in enumerate(xfs):
            style = Style()

            fontId = self._safe_int(xf.get('fontId'), -1)
            if 0 <= fontId < len(fonts):
                style.font = fonts[fontId].copy()

            fillId = self._safe_int(xf.get('fillId'), -1)
            if 0 <= fillId < len(fills):
                style.fill = fills[fillId].copy()

            borderId = self._safe_int(xf.get('borderId'), -1)
            if 0 <= borderId < len(borders):
                style.border = borders[borderId].copy()

            numFmtId = self._safe_int(xf.get('numFmtId'), 0)
            fmt_code = num_fmt_by_id.get(numFmtId)
            if fmt_code is not None:
                style.numberFormat = fmt_code

            align_elem = xf.find(MAIN + 'alignment')
            if align_elem is not None:
                style.alignment = alignment_from_element(align_elem)

            prot_elem = xf.find(MAIN + 'protection')
            if prot_elem is not None:
                style.protection = protection_from_element(prot_elem)

            self._xf_style_map[idx] = style

    cdef int _safe_int(self, object value, int default):
        """Безпечний парсинг цілого з XML-атрибута (None/некоректне -> default)."""
        if value is None:
            return default
        try:
            return int(value)
        except (ValueError, TypeError):
            return default

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

        # rId начинается с 2, т.к. rId1 занят styles.xml
        for i, sheet in enumerate(self._sheets.values()):
            elements.append(f'<sheet name="{sheet.title}" sheetId="{i + 1}" r:id="rId{i + 2}"/>')

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

        # Добавляем relationship для styles.xml
        relationships.append(
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
            'Target="styles.xml"/>'
        )

        # Добавляем relationships для листов (начиная с rId2)
        for i, sheet in enumerate(self._sheets.values()):
            relationships.append(
                f'<Relationship Id="rId{i + 2}" '
                f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
                f'Target="worksheets/sheet{i + 1}.xml"/>'
            )

        return "".join(relationships)

    cdef bytes _get_workbook_xml(self):
        """
        Генерация XML содержимого для workbook.xml.
        Если файл был загружен - используем оригинальный.
        """
        if self._original_workbook_xml is not None:
            return self._original_workbook_xml
        return WORKBOOK_XML_TEMPLATE.format(sheets=self._generate_sheet_elements()).encode('utf-8')

    cdef bytes _get_content_types_xml(self):
        """
        Генерация XML содержимого для [Content_Types].xml.
        Если файл был загружен - используем оригинальный.
        """
        if self._original_content_types_xml is not None:
            return self._original_content_types_xml
        return CONTENT_TYPES_XML_TEMPLATE.format(sheet_overrides=self._generate_sheet_overrides()).encode('utf-8')

    cdef bytes _get_styles_xml(self):
        """Генерация styles.xml.

        Для загруженной книги возвращаем оригинальный styles.xml (COMPAT-2: не
        теряем компоненты, которые мы не моделируем — themes, dxfs, indexed colors
        и т.д.), дописывая в конец таблиц лишь НОВЫЕ стили добавленных ячеек
        (merge). Для новой книги строим полный дедуплицированный styles.xml
        (fonts/fills/borders/numFmts + cellXfs со ссылками), чтобы визуальные стили
        (D-2) сохранялись и читались openpyxl.
        """
        if self._original_styles_xml is not None:
            return self._merge_loaded_styles_xml()

        cdef str numFmts_xml, fonts_xml, fills_xml, borders_xml, cellXfs_xml, numFmts_section
        cdef int numFmts_count, fonts_count, fills_count, borders_count, cellXfs_count

        (numFmts_xml, fonts_xml, fills_xml, borders_xml, cellXfs_xml,
         numFmts_count, fonts_count, fills_count, borders_count, cellXfs_count) = self._collect_styles()

        numFmts_section = ''
        if numFmts_count > 0:
            numFmts_section = f'\n    <numFmts count="{numFmts_count}">{numFmts_xml}</numFmts>'

        return STYLES_XML_TEMPLATE.format(
            numFmts_section=numFmts_section,
            fonts=fonts_xml, fonts_count=fonts_count,
            fills=fills_xml, fills_count=fills_count,
            borders=borders_xml, borders_count=borders_count,
            cellXfs=cellXfs_xml, cellXfs_count=cellXfs_count,
        ).encode('utf-8')

    def save_virtual_workbook(self) -> bytes:
        """
        Оптимизированное создание ZIP файла в памяти
        """
        cdef object buffer = BytesIO()
        cdef int i
        cdef object sheet

        self._check_open()

        with ZipFile(buffer, 'w', compression=ZIP_DEFLATED) as zip_file:
            zip_file.writestr("xl/workbook.xml", self._get_workbook_xml())
            zip_file.writestr("[Content_Types].xml", self._get_content_types_xml())
            
            # Используем оригинальный rels если есть
            if self._original_workbook_rels_xml is not None:
                zip_file.writestr("xl/_rels/workbook.xml.rels", self._original_workbook_rels_xml)
            else:
                zip_file.writestr("xl/_rels/workbook.xml.rels",
                                  RELATIONSHIPS_XML_TEMPLATE.format(relationships=self._generate_relationships()))
            
            zip_file.writestr("_rels/.rels", MAIN_RELATIONSHIPS_XML_TEMPLATE)
            # styles.xml с number formats
            zip_file.writestr("xl/styles.xml", self._get_styles_xml())
            
            # Сохраняем sharedStrings.xml если был в оригинале
            if self._original_shared_strings_xml is not None:
                zip_file.writestr("xl/sharedStrings.xml", self._original_shared_strings_xml)

            for i, sheet in enumerate(self._sheets.values()):
                zip_file.writestr(f"xl/worksheets/sheet{i + 1}.xml", sheet.get_xml_data())

        buffer.seek(0)
        return buffer.getvalue()

    def save(self, file_path):
        """
        Оптимизированное сохранение в файл или file-like объект (BytesIO).
        Если файл был загружен из архива - копируем все оригинальные файлы
        и перезаписываем только измененные листы.
        """
        cdef bint is_file_like
        cdef bytes data
        cdef str dir_path
        cdef object sheet

        self._check_open()

        # Поддержка file-like объектов (BytesIO, StringIO и т.д.)
        is_file_like = hasattr(file_path, 'write')
        
        if is_file_like:
            # Сохраняем в BytesIO или другой file-like объект
            data = self.save_virtual_workbook()
            file_path.write(data)
            return
        
        # Обычное сохранение в файл
        if not file_path:
            raise ValueError("File path cannot be empty")

        dir_path = os.path.dirname(file_path)
        if dir_path:
            os.makedirs(dir_path, exist_ok=True)

        with ZipFile(file_path, 'w', compression=ZIP_DEFLATED) as zip_file:
            # Если файл был загружен из архива - копируем все оригинальные файлы
            if self._archive is not None:
                for item in self._archive.namelist():
                    # Пропускаем файлы, которые будем перезаписывать
                    if item in ['xl/workbook.xml', '[Content_Types].xml', 
                                'xl/_rels/workbook.xml.rels', '_rels/.rels',
                                'xl/styles.xml', 'xl/sharedStrings.xml']:
                        continue
                    # Пропускаем листы - их перезапишем
                    if item.startswith('xl/worksheets/') and item.endswith('.xml'):
                        continue
                    # Копируем остальные файлы как есть
                    zip_file.writestr(item, self._archive.read(item))
            
            # Записываем обновленные файлы
            zip_file.writestr("xl/workbook.xml", self._get_workbook_xml())
            zip_file.writestr("[Content_Types].xml", self._get_content_types_xml())
            
            # Используем оригинальный rels если есть
            if self._original_workbook_rels_xml is not None:
                zip_file.writestr("xl/_rels/workbook.xml.rels", self._original_workbook_rels_xml)
            else:
                zip_file.writestr("xl/_rels/workbook.xml.rels",
                                  RELATIONSHIPS_XML_TEMPLATE.format(relationships=self._generate_relationships()))
            
            zip_file.writestr("_rels/.rels", MAIN_RELATIONSHIPS_XML_TEMPLATE)
            zip_file.writestr("xl/styles.xml", self._get_styles_xml())
            
            # Сохраняем sharedStrings.xml если был в оригинале
            if self._original_shared_strings_xml is not None:
                zip_file.writestr("xl/sharedStrings.xml", self._original_shared_strings_xml)

            for i, sheet in enumerate(self._sheets.values()):
                zip_file.writestr(f"xl/worksheets/sheet{i + 1}.xml", sheet.get_xml_data())

    @property
    def sheets(self):
        return self._sheets

    def close(self):
        if self._archive is not None:
            self._archive.close()
            self._archive = None
            # Позначаємо саме факт закриття завантаженої книги: у lazy-режимі
            # оригінальні XML листів у памʼяті не тримаються, тож після close
            # зберегти їх нізвідки — save має впасти, а не писати порожні листи.
            self._closed = True

    cdef _check_open(self):
        """Заборонити збереження книги, архів якої вже закрито (design D7.3)."""
        if self._closed:
            raise ValueError(
                "Cannot save a workbook after close(): the source archive is no "
                "longer available to read unmodified sheets from."
            )

    cdef str _build_xf(self, int numFmtId, int fontId, int fillId, int borderId,
                       str align_xml, str prot_xml):
        """Будує <xf> для cellXfs з реальними посиланнями та apply*-прапорцями."""
        cdef list attrs = [
            f'numFmtId="{numFmtId}"', f'fontId="{fontId}"',
            f'fillId="{fillId}"', f'borderId="{borderId}"', 'xfId="0"',
        ]
        if numFmtId != 0:
            attrs.append('applyNumberFormat="1"')
        if fontId != 0:
            attrs.append('applyFont="1"')
        if fillId != 0:
            attrs.append('applyFill="1"')
        if borderId != 0:
            attrs.append('applyBorder="1"')
        if align_xml:
            attrs.append('applyAlignment="1"')
        if prot_xml:
            attrs.append('applyProtection="1"')
        cdef str head = '<xf ' + ' '.join(attrs)
        if align_xml or prot_xml:
            return head + '>' + align_xml + prot_xml + '</xf>'
        return head + '/>'

    cdef tuple _collect_styles(self):
        """Будує дедуплікований реєстр стилів для НОВОЇ книги (D-2).

        Обходить усі комірки; кожен компонент (font/fill/border/numFmt)
        дедуплікується за його XML-представленням, cellXfs отримує реальні
        посилання fontId/fillId/borderId/numFmtId. У cell._style_id зберігається
        кінцевий xfId (для не-дефолтних стилів). Резервуються індекси 0 (дефолт)
        та fillId 0=none/1=gray125, як вимагає Excel.

        Повертає (numFmts_xml, fonts_xml, fills_xml, borders_xml, cellXfs_xml,
                  numFmts_count, fonts_count, fills_count, borders_count, cellXfs_count).
        """
        cdef dict font_ids = {}
        cdef dict fill_ids = {}
        cdef dict border_ids = {}
        cdef dict numfmt_ids = {}
        cdef dict xf_ids = {}
        cdef list font_elems = []
        cdef list fill_elems = []
        cdef list border_elems = []
        cdef list numfmt_elems = []
        cdef list xf_elems = []
        cdef int next_numFmtId = 164
        cdef object sheet, cell, st
        cdef str font_xml, fill_xml, border_xml, fmt, align_xml, prot_xml, xf_xml
        cdef int fontId, fillId, borderId, numFmtId, xfId

        # Зарезервовані дефолти (індекс 0 кожної таблиці)
        cdef str default_font = '<font><sz val="11"/><name val="Calibri"/></font>'
        cdef str none_fill = '<fill><patternFill patternType="none"/></fill>'
        cdef str gray_fill = '<fill><patternFill patternType="gray125"/></fill>'
        cdef str default_border = '<border><left/><right/><top/><bottom/><diagonal/></border>'
        cdef str default_xf = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'

        font_ids[default_font] = 0
        font_elems.append(default_font)
        fill_ids[none_fill] = 0
        fill_ids[gray_fill] = 1
        fill_elems.append(none_fill)
        fill_elems.append(gray_fill)
        border_ids[default_border] = 0
        border_elems.append(default_border)
        xf_ids[default_xf] = 0
        xf_elems.append(default_xf)

        for sheet in self._sheets.values():
            for cell in getattr(sheet, '_cells', {}).values():
                st = cell.style
                if st is None:
                    cell._style_id = -1
                    continue

                # font
                font_xml = st.font._to_xml() if st.font is not None else default_font
                fontId = font_ids.get(font_xml, -1)
                if fontId == -1:
                    fontId = len(font_elems)
                    font_ids[font_xml] = fontId
                    font_elems.append(font_xml)

                # fill
                fill_xml = st.fill._to_xml() if st.fill is not None else none_fill
                fillId = fill_ids.get(fill_xml, -1)
                if fillId == -1:
                    fillId = len(fill_elems)
                    fill_ids[fill_xml] = fillId
                    fill_elems.append(fill_xml)

                # border
                border_xml = st.border._to_xml() if st.border is not None else default_border
                borderId = border_ids.get(border_xml, -1)
                if borderId == -1:
                    borderId = len(border_elems)
                    border_ids[border_xml] = borderId
                    border_elems.append(border_xml)

                # numFmt (General/None -> 0; інші -> користувацький numFmtId >= 164)
                numFmtId = 0
                fmt = st.numberFormat
                if fmt and fmt != 'General':
                    numFmtId = numfmt_ids.get(fmt, -1)
                    if numFmtId == -1:
                        numFmtId = next_numFmtId
                        next_numFmtId += 1
                        numfmt_ids[fmt] = numFmtId
                        numfmt_elems.append(f'<numFmt numFmtId="{numFmtId}" formatCode="{_esc_xml_attr(fmt)}"/>')

                align_xml = st.alignment._to_xml() if st.alignment is not None else ''
                prot_xml = st.protection._to_xml() if st.protection is not None else ''

                xf_xml = self._build_xf(numFmtId, fontId, fillId, borderId, align_xml, prot_xml)
                xfId = xf_ids.get(xf_xml, -1)
                if xfId == -1:
                    xfId = len(xf_elems)
                    xf_ids[xf_xml] = xfId
                    xf_elems.append(xf_xml)

                # Дефолтний стиль (xfId=0) не вимагає атрибута s= у комірці
                cell._style_id = xfId if xfId > 0 else -1

        return (
            "".join(numfmt_elems), "".join(font_elems), "".join(fill_elems),
            "".join(border_elems), "".join(xf_elems),
            len(numfmt_elems), len(font_elems), len(fill_elems),
            len(border_elems), len(xf_elems),
        )

    # ------------------------------------------------------------------
    # Merge нових стилів у styles.xml завантаженої книги
    # ------------------------------------------------------------------
    # Оригінальний styles.xml зберігається байт-у-байт (COMPAT-2), а стилі
    # доданих/нових комірок (cell._style_id == -1 з не-дефолтним Style)
    # дописуються в кінець відповідних таблиць. Якщо нових стилів немає —
    # повертається оригінал без змін (тож round-trip незмінених книг лишається
    # ідентичним).

    cdef int _section_count(self, str data, str tag):
        """Кількість елементів секції styles.xml за атрибутом count (з fallback)."""
        cdef object m = re.search(r'<' + tag + r'\b[^>]*\bcount="(\d+)"', data)
        if m is not None:
            return int(m.group(1))
        # fallback: рахуємо одиничні елементи
        cdef dict singular = {'fonts': 'font', 'fills': 'fill', 'borders': 'border', 'cellXfs': 'xf'}
        cdef str child = singular.get(tag)
        if child is None:
            return 0
        return len(re.findall(r'<' + child + r'[ >/]', data))

    cdef int _max_numfmt_id(self, str data):
        """Найбільший numFmtId в оригіналі (для безконфліктної нумерації нових)."""
        cdef int mx = 163
        cdef object m
        cdef int v
        for m in re.finditer(r'<numFmt\b[^>]*\bnumFmtId="(\d+)"', data):
            v = int(m.group(1))
            if v > mx:
                mx = v
        return mx

    cdef str _insert_before_close(self, str data, str tag, str new_inner, int new_count):
        """Вставляє new_inner перед </tag> і оновлює count секції."""
        data = re.sub(
            r'(<' + tag + r'\b[^>]*\bcount=")\d+(")',
            lambda m: m.group(1) + str(new_count) + m.group(2),
            data, count=1,
        )
        cdef str close = '</' + tag + '>'
        cdef int idx = data.find(close)
        if idx != -1:
            return data[:idx] + new_inner + data[idx:]
        return data

    cdef str _insert_numfmts(self, str data, str new_inner, int added):
        """Дописує нові numFmt (створює секцію <numFmts>, якщо її немає)."""
        cdef object m
        cdef int old, idx
        if '</numFmts>' in data:
            m = re.search(r'<numFmts\b[^>]*\bcount="(\d+)"', data)
            old = int(m.group(1)) if m is not None else 0
            data = re.sub(
                r'(<numFmts\b[^>]*\bcount=")\d+(")',
                lambda mm: mm.group(1) + str(old + added) + mm.group(2),
                data, count=1,
            )
            idx = data.find('</numFmts>')
            return data[:idx] + new_inner + data[idx:]
        # секції немає — створюємо перед <fonts
        idx = data.find('<fonts')
        if idx != -1:
            return data[:idx] + f'<numFmts count="{added}">{new_inner}</numFmts>' + data[idx:]
        return data

    cdef tuple _collect_new_styles(self, int off_fonts, int off_fills,
                                   int off_borders, int off_xfs, int start_numfmt):
        """Збирає стилі нових/змінених комірок (style_id == -1) з offset-нумерацією.

        «Дефолтними» вважаються і голий openpyxl-дефолт, і стиль xf0 книги — щоб
        завантажені комірки без s=, що успадкували xf0, НЕ дублювали його (інакше
        round-trip незмінених книг переставав би бути ідентичним). Комірка, чий
        стиль резолвиться в дефолтний xf, серіалізується без s= (успадкує xf0).
        Решта отримують нові компоненти, дописані після оригінальних.
        """
        cdef dict nf_font = {}, nf_fill = {}, nf_border = {}, nf_numfmt = {}, nf_xf = {}
        cdef list e_font = [], e_fill = [], e_border = [], e_numfmt = [], e_xf = []
        cdef int next_numfmt = start_numfmt
        cdef object sheet, cell, st
        cdef str fxml, flxml, bxml, fmt, axml, pxml, xfxml
        cdef int fontId, fillId, borderId, numFmtId, xfId

        # Множини дефолтних представлень компонентів: голий дефолт + стиль xf0 книги
        cdef object xf0 = self._xf_style_map.get(0) if self._xf_style_map else None
        cdef set default_fonts = {'<font><sz val="11"/><name val="Calibri"/></font>'}
        cdef set default_fills = {'<fill><patternFill patternType="none"/></fill>'}
        cdef set default_borders = {'<border><left/><right/><top/><bottom/><diagonal/></border>'}
        cdef object default_numfmt = None
        if xf0 is not None:
            if xf0.font is not None:
                default_fonts.add(xf0.font._to_xml())
            if xf0.fill is not None:
                default_fills.add(xf0.fill._to_xml())
            if xf0.border is not None:
                default_borders.add(xf0.border._to_xml())
            default_numfmt = xf0.numberFormat

        for sheet in self._sheets.values():
            if not getattr(sheet, '_modified', False):
                continue
            for cell in getattr(sheet, '_cells', {}).values():
                if cell._style_id != -1:
                    continue  # існуюча незмінена комірка зберігає оригінальний xf
                st = cell.style
                if st is None:
                    continue

                fxml = st.font._to_xml() if st.font is not None else '<font><sz val="11"/><name val="Calibri"/></font>'
                if fxml in default_fonts:
                    fontId = 0
                elif fxml in nf_font:
                    fontId = nf_font[fxml]
                else:
                    fontId = off_fonts + len(e_font)
                    nf_font[fxml] = fontId
                    e_font.append(fxml)

                flxml = st.fill._to_xml() if st.fill is not None else '<fill><patternFill patternType="none"/></fill>'
                if flxml in default_fills:
                    fillId = 0
                elif flxml in nf_fill:
                    fillId = nf_fill[flxml]
                else:
                    fillId = off_fills + len(e_fill)
                    nf_fill[flxml] = fillId
                    e_fill.append(flxml)

                bxml = st.border._to_xml() if st.border is not None else '<border><left/><right/><top/><bottom/><diagonal/></border>'
                if bxml in default_borders:
                    borderId = 0
                elif bxml in nf_border:
                    borderId = nf_border[bxml]
                else:
                    borderId = off_borders + len(e_border)
                    nf_border[bxml] = borderId
                    e_border.append(bxml)

                numFmtId = 0
                fmt = st.numberFormat
                if fmt and fmt != 'General' and fmt != default_numfmt:
                    if fmt in nf_numfmt:
                        numFmtId = nf_numfmt[fmt]
                    else:
                        numFmtId = next_numfmt
                        next_numfmt += 1
                        nf_numfmt[fmt] = numFmtId
                        e_numfmt.append(f'<numFmt numFmtId="{numFmtId}" formatCode="{_esc_xml_attr(fmt)}"/>')

                axml = st.alignment._to_xml() if st.alignment is not None else ''
                pxml = st.protection._to_xml() if st.protection is not None else ''

                # Стиль резолвиться в дефолтний xf0 — не плодимо xf, лишаємо без s=
                if fontId == 0 and fillId == 0 and borderId == 0 and numFmtId == 0 and not axml and not pxml:
                    continue

                xfxml = self._build_xf(numFmtId, fontId, fillId, borderId, axml, pxml)
                if xfxml in nf_xf:
                    xfId = nf_xf[xfxml]
                else:
                    xfId = off_xfs + len(e_xf)
                    nf_xf[xfxml] = xfId
                    e_xf.append(xfxml)
                cell._style_id = xfId

        return (
            "".join(e_numfmt), "".join(e_font), "".join(e_fill),
            "".join(e_border), "".join(e_xf),
            len(e_numfmt), len(e_font), len(e_fill), len(e_border), len(e_xf),
        )

    cdef bytes _merge_loaded_styles_xml(self):
        """styles.xml завантаженої книги: оригінал + дописані нові стилі (або оригінал)."""
        cdef str data = self._original_styles_xml.decode('utf-8')
        cdef int off_fonts = self._section_count(data, 'fonts')
        cdef int off_fills = self._section_count(data, 'fills')
        cdef int off_borders = self._section_count(data, 'borders')
        cdef int off_xfs = self._section_count(data, 'cellXfs')
        cdef int start_numfmt = self._max_numfmt_id(data) + 1
        if start_numfmt < 164:
            start_numfmt = 164

        cdef tuple res = self._collect_new_styles(off_fonts, off_fills, off_borders, off_xfs, start_numfmt)
        cdef str nfmt = res[0], fonts = res[1], fills = res[2], borders = res[3], xfs = res[4]
        cdef int c_nf = res[5], c_f = res[6], c_fl = res[7], c_b = res[8], c_xf = res[9]

        if c_xf == 0:
            # Нових стилів немає — повертаємо оригінал байт-у-байт (COMPAT-2)
            return self._original_styles_xml

        if c_f:
            data = self._insert_before_close(data, 'fonts', fonts, off_fonts + c_f)
        if c_fl:
            data = self._insert_before_close(data, 'fills', fills, off_fills + c_fl)
        if c_b:
            data = self._insert_before_close(data, 'borders', borders, off_borders + c_b)
        data = self._insert_before_close(data, 'cellXfs', xfs, off_xfs + c_xf)
        if c_nf:
            data = self._insert_numfmts(data, nfmt, c_nf)

        return data.encode('utf-8')

    def __getitem__(self, str key):
        """Доступ к листу по имени через синтаксис wb['Sheet1'], совместимый с openpyxl."""
        return self.get_sheet_by_name(key)
