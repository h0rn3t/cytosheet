import os
from io import BytesIO
from zipfile import ZipFile

from lxml import etree

from .worksheet import Worksheet

WORKBOOK_XML_TEMPLATE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>{sheets}</sheets>
</workbook>"""

CONTENT_TYPES_XML_TEMPLATE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    {sheet_overrides}
</Types>"""

RELATIONSHIPS_XML_TEMPLATE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    {relationships}
</Relationships>"""

MAIN_RELATIONSHIPS_XML_TEMPLATE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""


class Workbook:
    def __init__(self, sheet_name=None):
        self._sheets = {}
        self._shared_strings = []
        self._active_sheet_index = 0
        self._add_sheet(sheet_name)

    def _add_sheet(self, sheet_name=None):
        if sheet_name is None:
            sheet_name = "Sheet"
        new_sheet = Worksheet(shared_strings=self._shared_strings, title=sheet_name)
        self._sheets[sheet_name] = new_sheet

    def get_sheet_by_name(self, sheet_name):
        if sheet_name in self._sheets:
            return self._sheets[sheet_name]
        lower_sheets = {name.lower(): ws for name, ws in self._sheets.items()}
        sheet_name_lower = sheet_name.lower()
        if sheet_name_lower in lower_sheets:
            return lower_sheets[sheet_name_lower]
        raise KeyError(f"No sheet named '{sheet_name}' exists.")

    def __getitem__(self, key):
        return self.get_sheet_by_name(key)

    @property
    def worksheets(self):
        return list(self._sheets.values())

    def close(self):
        self._sheets.clear()
        self._shared_strings.clear()

    def _parse_shared_strings(self, xml_data: bytes):
        root = etree.fromstring(xml_data)
        strings = root.xpath('//si')
        for s in strings:
            text = s.xpath('string(.)')[0]
            self._shared_strings.append(text)

    def create_sheet(self, title=None):
        if title is None:
            title = f"Sheet{len(self._sheets) + 1}"
        new_sheet = Worksheet(self._shared_strings, title)
        self._sheets[title] = new_sheet
        return new_sheet

    @property
    def active(self):
        if 0 <= self._active_sheet_index < len(self._sheets):
            active_name = list(self._sheets.keys())[self._active_sheet_index]
            return self._sheets[active_name]
        raise IndexError("No active sheet available")

    def remove_sheet(self, title):
        if title in self._sheets:
            del self._sheets[title]
        else:
            raise ValueError(f"Worksheet '{title}' does not exist in workbook.")

    @property
    def sheetnames(self):
        return list(self._sheets.keys())

    def set_active_sheet(self, title):
        if title in self._sheets:
            self._active_sheet_index = list(self._sheets.keys()).index(title)
        else:
            raise KeyError(f"No sheet named '{title}' exists.")

    def _generate_sheet_elements(self):
        return "".join(
            f'<sheet name="{sheet.title}" sheetId="{i + 1}" r:id="rId{i + 1}"/>'
            for i, sheet in enumerate(self._sheets.values())
        )

    def _generate_sheet_overrides(self):
        return "".join(
            f'<Override PartName="/xl/worksheets/sheet{i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            for i in range(len(self._sheets.values()))
        )

    def _generate_relationships(self):
        return "".join(
            f'<Relationship Id="rId{i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/{sheet.title}.xml"/>'
            for i, sheet in enumerate(self._sheets.values())
        )

    def _get_workbook_xml(self):
        return WORKBOOK_XML_TEMPLATE.format(sheets=self._generate_sheet_elements()).encode('utf-8')

    def _get_content_types_xml(self):
        return CONTENT_TYPES_XML_TEMPLATE.format(sheet_overrides=self._generate_sheet_overrides()).encode('utf-8')

    def save_virtual_workbook(self) -> bytes:
        buffer = BytesIO()
        with ZipFile(buffer, 'w') as zip_file:
            zip_file.writestr("xl/workbook.xml", self._get_workbook_xml())
            zip_file.writestr("[Content_Types].xml", self._get_content_types_xml())
            for sheet in self._sheets.values():
                zip_file.writestr(f"xl/worksheets/{sheet.title}.xml", sheet.get_xml_data())
        buffer.seek(0)
        return buffer.getvalue()

    def save(self, file_path: str):
        if not file_path:
            raise ValueError("File path cannot be empty")
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        with ZipFile(file_path, 'w') as zip_file:
            zip_file.writestr("xl/workbook.xml", self._get_workbook_xml())
            zip_file.writestr("[Content_Types].xml", self._get_content_types_xml())
            zip_file.writestr(
                "xl/_rels/workbook.xml.rels",
                RELATIONSHIPS_XML_TEMPLATE.format(relationships=self._generate_relationships()),
            )
            zip_file.writestr("_rels/.rels", MAIN_RELATIONSHIPS_XML_TEMPLATE)
            for sheet in self._sheets.values():
                zip_file.writestr(f"xl/worksheets/{sheet.title}.xml", sheet.get_xml_data())

    @property
    def sheets(self):
        return self._sheets
