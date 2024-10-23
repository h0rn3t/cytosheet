from zipfile import ZipFile
from lxml import etree
from .worksheet import Worksheet

cdef class Workbook:
    cdef public dict _sheets
    cdef public list _shared_strings

    def __init__(self):
        self._sheets = {}
        self._shared_strings = []

    def load_workbook(self, filename: str):
        with ZipFile(filename, 'r') as archive:
            if 'xl/sharedStrings.xml' in archive.namelist():
                self._parse_shared_strings(archive.read('xl/sharedStrings.xml'))

            # Затем парсим листы
            sheet_files = [f for f in archive.namelist() if f.startswith("xl/worksheets/sheet")]
            for sheet_file in sheet_files:
                xml_data = archive.read(sheet_file)
                sheet_name = sheet_file.split('/')[-1].replace('.xml', '')  # Название листа
                worksheet = Worksheet(sheet_name, self._shared_strings)
                worksheet._parse_sheet(xml_data)
                self._sheets[sheet_name] = worksheet

    def get_sheet_by_name(self, sheet_name: str):
        if sheet_name in self._sheets:
            return self._sheets[sheet_name]
        lower_sheets = {name.lower(): ws for name, ws in self._sheets.items()}
        sheet_name_lower = sheet_name.lower()
        if sheet_name_lower in lower_sheets:
            return lower_sheets[sheet_name_lower]
        raise KeyError(f"No sheet named '{sheet_name}' exists.")

    cdef void _parse_shared_strings(self, bytes xml_data):
        """parse sharedStrings.xml"""
        root = etree.fromstring(xml_data)
        strings = root.xpath('//si')
        for s in strings:
            self._shared_strings.append(s.xpath('string(.)')[0])

