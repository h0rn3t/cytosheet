from zipfile import ZipFile

from .workbook import Workbook
from .worksheet import Worksheet


def load_workbook(filename: str, read_only: bool = False):
    """Simplified loader used when Cython extensions are not built."""
    workbook = Workbook()

    with ZipFile(filename, 'r') as archive:
        if 'xl/sharedStrings.xml' in archive.namelist():
            workbook._parse_shared_strings(archive.read('xl/sharedStrings.xml'))

        sheet_files = [f for f in archive.namelist() if f.startswith('xl/worksheets/')]
        for sheet_file in sheet_files:
            xml_data = archive.read(sheet_file)
            sheet_name = sheet_file.split('/')[-1].replace('.xml', '')
            worksheet = Worksheet(workbook._shared_strings, sheet_name)
            worksheet._parse_sheet(xml_data)
            workbook.sheets[sheet_name] = worksheet

        if workbook.sheets:
            workbook._active_sheet_index = 0

    return workbook
