from cytosheet import Workbook as CSWorkbook, load_workbook as cs_load_workbook
from .worksheet import Worksheet


cdef class Workbook:
    """Lightweight wrapper around :mod:`cytosheet` Workbook."""

    cdef CSWorkbook _wb

    def __cinit__(self, bint write_only=False, bint guess_types=False):
        self._wb = CSWorkbook()

    @classmethod
    def load_workbook(cls, filename: str, bint read_only=False):
        cs_wb = cs_load_workbook(filename, read_only=read_only)
        wb = cls()
        wb._wb = cs_wb
        return wb

    def save(self, filename: str, bint as_template=False, bint keep_vba=False) -> None:
        self._wb.save(filename)

    @property
    def active(self):
        return Worksheet._wrap(self, self._wb.active)

    @property
    def sheetnames(self):
        return self._wb.sheetnames

    def create_sheet(self, title: str=None, index: int=None):
        cs_ws = self._wb.create_sheet(title)
        return Worksheet._wrap(self, cs_ws)

    def remove(self, sheet):
        self._wb.remove_sheet(sheet.title)

    def __getitem__(self, key: str):
        try:
            cs_ws = self._wb[key]
            return Worksheet._wrap(self, cs_ws)
        except KeyError:
            raise

    def __iter__(self):
        for name in self.sheetnames:
            yield self[name]
