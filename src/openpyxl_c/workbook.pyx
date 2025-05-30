cdef class Workbook:
    """Placeholder Workbook class emulating a subset of the OpenPyXL API."""

    def __cinit__(self, bint write_only=False, bint guess_types=False):
        self._sheets = []
        self._active = None

    @classmethod
    def load_workbook(cls, filename: str, bint read_only=False):
        raise NotImplementedError("Workbook.load_workbook is not implemented")

    def save(self, filename: str, bint as_template=False, bint keep_vba=False) -> None:
        raise NotImplementedError("Workbook.save is not implemented")

    @property
    def active(self):
        return self._active

    @property
    def sheetnames(self):
        return [getattr(ws, "title", "Sheet") for ws in self._sheets]

    def create_sheet(self, title: str=None, index: int=None):
        raise NotImplementedError("Workbook.create_sheet is not implemented")

    def remove(self, sheet):
        raise NotImplementedError("Workbook.remove is not implemented")

    def __getitem__(self, key: str):
        raise KeyError(key)

    def __iter__(self):
        for ws in self._sheets:
            yield ws
