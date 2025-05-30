from cytosheet.excel import load_workbook as cs_load_workbook


cdef class Reader:
    """Simplified XLSX reader delegating to :func:`cytosheet.excel.load_workbook`."""

    cdef str filename
    cdef bint read_only
    cdef object workbook

    def __cinit__(self, filename: str, bint read_only=False):
        self.filename = filename
        self.read_only = read_only
        self.workbook = cs_load_workbook(filename, read_only)

    def _unpack_archive(self) -> None:
        pass

    def _parse_workbook(self) -> None:
        pass

    def _parse_worksheet(self, sheet_path: str):
        return None

    def _parse_shared_strings(self) -> None:
        pass

    def _parse_styles(self) -> None:
        pass

    def _parse_charts(self) -> None:
        pass

    def _parse_comments(self) -> None:
        pass

    def _parse_data_validations(self) -> None:
        pass

    def _parse_conditional_formatting(self) -> None:
        pass
