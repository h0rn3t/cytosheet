cdef class Writer:
    """Simplified XLSX writer delegating to :mod:`cytosheet`."""

    def __cinit__(self, workbook, filename: str):
        self.workbook = workbook
        self.filename = filename

    def write(self) -> None:
        self.workbook.save(self.filename)

    def _write_workbook(self) -> bytes:
        return b""

    def _write_worksheet(self, ws) -> bytes:
        return b""

    def _write_shared_strings(self) -> bytes:
        return b""

    def _write_styles(self) -> bytes:
        return b""

    def _write_charts(self) -> bytes:
        return b""

    def _write_comments(self) -> bytes:
        return b""

    def _write_data_validations(self) -> bytes:
        return b""

    def _write_conditional_formatting(self) -> bytes:
        return b""
