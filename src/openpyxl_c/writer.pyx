cdef class Writer:
    """Placeholder XLSX writer."""

    def __cinit__(self, workbook, filename: str):
        self.workbook = workbook
        self.filename = filename

    def write(self) -> None:
        raise NotImplementedError

    def _write_workbook(self) -> bytes:
        raise NotImplementedError

    def _write_worksheet(self, ws) -> bytes:
        raise NotImplementedError

    def _write_shared_strings(self) -> bytes:
        raise NotImplementedError

    def _write_styles(self) -> bytes:
        raise NotImplementedError

    def _write_charts(self) -> bytes:
        raise NotImplementedError

    def _write_comments(self) -> bytes:
        raise NotImplementedError

    def _write_data_validations(self) -> bytes:
        raise NotImplementedError

    def _write_conditional_formatting(self) -> bytes:
        raise NotImplementedError
