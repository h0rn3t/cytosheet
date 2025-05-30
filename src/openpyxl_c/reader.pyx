cdef class Reader:
    """Placeholder XLSX reader."""

    def __cinit__(self, filename: str, bint read_only=False):
        self.filename = filename
        self.read_only = read_only

    def _unpack_archive(self) -> None:
        raise NotImplementedError

    def _parse_workbook(self) -> None:
        raise NotImplementedError

    def _parse_worksheet(self, sheet_path: str):
        raise NotImplementedError

    def _parse_shared_strings(self) -> None:
        raise NotImplementedError

    def _parse_styles(self) -> None:
        raise NotImplementedError

    def _parse_charts(self) -> None:
        raise NotImplementedError

    def _parse_comments(self) -> None:
        raise NotImplementedError

    def _parse_data_validations(self) -> None:
        raise NotImplementedError

    def _parse_conditional_formatting(self) -> None:
        raise NotImplementedError
