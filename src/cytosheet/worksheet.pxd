# worksheet.pxd

from libcpp.set cimport set as cpp_set
from libcpp.string cimport string
from libcpp cimport bool

cimport cython
from .cell cimport Cell

cdef class Worksheet:
    cdef dict _cells
    cdef str title
    cdef list _shared_strings
    cdef dict _data
    cdef set _merged_cells  # native Python set

    cpdef Cell cell(self, int row, int column, object value=*)
    cpdef void _parse_sheet(self, bytes xml_data)
    cpdef append(self, list values)

    cdef str _column_letter(self, int idx)
    cdef int _column_index_from_key(self, str key)
    cdef int _row_index_from_key(self, str key)
    cdef list get_row_cells(self, int row, int min_col, int max_col)
    cdef list get_col_cells(self, int col, int min_row, int max_row)
