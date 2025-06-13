cdef class Workbook:
    cdef dict _sheets
    cdef list _shared_strings
    cdef int _active_sheet_index

    cpdef void _parse_shared_strings(self, bytes xml_data)
    cdef bytes _get_workbook_xml(self)
    cdef bytes _get_content_types_xml(self)
    cdef str _generate_sheet_elements(self)
    cdef str _generate_sheet_overrides(self)
    cdef str _generate_relationships(self)