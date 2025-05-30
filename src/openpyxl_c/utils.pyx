def coordinate_from_string(coord: str):
    """Split an Excel cell reference into column and row indexes."""
    cdef int i = 0
    cdef int col = 0
    cdef int row = 0
    for i in range(len(coord)):
        if coord[i].isdigit():
            break
        col = col * 26 + (ord(coord[i].upper()) - ord('A') + 1)
    if i < len(coord):
        row = int(coord[i:])
    return col, row

def get_column_letter(idx: int) -> str:
    """Return the Excel column letter for a 1-based index."""
    cdef list letters = []
    cdef int num = idx
    while num > 0:
        num -= 1
        letters.append(chr(num % 26 + ord('A')))
        num //= 26
    letters.reverse()
    return "".join(letters)

def range_boundaries(range_string: str):
    """Return numerical boundaries (min_col, min_row, max_col, max_row)."""
    parts = range_string.split(':')
    start = coordinate_from_string(parts[0])
    end = coordinate_from_string(parts[1]) if len(parts) > 1 else start
    return start[0], start[1], end[0], end[1]

def quote_sheetname(name: str) -> str:
    if ' ' in name or any(ch in name for ch in ('"', "'", "[", "]")):
        return f"'{name}'"
    return name

def xml_escape(text: str) -> str:
    return (text.replace('&', '&amp;')
                .replace('<', '&lt;')
                .replace('>', '&gt;')
                .replace('"', '&quot;')
                .replace("'", '&apos;'))
