cdef class XMLNode:
    pass

cdef class Comment:
    def __cinit__(self, text: str, author: str):
        self.text = text
        self.author = author

    def to_tree(self) -> XMLNode:
        return XMLNode()
