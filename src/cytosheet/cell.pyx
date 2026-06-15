cdef class Cell:
    cdef str position
    cdef public object value
    cdef public object style
    cdef public object parent
    cdef public bint is_merged_cell
    cdef public str merged_range
    cdef public str data_type  # 'n', 's', 'b', 'f', 'd', 'e'
    cdef public int _style_id  # внутренний индекс XF для styles.xml
    cdef public int _shared_string_index  # индекс в sharedStrings для оригинальных ячеек

    def __init__(self, str position=None, object value=None, object style=None, object parent=None):
        self.position = position if position is not None else ""
        self.value = value
        self.parent = parent
        self.is_merged_cell = False
        self.merged_range = None
        self.data_type = None
        self._style_id = -1
        self._shared_string_index = -1  # -1 означает что это не ссылка на sharedString
        # Кожна комірка отримує ВЛАСНИЙ Style, щоб стилі не "протікали" між
        # комірками й книгами (раніше спільний DEFAULT_STYLE мутувався глобально).
        if style is None:
            from .styles import Style
            self.style = Style()
        else:
            self.style = style

    def __repr__(self):
        if self.is_merged_cell:
            return f"<MergedCell position={self.position}, merged_range={self.merged_range}>"
        # Показываем data_type для отладки, если он задан
        if self.data_type is not None:
            return f"<Cell position={self.position}, value={self.value}, data_type={self.data_type}>"
        return f"<Cell position={self.position}, value={self.value}>"

    cpdef str get_position(self):
        """Быстрый доступ к позиции ячейки"""
        return self.position

    cpdef void set_value(self, object value):
        """Быстрая установка значения с обновлением типа данных для формул."""
        self.value = value
        # Значення змінено вручну — комірка більше НЕ посилається на вихідний
        # sharedString, інакше серіалізатор повторно віддав би старий індекс (D-4).
        self._shared_string_index = -1
        # Простая эвристика: если это строка, начинающаяся с '=', считаем формулой
        if isinstance(value, str) and value.startswith('='):
            self.data_type = 'f'
        else:
            # type inference для других случаев можно доработать позже
            pass

    cpdef object get_value(self):
        """Быстрый доступ к значению"""
        return self.value

    cpdef void set_style(self, object style):
        """Быстрая установка стиля"""
        self.style = style

    cpdef object get_style(self):
        """Быстрый доступ к стилю"""
        return self.style

    # ------------------------------------------------------------------
    # Свойства, совместимые с openpyxl
    # ------------------------------------------------------------------

    @property
    def coordinate(self):
        """Координата ячейки (A1), алиас для position."""
        return self.position

    @property
    def row(self):
        """Номер строки (1-based)."""
        cdef int i = 0
        cdef int n = len(self.position)
        while i < n and self.position[i].isalpha():
            i += 1
        if i == 0 or i == n:
            return 0
        return int(self.position[i:])

    @property
    def column(self):
        """Номер колонки (1-based)."""
        cdef int i = 0
        cdef int n = len(self.position)
        cdef str col_letters
        while i < n and self.position[i].isalpha():
            i += 1
        if i == 0:
            return 0
        col_letters = self.position[:i]
        # Воспользуемся утилитой из worksheet через parent, если доступна
        if self.parent is not None and hasattr(self.parent, '_col_to_num'):
            return self.parent._col_to_num(col_letters)
        # Fallback: собственная реализация
        cdef int result = 0
        cdef object ch
        for ch in col_letters:
            result = result * 26 + (ord((<str>ch).upper()) - ord('A') + 1)
        return result

    @property
    def column_letter(self):
        """Буквенное обозначение колонки (A, B, ...)."""
        # Если есть parent с _num_to_col, можно вычислить из column
        if self.parent is not None and hasattr(self.parent, '_num_to_col'):
            return self.parent._num_to_col(self.column)
        # Простейший fallback
        cdef int n = self.column
        cdef list letters = []
        cdef int rem
        if n <= 0:
            return ""
        while n > 0:
            n, rem = divmod(n - 1, 26)
            letters.append(chr(ord('A') + rem))
        letters.reverse()
        return "".join(letters)

    cdef void _mark_modified(self):
        """Позначити зміну стилю комірки: лист змінено (D-3) + xf застарів.

        Скидаємо _style_id у -1, щоб серіалізатор/merge перебудував xf за поточним
        (повним) Style комірки. Інакше зміна стилю вже-завантаженої комірки
        (що мала s=) загубилась би — емітився б старий xfId.
        """
        self._style_id = -1
        if self.parent is not None:
            try:
                self.parent._modified = True
            except Exception:
                pass

    # Прокси к style, совместимые с openpyxl

    @property
    def font(self):
        return self.style.font if self.style is not None else None

    @font.setter
    def font(self, value):
        # Стиль всегда существует, просто прокидываем ссылку
        self.style.font = value
        self._mark_modified()

    @property
    def border(self):
        return self.style.border if self.style is not None else None

    @border.setter
    def border(self, value):
        self.style.border = value
        self._mark_modified()

    @property
    def fill(self):
        return self.style.fill if self.style is not None else None

    @fill.setter
    def fill(self, value):
        self.style.fill = value
        self._mark_modified()

    @property
    def alignment(self):
        return self.style.alignment if self.style is not None else None

    @alignment.setter
    def alignment(self, value):
        self.style.alignment = value
        self._mark_modified()

    @property
    def protection(self):
        return self.style.protection if self.style is not None else None

    @protection.setter
    def protection(self, value):
        self.style.protection = value
        self._mark_modified()

    @property
    def number_format(self):
        if self.style is None:
            return None
        return self.style.numberFormat

    @number_format.setter
    def number_format(self, fmt):
        if self.style is None:
            from .styles import Style
            self.style = Style()
        self.style.numberFormat = fmt
        self._mark_modified()
