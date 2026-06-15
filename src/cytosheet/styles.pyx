from lxml import etree


# ------------------------------------------------------------------
# Допоміжні функції рівня модуля (екранування, читання атрибутів)
# ------------------------------------------------------------------

cdef str _esc_attr(str s):
    """Екранує значення XML-атрибута (&, <, >, ")."""
    return (s.replace('&', '&amp;')
             .replace('<', '&lt;')
             .replace('>', '&gt;')
             .replace('"', '&quot;'))


cdef str _localname(object elem):
    """Локальне імʼя тега без namespace ('{ns}font' -> 'font')."""
    cdef str tag = elem.tag
    if not isinstance(tag, str):
        return ''
    if '}' in tag:
        return tag.split('}', 1)[1]
    return tag


cdef bint _flag_child(object child):
    """Булевий прапорець-елемент шрифту (<b/> = True, <b val="0"/> = False)."""
    cdef object v = child.get('val')
    if v is None:
        return True
    return v not in ('0', 'false', 'False')


cdef bint _attr_true(object v):
    """Булевий XML-атрибут ('1'/'true' -> True)."""
    return v in ('1', 'true', 'True')


cdef str _fmt_size(object size):
    """Форматує розмір шрифту для XML: None/11.0 -> '11', 11.5 -> '11.5'."""
    if size is None:
        return '11'
    cdef double s = float(size)
    if s == int(s):
        return str(int(s))
    return str(s)


cdef class Color:
    """
    Represents a color in ARGB format.
    """
    cdef public str rgb
    cdef public str indexed
    cdef public str auto_color  # renamed from 'auto' which is a C++ keyword
    cdef public str theme
    cdef public float tint

    def __init__(self, rgb=None, indexed=None, auto=None, theme=None, tint=0.0):
        self.rgb = rgb
        self.indexed = indexed
        self.auto_color = auto  # renamed from 'auto' which is a C++ keyword
        self.theme = theme
        self.tint = tint

    def __repr__(self):
        if self.rgb:
            return f"<Color rgb={self.rgb}>"
        elif self.indexed:
            return f"<Color indexed={self.indexed}>"
        elif self.theme:
            return f"<Color theme={self.theme} tint={self.tint}>"
        elif self.auto_color:
            return f"<Color auto={self.auto_color}>"
        return "<Color>"

    cpdef Color copy(self):
        """Повертає незалежну копію кольору (для ізоляції стилю комірки, D-6)."""
        return Color(rgb=self.rgb, indexed=self.indexed, auto=self.auto_color,
                     theme=self.theme, tint=self.tint)

    cpdef str _attr_str(self):
        """Атрибути для <color>/<fgColor>/<bgColor>; '' якщо колір не задано.

        Пріоритет як в OOXML: rgb > theme(+tint) > indexed.
        """
        if self.rgb:
            return f'rgb="{self.rgb}"'
        if self.theme is not None and self.theme != '':
            if self.tint:
                return f'theme="{self.theme}" tint="{self.tint}"'
            return f'theme="{self.theme}"'
        if self.indexed:
            return f'indexed="{self.indexed}"'
        return ''


cdef class Side:
    """
    Border options for use in styles.
    """
    cdef public str style
    cdef public Color color

    def __init__(self, style=None, color=None):
        self.style = style
        self.color = color if color is not None else Color()

    def __repr__(self):
        return f"<Side style={self.style}>"

    cpdef Side copy(self):
        return Side(style=self.style,
                    color=self.color.copy() if self.color is not None else None)

    cpdef str _to_xml(self, str tag):
        """Серіалізує сторону рамки як <left/.../> з опційним кольором."""
        cdef str ca
        if not self.style:
            return f'<{tag}/>'
        ca = self.color._attr_str() if self.color is not None else ''
        if ca:
            return f'<{tag} style="{self.style}"><color {ca}/></{tag}>'
        return f'<{tag} style="{self.style}"/>'


cdef class Border:
    """
    Border positioning for use in styles.
    """
    cdef public Side left
    cdef public Side right
    cdef public Side top
    cdef public Side bottom
    cdef public Side diagonal
    cdef public bint diagonalUp
    cdef public bint diagonalDown
    cdef public bint outline

    def __init__(self, left=None, right=None, top=None, bottom=None, diagonal=None,
                 diagonalUp=False, diagonalDown=False, outline=True):
        self.left = left if left is not None else Side()
        self.right = right if right is not None else Side()
        self.top = top if top is not None else Side()
        self.bottom = bottom if bottom is not None else Side()
        self.diagonal = diagonal if diagonal is not None else Side()
        self.diagonalUp = diagonalUp
        self.diagonalDown = diagonalDown
        self.outline = outline

    def __repr__(self):
        return f"<Border left={self.left} right={self.right} top={self.top} bottom={self.bottom}>"

    def __add__(self, other):
        """
        Add borders together, with the other border taking precedence.
        """
        cdef Side left, right, top, bottom, diagonal

        if not isinstance(other, Border):
            return self

        left = other.left if other.left.style else self.left
        right = other.right if other.right.style else self.right
        top = other.top if other.top.style else self.top
        bottom = other.bottom if other.bottom.style else self.bottom
        diagonal = other.diagonal if other.diagonal.style else self.diagonal

        return Border(
            left=left, right=right, top=top, bottom=bottom, diagonal=diagonal,
            diagonalUp=other.diagonalUp or self.diagonalUp,
            diagonalDown=other.diagonalDown or self.diagonalDown,
            outline=other.outline
        )

    cpdef Border copy(self):
        return Border(
            left=self.left.copy() if self.left is not None else None,
            right=self.right.copy() if self.right is not None else None,
            top=self.top.copy() if self.top is not None else None,
            bottom=self.bottom.copy() if self.bottom is not None else None,
            diagonal=self.diagonal.copy() if self.diagonal is not None else None,
            diagonalUp=self.diagonalUp, diagonalDown=self.diagonalDown,
            outline=self.outline,
        )

    cpdef str _to_xml(self):
        """Серіалізує рамку; дефолтна -> <border><left/>...<diagonal/></border>."""
        return ('<border>'
                + self.left._to_xml('left')
                + self.right._to_xml('right')
                + self.top._to_xml('top')
                + self.bottom._to_xml('bottom')
                + self.diagonal._to_xml('diagonal')
                + '</border>')


cdef class Font:
    """
    Font options for use in styles.
    """
    cdef public str name
    cdef public object size  # Changed from float to object to allow None
    cdef public bint bold
    cdef public bint italic
    cdef public bint underline
    cdef public bint strike
    cdef public Color color

    def __init__(self, name='Calibri', size=11.0, bold=False, italic=False, underline=False,
                 strike=False, color=None):
        # Дефолти збігаються з openpyxl Font() (Calibri, 11), щоб комірки без
        # явного шрифту (зокрема у файлах без styles.xml) читалися однаково.
        self.name = name
        self.size = size  # Can be None or a float
        self.bold = bold
        self.italic = italic
        self.underline = underline
        self.strike = strike
        self.color = color if color is not None else Color()

    def __repr__(self):
        return f"<Font name={self.name} size={self.size} bold={self.bold}>"

    @property
    def sz(self):
        """Алиас для size (openpyxl-сумісність: Font.sz)."""
        return self.size

    @sz.setter
    def sz(self, value):
        self.size = value

    cpdef Font copy(self):
        return Font(name=self.name, size=self.size, bold=self.bold,
                    italic=self.italic, underline=self.underline,
                    strike=self.strike,
                    color=self.color.copy() if self.color is not None else None)

    cpdef str _to_xml(self):
        """Серіалізує шрифт; дефолтний -> <font><sz val="11"/><name val="Calibri"/></font>.

        Дефолти (name=None -> Calibri, size=None -> 11) узгоджені з fontId=0,
        тож комірки без явного шрифту дедуплікуються в дефолтний шрифт.
        """
        parts = ['<font>']
        if self.bold:
            parts.append('<b/>')
        if self.italic:
            parts.append('<i/>')
        if self.underline:
            parts.append('<u/>')
        if self.strike:
            parts.append('<strike/>')
        parts.append(f'<sz val="{_fmt_size(self.size)}"/>')
        ca = self.color._attr_str() if self.color is not None else ''
        if ca:
            parts.append(f'<color {ca}/>')
        parts.append(f'<name val="{_esc_attr(self.name) if self.name else "Calibri"}"/>')
        parts.append('</font>')
        return ''.join(parts)


cdef class PatternFill:
    """
    Pattern fill for use in styles.
    """
    cdef public str patternType
    cdef public Color fgColor
    cdef public Color bgColor

    def __init__(self, patternType=None, fgColor=None, bgColor=None):
        self.patternType = patternType
        self.fgColor = fgColor if fgColor is not None else Color()
        self.bgColor = bgColor if bgColor is not None else Color()

    def __repr__(self):
        return f"<PatternFill patternType={self.patternType}>"

    cpdef PatternFill copy(self):
        return PatternFill(
            patternType=self.patternType,
            fgColor=self.fgColor.copy() if self.fgColor is not None else None,
            bgColor=self.bgColor.copy() if self.bgColor is not None else None,
        )

    cpdef str _to_xml(self):
        """Серіалізує заливку; дефолтна -> <fill><patternFill patternType="none"/></fill>."""
        cdef str pt = self.patternType
        if not pt or pt == 'none':
            return '<fill><patternFill patternType="none"/></fill>'
        if pt == 'gray125':
            return '<fill><patternFill patternType="gray125"/></fill>'
        fg = self.fgColor._attr_str() if self.fgColor is not None else ''
        bg = self.bgColor._attr_str() if self.bgColor is not None else ''
        inner = f'<patternFill patternType="{pt}">'
        if fg:
            inner += f'<fgColor {fg}/>'
        if bg:
            inner += f'<bgColor {bg}/>'
        inner += '</patternFill>'
        return f'<fill>{inner}</fill>'


cdef class Alignment:
    """Alignment options for use in styles (openpyxl-совместный API).

    Публичные аргументы совпадают с openpyxl:
    - horizontal
    - vertical
    - text_rotation
    - wrap_text
    - shrink_to_fit
    - indent
    """
    cdef public str horizontal
    cdef public str vertical
    cdef public int textRotation
    cdef public bint wrapText
    cdef public bint shrinkToFit
    cdef public int indent

    def __init__(
        self,
        horizontal=None,
        vertical=None,
        text_rotation=None,
        wrap_text=None,
        shrink_to_fit=None,
        indent=0,
    ):
        self.horizontal = horizontal
        self.vertical = vertical

        if text_rotation is None:
            self.textRotation = 0
        else:
            self.textRotation = int(text_rotation)

        self.wrapText = False if wrap_text is None else bool(wrap_text)
        self.shrinkToFit = False if shrink_to_fit is None else bool(shrink_to_fit)
        self.indent = int(indent)

    def __repr__(self):
        return f"<Alignment horizontal={self.horizontal} vertical={self.vertical}>"

    cpdef Alignment copy(self):
        return Alignment(
            horizontal=self.horizontal, vertical=self.vertical,
            text_rotation=self.textRotation, wrap_text=self.wrapText,
            shrink_to_fit=self.shrinkToFit, indent=self.indent,
        )

    cpdef str _to_xml(self):
        """Серіалізує вирівнювання; '' якщо дефолтне (нічого не задано)."""
        attrs = []
        if self.horizontal:
            attrs.append(f'horizontal="{self.horizontal}"')
        if self.vertical:
            attrs.append(f'vertical="{self.vertical}"')
        if self.textRotation:
            attrs.append(f'textRotation="{self.textRotation}"')
        if self.wrapText:
            attrs.append('wrapText="1"')
        if self.shrinkToFit:
            attrs.append('shrinkToFit="1"')
        if self.indent:
            attrs.append(f'indent="{self.indent}"')
        if not attrs:
            return ''
        return '<alignment ' + ' '.join(attrs) + '/>'


cdef class Protection:
    """
    Protection options for use in styles.
    """
    cdef public bint locked
    cdef public bint hidden

    def __init__(self, locked=True, hidden=False):
        self.locked = locked
        self.hidden = hidden

    def __repr__(self):
        return f"<Protection locked={self.locked} hidden={self.hidden}>"

    cpdef Protection copy(self):
        return Protection(locked=self.locked, hidden=self.hidden)

    cpdef str _to_xml(self):
        """Серіалізує захист; '' для дефолту (locked=True, hidden=False)."""
        if self.locked and not self.hidden:
            return ''
        return f'<protection locked="{1 if self.locked else 0}" hidden="{1 if self.hidden else 0}"/>'


cdef class Style:
    """
    Style for use in cells.
    """
    cdef public Font font
    cdef public Border border
    cdef public PatternFill fill
    cdef public Alignment alignment
    cdef public Protection protection
    cdef public str numberFormat

    def __init__(self, font=None, border=None, fill=None, alignment=None, protection=None, numberFormat=None):
        self.font = font if font is not None else Font()
        self.border = border if border is not None else Border()
        self.fill = fill if fill is not None else PatternFill()
        self.alignment = alignment if alignment is not None else Alignment()
        self.protection = protection if protection is not None else Protection()
        self.numberFormat = numberFormat

    def __repr__(self):
        return f"<Style font={self.font} border={self.border} alignment={self.alignment}>"

    cpdef Style copy(self):
        """Глибока копія стилю: кожна комірка має власний Style, без протікання (D-6)."""
        return Style(
            font=self.font.copy() if self.font is not None else None,
            border=self.border.copy() if self.border is not None else None,
            fill=self.fill.copy() if self.fill is not None else None,
            alignment=self.alignment.copy() if self.alignment is not None else None,
            protection=self.protection.copy() if self.protection is not None else None,
            numberFormat=self.numberFormat,
        )


# ------------------------------------------------------------------
# Парсинг компонентів зі styles.xml (фабрики з lxml-елементів).
# Використовуються у workbook.pyx::_parse_styles (D-1).
# ------------------------------------------------------------------

def color_from_element(elem):
    """Будує Color з <color>/<fgColor>/<bgColor>."""
    if elem is None:
        return Color()
    cdef object tint = elem.get('tint')
    return Color(
        rgb=elem.get('rgb'),
        indexed=elem.get('indexed'),
        theme=elem.get('theme'),
        tint=float(tint) if tint else 0.0,
    )


def font_from_element(elem):
    """Будує Font з <font> (b/i/u/strike/sz/color/name)."""
    cdef Font f = Font()
    cdef str ln
    cdef object v
    for child in elem:
        ln = _localname(child)
        if ln == 'b':
            f.bold = _flag_child(child)
        elif ln == 'i':
            f.italic = _flag_child(child)
        elif ln == 'u':
            f.underline = _flag_child(child)
        elif ln == 'strike':
            f.strike = _flag_child(child)
        elif ln == 'sz':
            v = child.get('val')
            if v is not None:
                try:
                    f.size = float(v)
                except (ValueError, TypeError):
                    pass
        elif ln == 'name':
            f.name = child.get('val')
        elif ln == 'color':
            f.color = color_from_element(child)
    return f


def fill_from_element(elem):
    """Будує PatternFill з <fill> (підтримує patternFill; gradientFill -> дефолт)."""
    cdef PatternFill pf = PatternFill()
    cdef str sln
    for child in elem:
        if _localname(child) == 'patternFill':
            pf.patternType = child.get('patternType')
            for sub in child:
                sln = _localname(sub)
                if sln == 'fgColor':
                    pf.fgColor = color_from_element(sub)
                elif sln == 'bgColor':
                    pf.bgColor = color_from_element(sub)
            break
    return pf


def side_from_element(elem):
    """Будує Side з <left>/<right>/<top>/<bottom>/<diagonal>."""
    cdef Side s = Side()
    s.style = elem.get('style')
    for child in elem:
        if _localname(child) == 'color':
            s.color = color_from_element(child)
    return s


def border_from_element(elem):
    """Будує Border з <border>."""
    cdef Border b = Border()
    cdef str ln
    for child in elem:
        ln = _localname(child)
        if ln == 'left':
            b.left = side_from_element(child)
        elif ln == 'right':
            b.right = side_from_element(child)
        elif ln == 'top':
            b.top = side_from_element(child)
        elif ln == 'bottom':
            b.bottom = side_from_element(child)
        elif ln == 'diagonal':
            b.diagonal = side_from_element(child)
    return b


def alignment_from_element(elem):
    """Будує Alignment з <alignment> (None -> дефолтне)."""
    if elem is None:
        return Alignment()
    cdef object tr = elem.get('textRotation')
    cdef object ind = elem.get('indent')
    return Alignment(
        horizontal=elem.get('horizontal'),
        vertical=elem.get('vertical'),
        text_rotation=int(tr) if tr else None,
        wrap_text=_attr_true(elem.get('wrapText')),
        shrink_to_fit=_attr_true(elem.get('shrinkToFit')),
        indent=int(ind) if ind else 0,
    )


def protection_from_element(elem):
    """Будує Protection з <protection> (None -> дефолтне)."""
    if elem is None:
        return Protection()
    cdef object locked = elem.get('locked')
    cdef object hidden = elem.get('hidden')
    return Protection(
        locked=_attr_true(locked) if locked is not None else True,
        hidden=_attr_true(hidden) if hidden is not None else False,
    )


# Default style
DEFAULT_STYLE = Style()