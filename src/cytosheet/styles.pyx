from lxml import etree

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

    def __init__(self, name=None, size=None, bold=False, italic=False, underline=False, 
                 strike=False, color=None):
        self.name = name
        self.size = size  # Can be None or a float
        self.bold = bold
        self.italic = italic
        self.underline = underline
        self.strike = strike
        self.color = color if color is not None else Color()

    def __repr__(self):
        return f"<Font name={self.name} size={self.size} bold={self.bold}>"

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

# Default style
DEFAULT_STYLE = Style()
