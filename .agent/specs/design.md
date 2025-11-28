# Архітектурний дизайн - Cytosheet

## Огляд архітектури

Cytosheet побудовано на основі Cython для високої продуктивності при збереженні сумісності з openpyxl API.

### Принципи дизайну

1. **API-first**: Зовнішнє API ідентичне openpyxl
2. **Performance-critical**: Використання Cython для критичних операцій
3. **Memory-efficient**: Lazy loading та streaming для великих файлів
4. **OOXML compliance**: Повна відповідність стандарту Office Open XML

---

## Модульна архітектура

```
cytosheet/
├── __init__.py          # Публічне API
├── cell.pyx             # Cell клас (Cython)
├── workbook.pyx         # Workbook клас (Cython)
├── worksheet.pyx        # Worksheet клас (Cython)
├── excel.pyx            # Функції load_workbook (Cython)
├── styles.pyx           # Стилі (Cython)
├── formula.pyx          # Підтримка формул [TODO]
├── charts.py            # Діаграми [TODO]
├── drawing.py           # Зображення [TODO]
└── utils/
    ├── xml.pyx          # XML утиліти (Cython)
    └── constants.py     # Константи OOXML
```

---

## Ключові компоненти

### 1. Cell (cell.pyx)

**Призначення**: Представлення однієї ячейки з даними та стилями.

**Cython оптимізації**:
```cython
cdef class Cell:
    cdef str position          # Координата ячейки (A1)
    cdef public object value   # Значення (int, float, str, formula)
    cdef public object style   # Стиль ячейки
    cdef public object parent  # Посилання на Worksheet
    cdef public bint is_merged_cell  # Прапорець об'єднаної ячейки
    cdef public str merged_range     # Діапазон об'єднання
```

**Методи**:
- `get_position()` - швидкий доступ до координат
- `set_value(value)` - встановлення значення з автоматичним визначенням типу
- `get_value()` - отримання значення
- `set_style(style)` - застосування стилю
- `get_style()` - отримання стилю

**Типи значень**:
```python
cell.value = 42              # int
cell.value = 3.14            # float
cell.value = "Text"          # str
cell.value = datetime.now()  # datetime
cell.value = "=SUM(A1:A10)"  # formula string
```

---

### 2. Worksheet (worksheet.pyx)

**Призначення**: Робота з одним аркушем (worksheet) робочої книги.

**Структура даних**:
```cython
cdef class Worksheet:
    cdef public dict _cells              # {coordinate: Cell}
    cdef public str title                # Назва аркуша
    cdef public list _shared_strings     # Shared strings таблиця
    cdef public object _archive          # ZipFile для lazy loading
    cdef public str _sheet_path          # Шлях до XML в архіві
    cdef public bint _preloaded          # Чи завантажені дані
    cdef public bint _is_small_file      # Для вибору стратегії парсингу
    cdef public list _merged_cells       # Список об'єднаних діапазонів
```

**Ключові методи**:

1. **Доступ до ячейок**:
```python
ws['A1'] = 'Value'    # __setitem__
value = ws['A1']      # __getitem__
```

2. **Streaming читання**:
```python
for row in ws.iter_rows(values_only=True):
    process(row)
```

3. **Парсинг стратегії**:
- **Малі файли** (<50KB): `_parse_sheet_simple()` - string parsing
- **Середні файли** (50KB-100KB): `_parse_sheet_standard()` - стандартний iterparse
- **Великі файли** (>100KB): `_parse_sheet_chunked()` - батчевий парсинг

**Оптимізації парсингу**:
```cython
cdef void _parse_sheet_chunked(self, bytes xml_data):
    # Батчева обробка по 1000 ячейок
    cdef dict temp_batch = {}
    cdef int batch_size = 1000
    
    for event, cell_elem in context:
        # Обробка ячейки
        temp_batch[col_ref] = Cell(...)
        
        if len(temp_batch) >= batch_size:
            self._cells.update(temp_batch)
            temp_batch.clear()
```

4. **Об'єднання ячейок**:
```python
ws.merge_cells('A1:B2')    # Об'єднання
ws.unmerge_cells('A1:B2')  # Роз'єднання
```

---

### 3. Workbook (workbook.pyx)

**Призначення**: Управління робочою книгою та її аркушами.

**Структура**:
```cython
cdef class Workbook:
    cdef public dict _sheets              # {title: Worksheet}
    cdef public list _shared_strings      # Глобальні shared strings
    cdef public int _active_sheet_index   # Індекс активного аркуша
    cdef public bint _lazy                # Режим lazy loading
    cdef public object _archive           # ZipFile для читання
```

**Життєвий цикл**:

1. **Створення нової книги**:
```python
wb = Workbook()  # Створює один лист 'Sheet'
```

2. **Завантаження існуючої**:
```python
wb = load_workbook('file.xlsx', lazy=False)  # Повне завантаження
wb = load_workbook('file.xlsx', lazy=True)   # Lazy loading
```

3. **Робота з аркушами**:
```python
ws = wb.active                    # Активний аркуш
ws = wb['SheetName']              # Доступ за іменем
ws = wb.create_sheet('NewSheet')  # Створення
wb.remove_sheet('OldSheet')       # Видалення
```

4. **Збереження**:
```python
wb.save('output.xlsx')            # Збереження у файл
data = wb.save_virtual_workbook() # Збереження в пам'ять
```

**Генерація XLSX структури**:

XLSX це ZIP архів з XML файлами:
```
archive.zip
├── [Content_Types].xml      # MIME типи
├── _rels/
│   └── .rels                # Основні зв'язки
└── xl/
    ├── workbook.xml         # Структура книги
    ├── _rels/
    │   └── workbook.xml.rels # Зв'язки аркушів
    ├── worksheets/
    │   ├── sheet1.xml       # Дані аркуша 1
    │   └── sheet2.xml       # Дані аркуша 2
    ├── sharedStrings.xml    # Таблиця строк
    └── styles.xml           # Таблиця стилів
```

**Оптимізація генерації XML**:
```cython
cdef str _generate_sheet_elements(self):
    # Використовуємо string concatenation замість XML builders
    cdef list elements = []
    for i, sheet in enumerate(self._sheets.values()):
        elements.append(f'<sheet name="{sheet.title}" sheetId="{i + 1}"/>')
    return "".join(elements)
```

---

### 4. Styles (styles.pyx)

**Призначення**: Підтримка форматування ячейок.

**Класи стилів**:

```python
# Колір
class Color:
    rgb: str = 'FF000000'  # ARGB формат

# Сторона рамки
class Side:
    style: str = 'thin'    # thin, medium, thick, etc.
    color: Color = Color()

# Рамка
class Border:
    left: Side
    right: Side
    top: Side
    bottom: Side
    
    def __add__(self, other):  # Комбінація рамок
        return Border(...)

# Шрифт
class Font:
    name: str = 'Calibri'
    size: float = 11
    bold: bool = False
    italic: bool = False
    underline: str = 'none'
    color: Color = Color()

# Заливка
class PatternFill:
    patternType: str = 'solid'  # solid, gray125, etc.
    fgColor: Color = Color()
    bgColor: Color = Color()

# Вирівнювання
class Alignment:
    horizontal: str = 'general'  # left, center, right, etc.
    vertical: str = 'bottom'     # top, center, bottom
    wrap_text: bool = False
    shrink_to_fit: bool = False

# Захист
class Protection:
    locked: bool = True
    hidden: bool = False

# Комплексний стиль
class Style:
    font: Font
    border: Border
    fill: PatternFill
    alignment: Alignment
    protection: Protection
    number_format: str = 'General'
```

**Використання**:
```python
cell.font = Font(bold=True, size=14)
cell.border = Border(left=Side(style='thick'))
cell.fill = PatternFill(patternType='solid', fgColor=Color(rgb='FFFF00'))
cell.alignment = Alignment(horizontal='center')
```

**Оптимізації**:
- Immutable об'єкти стилів (копіювання при модифікації)
- Кешування часто використовуваних стилів
- Індексація стилів у styles.xml

---

### 5. Formula Support [TODO]

**Призначення**: Підтримка формул Excel.

**Архітектура**:
```python
class Formula:
    text: str           # "=SUM(A1:A10)"
    tokens: List[Token] # Розібраний вираз
    
    @staticmethod
    def parse(text: str) -> Formula:
        """Парсинг формули у токени"""
        
    def to_xml(self) -> str:
        """Серіалізація у XML"""
```

**Типи формул**:
- Стандартні функції: SUM, AVERAGE, COUNT, IF, VLOOKUP
- Посилання на ячейки: A1, $A$1, Sheet2!A1
- Діапазони: A1:B10
- Оператори: +, -, *, /, ^, &

**Збереження у XML**:
```xml
<c r="A3">
    <f>SUM(A1:A2)</f>
    <v>30</v>  <!-- Cached value -->
</c>
```

---

## Робота з OOXML

### XML Parsing

**Бібліотека**: lxml (C бібліотека libxml2)

**Стратегії парсингу**:

1. **DOM parsing** (малі файли):
```python
tree = etree.fromstring(xml_data)
cells = tree.xpath('//c')
```

2. **SAX parsing** (великі файли):
```python
context = etree.iterparse(stream, events=('end',), tag='c')
for event, elem in context:
    process(elem)
    elem.clear()  # Очистка пам'яті
```

3. **Hybrid approach** (cytosheet):
```python
# Вибір стратегії на основі розміру
if xml_size < 50KB:
    use_simple_string_parsing()
elif xml_size < 100KB:
    use_standard_iterparse()
else:
    use_chunked_batched_parsing()
```

### XML Generation

**Підхід**: String templating замість XML builders для швидкості.

```python
# Швидко
xml = f'<c r="{ref}" t="n"><v>{value}</v></c>'

# Повільно
cell = etree.Element('c', r=ref, t='n')
v = etree.SubElement(cell, 'v')
v.text = str(value)
```

---

## Оптимізації продуктивності

### 1. Cython типізація

```cython
# Замість
def get_value(self):
    return self.value

# Використовуємо
cpdef object get_value(self):
    return self.value
```

**Виграш**: 2-3x швидше для часто викликаних методів.

### 2. Batch processing

```python
# Замість побічної обробки
for cell in cells:
    process(cell)

# Батчева
batch = []
for cell in cells:
    batch.append(cell)
    if len(batch) >= 1000:
        process_batch(batch)
        batch.clear()
```

### 3. Memory pools

```python
# Перевикористання об'єктів Cell
cell_pool = []

def get_cell():
    if cell_pool:
        return cell_pool.pop()
    return Cell()

def release_cell(cell):
    cell.reset()
    cell_pool.append(cell)
```

### 4. String interning

```python
# Для координат ячейок
coord_cache = {}

def get_coord(col, row):
    key = (col, row)
    if key not in coord_cache:
        coord_cache[key] = f'{col}{row}'
    return coord_cache[key]
```

---

## Lazy Loading

**Концепція**: Завантаження даних тільки при потребі.

```python
wb = load_workbook('large.xlsx', lazy=True)
ws = wb['Sheet1']  # Не завантажує дані

# Завантаження при першому доступі
for row in ws.iter_rows():  # Streaming читання
    process(row)
```

**Переваги**:
- Швидкий старт
- Низьке використання пам'яті
- Можливість обробки файлів >1GB

---

## Стратегія тестування

### 1. Unit Tests
```python
def test_cell_value():
    cell = Cell(position='A1')
    cell.value = 42
    assert cell.value == 42
```

### 2. Compatibility Tests
```python
def test_openpyxl_compatibility():
    # Test with openpyxl
    from openpyxl import Workbook as OpenpyxlWB
    wb1 = OpenpyxlWB()
    wb1.active['A1'] = 42
    
    # Test with cytosheet
    from cytosheet import Workbook
    wb2 = Workbook()
    wb2.active['A1'] = 42
    
    # Compare results
    assert wb2.active['A1'].value == wb1.active['A1'].value
```

### 3. Performance Benchmarks
```python
def benchmark_read(library, file):
    start = time.time()
    wb = library.load_workbook(file)
    ws = wb.active
    for row in ws.iter_rows():
        pass
    return time.time() - start
```

---

## Майбутні покращення

### 1. Паралелізація
- Багатопотокове читання великих файлів
- Паралельний парсинг множинних аркушів

### 2. Кешування
- Кешування розібраних формул
- Кешування XML структур

### 3. Compression
- Оптимізація розміру згенерованих XLSX
- Підтримка різних рівнів компресії

### 4. Extensibility
- Plugin система для кастомних функцій
- Hooks для preprocessing/postprocessing