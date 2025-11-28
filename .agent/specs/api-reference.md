# API Reference - Cytosheet (openpyxl compatibility)

## Зміст

- [Workbook API](#workbook-api)
- [Worksheet API](#worksheet-api)
- [Cell API](#cell-api)
- [Styles API](#styles-api)
- [Formula API](#formula-api)
- [Charts API](#charts-api)
- [Images API](#images-api)
- [Data Validation API](#data-validation-api)

---

## Workbook API

### Створення та завантаження

#### `Workbook(write_only=False, iso_dates=False)`
Створює нову робочу книгу.

**Параметри**:
- `write_only` (bool): Режим тільки для запису (оптимізація пам'яті) [TODO]
- `iso_dates` (bool): Використовувати ISO формат дат [TODO]

**Приклад**:
```python
from cytosheet import Workbook

wb = Workbook()
```

**Статус**: ✅ Реалізовано (базово)

---

#### `load_workbook(filename, read_only=False, keep_vba=False, data_only=False, keep_links=True, rich_text=False)`
Завантажує існуючу робочу книгу.

**Параметри**:
- `filename` (str): Шлях до XLSX файлу
- `read_only` (bool): Режим тільки для читання (lazy loading)
- `keep_vba` (bool): Зберігати VBA макроси [TODO]
- `data_only` (bool): Читати тільки значення, не формули [TODO]
- `keep_links` (bool): Зберігати зовнішні посилання [TODO]
- `rich_text` (bool): Підтримка багатого тексту [TODO]

**Приклад**:
```python
from cytosheet import load_workbook

wb = load_workbook('example.xlsx')
wb_lazy = load_workbook('large.xlsx', read_only=True)
```

**Статус**: ✅ Реалізовано (частково: filename, read_only)

---

### Властивості Workbook

#### `workbook.active`
Повертає активний аркуш.

**Тип**: `Worksheet`

**Приклад**:
```python
ws = wb.active
```

**Статус**: ✅ Реалізовано

---

#### `workbook.sheetnames`
Список назв всіх аркушів.

**Тип**: `List[str]`

**Приклад**:
```python
names = wb.sheetnames  # ['Sheet1', 'Sheet2']
```

**Статус**: ✅ Реалізовано

---

#### `workbook.worksheets`
Список всіх аркушів.

**Тип**: `List[Worksheet]`

**Приклад**:
```python
for ws in wb.worksheets:
    print(ws.title)
```

**Статус**: ✅ Реалізовано (через wb._sheets.values())

---

### Методи Workbook

#### `workbook.create_sheet(title=None, index=None)`
Створює новий аркуш.

**Параметри**:
- `title` (str, optional): Назва аркуша. За замовчуванням "SheetN"
- `index` (int, optional): Позиція вставки [TODO]

**Повертає**: `Worksheet`

**Приклад**:
```python
ws1 = wb.create_sheet("MySheet")
ws2 = wb.create_sheet("Data", 0)  # Вставка на початок
```

**Статус**: ✅ Реалізовано (без index)

---

#### `workbook.remove(worksheet)` / `workbook.remove_sheet(worksheet)`
Видаляє аркуш.

**Параметри**:
- `worksheet` (Worksheet або str): Аркуш або його назва

**Приклад**:
```python
wb.remove_sheet("OldSheet")
# або
ws = wb["OldSheet"]
wb.remove(ws)
```

**Статус**: ✅ Реалізовано (тільки за назвою)

---

#### `workbook.save(filename)`
Зберігає робочу книгу у файл.

**Параметри**:
- `filename` (str): Шлях для збереження

**Приклад**:
```python
wb.save('output.xlsx')
```

**Статус**: ✅ Реалізовано

---

#### `workbook.close()`
Закриває робочу книгу і звільняє ресурси.

**Приклад**:
```python
wb.close()
```

**Статус**: ✅ Реалізовано

---

#### `workbook['SheetName']`
Доступ до аркуша за назвою.

**Повертає**: `Worksheet`

**Приклад**:
```python
ws = wb['Sheet1']
```

**Статус**: ✅ Реалізовано

---

### Додаткові властивості [TODO]

- `workbook.properties` - метадані документа
- `workbook.security` - налаштування безпеки
- `workbook.defined_names` - іменовані діапазони
- `workbook.calculation` - режим обчислення формул

---

## Worksheet API

### Властивості Worksheet

#### `worksheet.title`
Назва аркуша.

**Тип**: `str`

**Приклад**:
```python
ws.title = "New Name"
print(ws.title)  # "New Name"
```

**Статус**: ✅ Реалізовано

---

#### `worksheet.max_row`
Номер останнього рядка з даними.

**Тип**: `int`

**Приклад**:
```python
last_row = ws.max_row
```

**Статус**: ❌ Не реалізовано

---

#### `worksheet.max_column`
Номер останньої колонки з даними.

**Тип**: `int`

**Приклад**:
```python
last_col = ws.max_column
```

**Статус**: ❌ Не реалізовано

---

#### `worksheet.dimensions`
Діапазон використаних ячейок (наприклад, "A1:D10").

**Тип**: `str`

**Приклад**:
```python
print(ws.dimensions)  # "A1:D10"
```

**Статус**: ❌ Не реалізовано

---

### Доступ до ячейок

#### `worksheet['A1']`
Доступ до однієї ячейки.

**Повертає**: `Cell`

**Приклад**:
```python
cell = ws['A1']
ws['A1'] = 42
```

**Статус**: ✅ Реалізовано

---

#### `worksheet['A1:C3']`
Доступ до діапазону ячейок.

**Повертає**: `Tuple[Tuple[Cell]]`

**Приклад**:
```python
cells = ws['A1:C3']
for row in cells:
    for cell in row:
        print(cell.value)
```

**Статус**: ❌ Не реалізовано

---

#### `worksheet.cell(row, column, value=None)`
Доступ до ячейки за координатами (1-based).

**Параметри**:
- `row` (int): Номер рядка (починається з 1)
- `column` (int): Номер колонки (починається з 1)
- `value` (optional): Значення для встановлення

**Повертає**: `Cell`

**Приклад**:
```python
cell = ws.cell(row=1, column=1)
ws.cell(row=1, column=1, value=42)
```

**Статус**: ❌ Не реалізовано

---

### Ітерація

#### `worksheet.iter_rows(min_row=None, max_row=None, min_col=None, max_col=None, values_only=False)`
Ітерація по рядках.

**Параметри**:
- `min_row` (int, optional): Початковий рядок
- `max_row` (int, optional): Кінцевий рядок
- `min_col` (int, optional): Початкова колонка
- `max_col` (int, optional): Кінцева колонка
- `values_only` (bool): Повертати тільки значення (не Cell об'єкти)

**Повертає**: `Generator[Tuple[Cell]]` або `Generator[Tuple[Any]]`

**Приклад**:
```python
# З об'єктами Cell
for row in ws.iter_rows(min_row=2, max_row=10):
    for cell in row:
        print(cell.value)

# Тільки значення
for row in ws.iter_rows(values_only=True):
    print(row)  # (value1, value2, value3)
```

**Статус**: ✅ Реалізовано

---

#### `worksheet.iter_cols(min_row=None, max_row=None, min_col=None, max_col=None, values_only=False)`
Ітерація по колонках.

**Параметри**: Аналогічно `iter_rows()`

**Приклад**:
```python
for col in ws.iter_cols(min_col=1, max_col=3, values_only=True):
    print(col)  # (value1, value2, value3)
```

**Статус**: ❌ Не реалізовано

---

#### `worksheet.rows`
Генератор всіх рядків.

**Повертає**: `Generator[Tuple[Cell]]`

**Приклад**:
```python
for row in ws.rows:
    for cell in row:
        print(cell.value)
```

**Статус**: ❌ Не реалізовано (використовуйте iter_rows())

---

#### `worksheet.columns`
Генератор всіх колонок.

**Повертає**: `Generator[Tuple[Cell]]`

**Приклад**:
```python
for col in ws.columns:
    for cell in col:
        print(cell.value)
```

**Статус**: ❌ Не реалізовано (використовуйте iter_cols())

---

#### `worksheet.values`
Генератор тільки значень по рядках.

**Повертає**: `Generator[Tuple[Any]]`

**Приклад**:
```python
for row in ws.values:
    print(row)  # (value1, value2, value3)
```

**Статус**: ❌ Не реалізовано (використовуйте iter_rows(values_only=True))

---

### Модифікація даних

#### `worksheet.append(iterable)`
Додає рядок даних в кінець аркуша.

**Параметри**:
- `iterable`: Список, кортеж або словник з даними

**Приклад**:
```python
ws.append([1, 2, 3])
ws.append(['Alice', 25, 'Developer'])
```

**Статус**: ❌ Не реалізовано

---

#### `worksheet.insert_rows(idx, amount=1)`
Вставляє порожні рядки.

**Параметри**:
- `idx` (int): Індекс для вставки
- `amount` (int): Кількість рядків

**Приклад**:
```python
ws.insert_rows(2, 3)  # Вставити 3 рядки після рядка 1
```

**Статус**: ❌ Не реалізовано

---

#### `worksheet.delete_rows(idx, amount=1)`
Видаляє рядки.

**Параметри**:
- `idx` (int): Початковий індекс
- `amount` (int): Кількість рядків

**Приклад**:
```python
ws.delete_rows(2, 3)  # Видалити рядки 2, 3, 4
```

**Статус**: ❌ Не реалізовано

---

#### `worksheet.insert_cols(idx, amount=1)`
Вставляє порожні колонки.

**Статус**: ❌ Не реалізовано

---

#### `worksheet.delete_cols(idx, amount=1)`
Видаляє колонки.

**Статус**: ❌ Не реалізовано

---

### Об'єднання ячейок

#### `worksheet.merge_cells(range_string)` / `worksheet.merge_cells(start_row, start_column, end_row, end_column)`
Об'єднує ячейки.

**Параметри**:
- `range_string` (str): Діапазон типу "A1:B2"
- або координати: `start_row`, `start_column`, `end_row`, `end_column`

**Приклад**:
```python
ws.merge_cells('A1:B2')
# або
ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=2)
```

**Статус**: ✅ Реалізовано (тільки range_string)

---

#### `worksheet.unmerge_cells(range_string)`
Роз'єднує ячейки.

**Приклад**:
```python
ws.unmerge_cells('A1:B2')
```

**Статус**: ✅ Реалізовано

---

#### `worksheet.merged_cells`
Список об'єднаних діапазонів.

**Тип**: `List[str]`

**Приклад**:
```python
for range_str in ws.merged_cells:
    print(range_str)  # "A1:B2"
```

**Статус**: ✅ Реалізовано (через ws._merged_cells)

---

### Розміри рядків та колонок

#### `worksheet.row_dimensions[row_number]`
Доступ до параметрів рядка.

**Повертає**: `RowDimension`

**Приклад**:
```python
ws.row_dimensions[1].height = 30
ws.row_dimensions[1].hidden = True
```

**Статус**: ❌ Не реалізовано

---

#### `worksheet.column_dimensions[column_letter]`
Доступ до параметрів колонки.

**Повертає**: `ColumnDimension`

**Приклад**:
```python
ws.column_dimensions['A'].width = 20
ws.column_dimensions['B'].hidden = True
```

**Статус**: ❌ Не реалізовано

---

### Додаткові функції

#### `worksheet.freeze_panes`
Закріплення областей.

**Тип**: `str` або `None`

**Приклад**:
```python
ws.freeze_panes = 'B2'  # Закріпити рядок 1 та колонку A
ws.freeze_panes = None   # Відмінити закріплення
```

**Статус**: ❌ Не реалізовано

---

#### `worksheet.auto_filter`
Автофільтр для діапазону.

**Тип**: `str`

**Приклад**:
```python
ws.auto_filter.ref = 'A1:D10'
```

**Статус**: ❌ Не реалізовано

---

## Cell API

### Властивості Cell

#### `cell.value`
Значення ячейки.

**Тип**: `int`, `float`, `str`, `datetime`, `bool`, або `None`

**Приклад**:
```python
cell.value = 42
cell.value = "Text"
cell.value = datetime.now()
cell.value = "=SUM(A1:A10)"  # Формула
```

**Статус**: ✅ Реалізовано (без формул)

---

#### `cell.coordinate`
Координата ячейки (наприклад, "A1").

**Тип**: `str`

**Приклад**:
```python
print(cell.coordinate)  # "A1"
```

**Статус**: ✅ Реалізовано (cell.position)

---

#### `cell.row`
Номер рядка (1-based).

**Тип**: `int`

**Приклад**:
```python
print(cell.row)  # 1
```

**Статус**: ❌ Не реалізовано

---

#### `cell.column`
Номер колонки (1-based).

**Тип**: `int`

**Приклад**:
```python
print(cell.column)  # 1 (для колонки A)
```

**Статус**: ❌ Не реалізовано

---

#### `cell.column_letter`
Буквена позначення колонки.

**Тип**: `str`

**Приклад**:
```python
print(cell.column_letter)  # "A"
```

**Статус**: ❌ Не реалізовано

---

#### `cell.data_type`
Тип даних ячейки.

**Тип**: `str` - 'n' (number), 's' (string), 'b' (bool), 'f' (formula), 'd' (date), 'e' (error)

**Приклад**:
```python
print(cell.data_type)  # 'n' для числа
```

**Статус**: ❌ Не реалізовано

---

### Стилі Cell

#### `cell.font`
Шрифт ячейки.

**Тип**: `Font`

**Приклад**:
```python
from cytosheet import Font

cell.font = Font(name='Arial', size=12, bold=True)
```

**Статус**: ✅ Реалізовано (через cell.style.font)

---

#### `cell.border`
Рамка ячейки.

**Тип**: `Border`

**Приклад**:
```python
from cytosheet import Border, Side, Color

cell.border = Border(
    left=Side(style='thin', color=Color(rgb='FF0000'))
)
```

**Статус**: ✅ Реалізовано (через cell.style.border)

---

#### `cell.fill`
Заливка ячейки.

**Тип**: `PatternFill`

**Приклад**:
```python
from cytosheet import PatternFill, Color

cell.fill = PatternFill(
    patternType='solid',
    fgColor=Color(rgb='FFFF00')
)
```

**Статус**: ✅ Реалізовано (через cell.style.fill)

---

#### `cell.alignment`
Вирівнювання.

**Тип**: `Alignment`

**Приклад**:
```python
from cytosheet import Alignment

cell.alignment = Alignment(
    horizontal='center',
    vertical='center',
    wrap_text=True
)
```

**Статус**: ✅ Реалізовано (через cell.style.alignment)

---

#### `cell.number_format`
Формат числа.

**Тип**: `str`

**Приклад**:
```python
cell.number_format = '0.00'      # 2 знаки після коми
cell.number_format = '0.00%'     # Відсотки
cell.number_format = 'dd/mm/yyyy' # Дата
```

**Статус**: ❌ Не реалізовано

---

#### `cell.protection`
Захист ячейки.

**Тип**: `Protection`

**Приклад**:
```python
from cytosheet import Protection

cell.protection = Protection(locked=True, hidden=False)
```

**Статус**: ✅ Реалізовано (через cell.style.protection)

---

## Styles API

### Font

```python
from cytosheet import Font

font = Font(
    name='Arial',      # Назва шрифту
    size=12,          # Розмір
    bold=False,       # Жирний
    italic=False,     # Курсив
    underline='none', # none, single, double
    strike=False,     # Закреслений
    color=Color(rgb='FF000000')  # Колір
)
```

**Статус**: ✅ Реалізовано

---

### Border

```python
from cytosheet import Border, Side, Color

border = Border(
    left=Side(style='thin', color=Color(rgb='FF0000')),
    right=Side(style='thin', color=Color(rgb='FF0000')),
    top=Side(style='thin', color=Color(rgb='FF0000')),
    bottom=Side(style='thin', color=Color(rgb='FF0000')),
    diagonal=Side(style='thin', color=Color(rgb='FF0000')),
    diagonal_direction=0  # 0, 1, 2
)
```

**Side styles**: 'thin', 'medium', 'thick', 'double', 'dotted', 'dashed', etc.

**Статус**: ✅ Реалізовано

---

### PatternFill

```python
from cytosheet import PatternFill, Color

fill = PatternFill(
    patternType='solid',  # solid, gray125, darkGray, etc.
    fgColor=Color(rgb='FFFF00'),  # Колір переднього плану
    bgColor=Color(rgb='FFFFFF')   # Колір фону
)
```

**Статус**: ✅ Реалізовано

---

### Alignment

```python
from cytosheet import Alignment

alignment = Alignment(
    horizontal='center',    # general, left, center, right, justify
    vertical='center',      # top, center, bottom, justify
    text_rotation=0,        # Кут повороту (0-180)
    wrap_text=False,        # Перенос тексту
    shrink_to_fit=False,    # Зменшити до розміру
    indent=0                # Відступ
)
```

**Статус**: ✅ Реалізовано (частково)

---

### Protection

```python
from cytosheet import Protection

protection = Protection(
    locked=True,   # Заблокована
    hidden=False   # Прихована формула
)
```

**Статус**: ✅ Реалізовано

---

### Color

```python
from cytosheet import Color

# RGB
color = Color(rgb='FFFF00')  # Жовтий (ARGB формат)

# Indexed
color = Color(indexed=64)

# Theme
color = Color(theme=1)

# Auto
color = Color(auto=True)
```

**Статус**: ✅ Реалізовано (тільки RGB)

---

## Formula API [TODO]

### Запис формул

```python
# Проста формула
ws['A3'] = '=SUM(A1:A2)'

# Формула з посиланнями на інші аркуші (зберігається як текст)
ws['A3'] = '=Sheet2!A1 + Sheet3!B2'

# Абсолютні посилання (зберігаються як текст)
ws['A3'] = '=$A$1 + B1'
```

**Статус**: ⏳ Частково реалізовано
- Формула зберігається як рядок `"=..."`
- При записі у XML формується тег `<f>` без кешованого `<v>`

---

### Читання формул

```python
cell = ws['A3']
print(cell.value)       # '=SUM(A1:A2)'
print(cell.data_type)   # 'f' (при читанні з файлу)
```

**Статус**: ⏳ Частково реалізовано
- При парсингу `<f>` формула відновлюється як рядок `'=...'`
- Для lazy-режиму `iter_rows(values_only=True)` також повертає рядок формули
- `cell.data_type = 'f'` встановлюється при читанні, але поки не при всіх сценаріях запису

---

## Charts API [TODO]

### BarChart

```python
from cytosheet.chart import BarChart, Reference

chart = BarChart()
chart.title = "Sales Data"
chart.x_axis.title = "Month"
chart.y_axis.title = "Revenue"

data = Reference(ws, min_col=2, min_row=1, max_row=10)
cats = Reference(ws, min_col=1, min_row=2, max_row=10)

chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)

ws.add_chart(chart, 'E5')
```

**Статус**: ❌ Не реалізовано

---

### LineChart, PieChart, ScatterChart

Аналогічно BarChart.

**Статус**: ❌ Не реалізовано

---

## Images API [TODO]

```python
from cytosheet.drawing import Image

img = Image('logo.png')
img.width = 200
img.height = 100

ws.add_image(img, 'A1')
```

**Статус**: ❌ Не реалізовано

---

## Data Validation API [TODO]

```python
from cytosheet.worksheet.datavalidation import DataValidation

dv = DataValidation(
    type='list',
    formula1='"Item1,Item2,Item3"',
    allow_blank=True
)

ws.add_data_validation(dv)
dv.add('A1:A10')
```

**Статус**: ❌ Не реалізовано

---

## Легенда статусів

- ✅ **Реалізовано** - повністю працює
- ⏳ **Частково** - базова функціональність є
- ❌ **Не реалізовано** - функціонал відсутній
- 📝 **Документовано** - API описано, але не реалізовано

---

## Пріоритети реалізації (з урахуванням v1.0 та сценаріїв компанії)

### 🔴 Критичні (v1.0.0)
- **Читання/запис**:
  - `load_workbook(filename, read_only/lazy)` (чтение шаблонів і великих файлів)
  - `Workbook.save(filename)` / `save_virtual_workbook()`
- **Парсинг з підвищеною продуктивністю**:
  - Оптимізований парсинг листів (вже є: simple/standard/chunked)
  - Стабільність і сумісність з openpyxl для `iter_rows(values_only=True)`
- **Worksheet базовий API**:
  - `worksheet.cell(row, column, value=None)`
  - `worksheet.max_row`, `worksheet.max_column`, `worksheet.dimensions`
  - `worksheet.append(iterable)`
  - `worksheet.iter_cols(...)`
- **Формули**:
  - Запис/читання `cell.value = "=..."` як `<f>` у XML
  - Коректна робота в шаблонних сценаріях (SUM/IF, посилання, діапазони)
- **Стилі** (те, що реально використовується в корпоративних звітах):
  - `cell.font`, `cell.alignment`, `cell.border`, `cell.fill`, `cell.protection`
  - `cell.number_format` для чисел, дат, відсотків
  - `worksheet.row_dimensions[row].height`, `worksheet.column_dimensions[col].width`, `hidden`

### 🟡 Високі (v1.0.x)
- Freeze panes: `worksheet.freeze_panes`
- Insert/Delete rows/columns з оновленням формул
- Data Validation (мінімальний набір)

### 🟢 Бажані (v1.0+ / 1.x)
- Charts (діаграми)
- Images (зображення)
- Conditional Formatting (розширена)
- Named Ranges
- Comments
