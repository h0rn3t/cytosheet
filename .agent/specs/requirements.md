# Функціональні вимоги - Cytosheet

## Загальні вимоги

### FR-001: API Сумісність
**Пріоритет**: КРИТИЧНИЙ  
**Статус**: В розробці (60%)

Бібліотека cytosheet має бути 100% сумісною з публічним API openpyxl для основного функціоналу.

**Критерії прийняття**:
- ✅ Імпорт: `from cytosheet import Workbook, load_workbook` працює
- ✅ Базові операції: створення, читання, запис XLSX
- ⏳ Всі публічні методи openpyxl доступні
- ⏳ Сигнатури методів ідентичні
- ❌ Поведінка методів ідентична у 100% випадків

**Приклад**:
```python
# Має працювати однаково
from openpyxl import Workbook
# або
from cytosheet import Workbook

wb = Workbook()
ws = wb.active
ws['A1'] = 42
wb.save('file.xlsx')
```

---

### FR-002: Продуктивність
**Пріоритет**: ВИСОКИЙ  
**Статус**: Частково виконано

Cytosheet має бути швидшим за openpyxl мінімум у 2x для типових операцій.

**Критерії прийняття**:
- ✅ Читання файлів: 2-5x швидше
- ✅ Парсинг великих файлів (>10k рядків): 5-10x швидше
- ⏳ Запис файлів: 2-3x швидше
- ❌ Складні операції (формули, діаграми): не гірше openpyxl

**Бенчмарк цілі**:
| Операція | Openpyxl | Cytosheet | Speedup |
|----------|----------|-----------|---------|
| Читання 1k рядків | 0.5s | 0.1s | 5x |
| Читання 10k рядків | 3.0s | 0.5s | 6x |
| Запис 1k рядків | 0.8s | 0.3s | 2.7x |
| Пошук значення | 2.0s | 0.3s | 6.7x |

---

## Основний функціонал

### FR-010: Workbook Management
**Пріоритет**: КРИТИЧНИЙ  
**Статус**: ✅ Реалізовано

**Вимоги**:
- ✅ Створення нової робочої книги: `Workbook()`
- ✅ Завантаження існуючої: `load_workbook(filename)`
- ✅ Збереження: `workbook.save(filename)`
- ✅ Доступ до активного аркуша: `workbook.active`
- ✅ Створення нового аркуша: `workbook.create_sheet(title)`
- ✅ Видалення аркуша: `workbook.remove_sheet(title)`
- ✅ Список аркушів: `workbook.sheetnames`
- ✅ Доступ до аркуша за іменем: `workbook['SheetName']`

**Приклад API**:
```python
wb = Workbook()
ws1 = wb.active
ws2 = wb.create_sheet("Sheet2")
wb.sheetnames  # ['Sheet', 'Sheet2']
wb.save('workbook.xlsx')
```

---

### FR-011: Worksheet Operations
**Пріоритет**: КРИТИЧНИЙ  
**Статус**: ✅ Базово реалізовано

**Вимоги**:
- ✅ Доступ до ячейки: `worksheet['A1']`
- ✅ Присвоєння значення: `worksheet['A1'] = value`
- ✅ Ітерація по рядках: `worksheet.iter_rows()`
- ⏳ Ітерація по колонках: `worksheet.iter_cols()`
- ⏳ Додавання рядка: `worksheet.append(data)`
- ⏳ Вставка/видалення рядків: `insert_rows()`, `delete_rows()`
- ⏳ Вставка/видалення колонок: `insert_cols()`, `delete_cols()`
- ⏳ Максимальний рядок/колонка: `max_row`, `max_column`

**Приклад API**:
```python
ws = wb.active
ws['A1'] = 'Header'
ws.append([1, 2, 3])

for row in ws.iter_rows(min_row=2, values_only=True):
    print(row)
```

---

### FR-012: Cell Operations
**Пріоритет**: КРИТИЧНИЙ  
**Статус**: ✅ Базово реалізовано

**Вимоги**:
- ✅ Читання значення: `cell.value`
- ✅ Запис значення: `cell.value = data`
- ✅ Координати: `cell.coordinate`, `cell.row`, `cell.column`
- ⏳ Тип даних: автоматичне визначення (str, int, float, datetime, formula)
- ⏳ Формули: `cell.value = "=SUM(A1:A10)"`
- ⏳ Number format: `cell.number_format = "0.00"`

**Типи даних**:
```python
ws['A1'] = 42              # int
ws['A2'] = 3.14            # float
ws['A3'] = 'Text'          # str
ws['A4'] = datetime.now()  # datetime
ws['A5'] = '=A1+A2'        # formula
```

---

### FR-020: Styling
**Пріоритет**: ВИСОКИЙ  
**Статус**: ⏳ Частково реалізовано

**Вимоги**:
- ✅ Font: `Font(name, size, bold, italic, color)`
- ✅ Border: `Border(left, right, top, bottom)`
- ✅ PatternFill: `PatternFill(patternType, fgColor, bgColor)`
- ✅ Alignment: `Alignment(horizontal, vertical, wrap_text)`
- ✅ Protection: `Protection(locked, hidden)`
- ⏳ Number format: `cell.number_format = '0.00%'`
- ❌ Named styles: `wb.add_named_style()`

**Приклад API**:
```python
from cytosheet import Font, Border, Side, Color, PatternFill, Alignment

cell = ws['A1']
cell.font = Font(name='Arial', size=12, bold=True)
cell.border = Border(
    left=Side(style='thin', color=Color(rgb='FF0000'))
)
cell.fill = PatternFill(patternType='solid', fgColor=Color(rgb='FFFF00'))
cell.alignment = Alignment(horizontal='center', vertical='center')
```

---

### FR-021: Merged Cells
**Пріоритет**: ВИСОКИЙ  
**Статус**: ✅ Реалізовано

**Вимоги**:
- ✅ Об'єднання ячейок: `worksheet.merge_cells('A1:B2')`
- ✅ Роз'єднання: `worksheet.unmerge_cells('A1:B2')`
- ✅ Список об'єднаних: `worksheet.merged_cells.ranges`
- ✅ Збереження стану при запису/читанні

**Приклад API**:
```python
ws.merge_cells('A1:B2')
ws['A1'] = 'Merged Cell'
ws.unmerge_cells('A1:B2')
```

---

### FR-030: Formulas
**Пріоритет**: ВИСОКИЙ  
**Статус**: ❌ Не реалізовано

**Вимоги**:
- ❌ Запис формул: `cell.value = "=SUM(A1:A10)"`
- ❌ Читання формул: `cell.value` повертає формулу
- ❌ Збереження формул при запису
- ❌ Підтримка всіх стандартних функцій Excel

**Приклад API**:
```python
ws['A1'] = 10
ws['A2'] = 20
ws['A3'] = '=SUM(A1:A2)'  # Формула
ws['A3'].value  # '=SUM(A1:A2)'
```

---

### FR-040: Charts (Діаграми)
**Пріоритет**: СЕРЕДНІЙ  
**Статус**: ❌ Не реалізовано

**Вимоги**:
- ❌ Створення діаграм: `BarChart()`, `LineChart()`, etc.
- ❌ Додавання даних: `chart.add_data()`
- ❌ Вставка на аркуш: `worksheet.add_chart(chart, 'E5')`
- ❌ Збереження діаграм у файлі

**Приклад API**:
```python
from cytosheet.chart import BarChart, Reference

chart = BarChart()
data = Reference(ws, min_col=1, min_row=2, max_row=10)
chart.add_data(data)
ws.add_chart(chart, 'E5')
```

---

### FR-041: Images
**Пріоритет**: СЕРЕДНІЙ  
**Статус**: ❌ Не реалізовано

**Вимоги**:
- ❌ Вставка зображень: `worksheet.add_image(Image('logo.png'), 'A1')`
- ❌ Підтримка форматів: PNG, JPEG, GIF
- ❌ Масштабування та позиціонування

---

### FR-050: Data Validation
**Пріоритет**: СЕРЕДНІЙ  
**Статус**: ❌ Не реалізовано

**Вимоги**:
- ❌ Списки вибору: `DataValidation(type='list', formula1='"Item1,Item2"')`
- ❌ Числові обмеження
- ❌ Дати
- ❌ Довільні формули

---

### FR-051: Conditional Formatting
**Пріоритет**: СЕРЕДНІЙ  
**Статус**: ❌ Не реалізовано

**Вимоги**:
- ❌ Правила на основі значень
- ❌ Кольорові шкали
- ❌ Data bars
- ❌ Icon sets

---

### FR-060: Row/Column Dimensions
**Пріоритет**: ВИСОКИЙ  
**Статус**: ❌ Не реалізовано

**Вимоги**:
- ❌ Висота рядка: `worksheet.row_dimensions[1].height = 30`
- ❌ Ширина колонки: `worksheet.column_dimensions['A'].width = 20`
- ❌ Приховування: `worksheet.row_dimensions[1].hidden = True`
- ❌ Auto-size

**Приклад API**:
```python
ws.row_dimensions[1].height = 30
ws.column_dimensions['A'].width = 20
ws.column_dimensions['B'].hidden = True
```

---

### FR-061: Freeze Panes
**Пріоритет**: ВИСОКИЙ  
**Статус**: ❌ Не реалізовано

**Вимоги**:
- ❌ Закріплення рядків/колонок: `worksheet.freeze_panes = 'B2'`
- ❌ Відміна закріплення: `worksheet.freeze_panes = None`

---

### FR-070: Page Setup
**Пріоритет**: НИЗЬКИЙ  
**Статус**: ❌ Не реалізовано

**Вимоги**:
- ❌ Орієнтація: `worksheet.page_setup.orientation = 'landscape'`
- ❌ Розмір паперу
- ❌ Margins (відступи)
- ❌ Header/Footer

---

### FR-080: Protection
**Пріоритет**: НИЗЬКИЙ  
**Статус**: ❌ Не реалізовано

**Вимоги**:
- ❌ Захист аркуша: `worksheet.protection.sheet = True`
- ❌ Захист робочої книги
- ❌ Паролі

---

### FR-090: Named Ranges
**Пріоритет**: НИЗЬКИЙ  
**Статус**: ❌ Не реалізовано

**Вимоги**:
- ❌ Створення: `workbook.create_named_range('MyRange', worksheet, 'A1:B10')`
- ❌ Використання у формулах
- ❌ Видалення

---

### FR-100: Comments
**Пріоритет**: НИЗЬКИЙ  
**Статус**: ❌ Не реалізовано

**Вимоги**:
- ❌ Додавання коментаря: `cell.comment = Comment('Text', 'Author')`
- ❌ Видалення коментаря

---

## Нефункціональні вимоги

### NFR-001: Продуктивність пам'яті
Бібліотека має підтримувати:
- ✅ Lazy loading для великих файлів (>100 MB)
- ✅ Streaming читання (`iter_rows()`)
- ⏳ Streaming запис (write-only mode)
- Максимальне використання пам'яті: <2x розміру файлу

### NFR-002: Сумісність версій
- Python: 3.7+
- Cython: ≥3.0.11
- lxml: ≥4.9
- Формат файлів: XLSX (Office Open XML)

### NFR-003: Документація
- ✅ Docstrings для всіх публічних API
- ⏳ Повна документація у стилі openpyxl
- ⏳ Migration guide від openpyxl

### NFR-004: Тестування
- ✅ Unit tests для базового функціоналу
- ⏳ Тести сумісності з openpyxl (порівняння результатів)
- ⏳ Performance benchmarks
- ⏳ Покриття коду: >80%

---

## Легенда статусів

- ✅ **Реалізовано** - функціонал повністю працює
- ⏳ **В розробці** - частково реалізовано або в процесі
- ❌ **Не реалізовано** - функціонал відсутній
- 🔴 **Критично** - блокує основну функціональність
- 🟡 **Важливо** - потрібно для сумісності
- 🟢 **Бажано** - розширений функціонал