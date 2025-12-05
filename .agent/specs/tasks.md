# План розробки - Cytosheet

## Поточний стан проекту

**Версія**: 0.1.0
**Дата оновлення**: 2025-12-05
**Прогрес до v1.0**: 60%

**📊 Совместимость API**: см. `../.agent/COMPATIBILITY_SUMMARY.md`
- Базовый Core API: **~65%**
- Практические сценарии: **~90-95%**
- Полный API openpyxl: **~45%**

**📚 Детальный анализ**: `../.agent/api_compatibility_analysis.md` (392 строки)
**📈 Метрики**: `../.agent/COMPATIBILITY_METRICS.json`

---

## Пріоритизація завдань

### 🔴 Критично (Blocking для v1.0)
Без цих функцій неможлива базова сумісність з openpyxl.

### 🟡 Важливо (Nice to have для v1.0)
Потрібні для повної сумісності, але не блокують базове використання.

### 🟢 Бажано (Post v1.0)
Розширений функціонал, що покращує зручність використання.

---

## Milestone 1: Core API Completion (v0.2.0)

**Термін**: 2 тижні  
**Прогрес**: 🟡 60%

### TASK-001: Формули (Formulas) 🔴
**Пріоритет**: КРИТИЧНИЙ  
**Статус**: ⏳ В розробці  
**Час оцінка**: 3 дні

**Опис**: Додати повну підтримку формул Excel.

**Acceptance Criteria**:
- [x] Запис формули як рядка: `cell.value = "=SUM(A1:A10)"` та коректний запис у `<f>`
- [x] Читання формули: `cell.value` повертає текст формули, а не обчислене значення
- [x] Збереження/читання формул у XML (`<f>` тег) для простих кейсів (SUM, посилання, діапазони)
- [x] Підтримка формул у lazy-режимі (`worksheet.iter_rows(values_only=True)` повертає текст формули)
- [x] Атрибут `cell.data_type` визначає 'f' для формул у всіх режимах (установка через API, а не тільки при парсингу)
- [ ] Підтримка базових функцій: SUM, AVERAGE, COUNT, IF (з точки зору збереження/читання, без обчислювального движка)
- [ ] Підтримка посилань: A1, $A$1, Sheet2!A1 (зберігаються як текст без модифікації)
- [ ] Підтримка діапазонів: A1:B10 (зберігаються як текст без модифікації)

**Файли для зміни**:
- `src/cytosheet/cell.pyx` - додати атрибут `data_type` (✅ зроблено)
- `src/cytosheet/worksheet.pyx` - парсинг `<f>` тегів та генерація `<f>` при записі (✅ базово зроблено)
- `tests/test_formulas.py` - базові тести на запис/читання формул (✅ створено)
- `src/cytosheet/formula.pyx` - повноцінний модуль для роботи з формулами (📝 заплановано на v1.0+)

---

### TASK-002: Number Format 🔴
**Пріоритет**: КРИТИЧНИЙ  
**Статус**: ⏳ В розробці  
**Час оцінка**: 2 дні

**Опис**: Додати підтримку форматування чисел.

**Acceptance Criteria**:
- [x] Атрибут `cell.number_format`
- [x] Підтримка стандартних форматів: '0.00', '0.00%', 'dd/mm/yyyy' (через зберігання рядка формату)
- [x] Збереження форматів у styles.xml
- [x] Читання форматів з файлів

**Приклад API**:
```python
ws['A1'] = 0.5
ws['A1'].number_format = '0.00%'  # Відображається як "50.00%"

ws['B1'] = datetime.now()
ws['B1'].number_format = 'dd/mm/yyyy'  # "28/01/2025"
```

**Файли для зміни**:
- `src/cytosheet/cell.pyx` - додати атрибут `number_format`
- `src/cytosheet/styles.pyx` - розширити Style клас
- `src/cytosheet/workbook.pyx` - генерація styles.xml

---

### TASK-003: Row/Column Dimensions 🔴
**Пріоритет**: КРИТИЧНИЙ  
**Статус**: ✅ Виконано  
**Час оцінка**: 2 дні

**Опис**: Управління висотою рядків та шириною колонок.

**Acceptance Criteria**:
- [x] `worksheet.row_dimensions[row].height`
- [x] `worksheet.column_dimensions[col].width`
- [x] `worksheet.row_dimensions[row].hidden`
- [x] `worksheet.column_dimensions[col].hidden`
- [x] Збереження у worksheet XML
- [x] Читання з існуючих файлів

**Приклад API**:
```python
ws.row_dimensions[1].height = 30
ws.column_dimensions['A'].width = 20
ws.column_dimensions['B'].hidden = True
```

**Файли для зміни**:
- `src/cytosheet/worksheet.pyx` - додати атрибути `row_dimensions`, `column_dimensions`
- `src/cytosheet/dimensions.pyx` - новий модуль (створити)
- Парсинг `<row>`, `<col>` тегів у XML

---

### TASK-004: iter_cols() 🔴
**Пріоритет**: ВИСОКИЙ  
**Статус**: ✅ Виконано  
**Час оцінка**: 1 день

**Опис**: Додати метод ітерації по колонках (аналог iter_rows).

**Acceptance Criteria**:
- [x] Метод `worksheet.iter_cols(min_col, max_col, min_row, max_row, values_only)`
- [x] Підтримка lazy режиму
- [x] Тести сумісності з openpyxl

**Приклад API**:
```python
for col in ws.iter_cols(min_col=1, max_col=3, values_only=True):
    print(col)  # (cell1, cell2, cell3)
```

**Файли для зміни**:
- `src/cytosheet/worksheet.pyx` - додати метод `iter_cols()`

---

### TASK-005: append() метод 🔴
**Пріоритет**: ВИСОКИЙ  
**Статус**: ✅ Виконано  
**Час оцінка**: 1 день

**Опис**: Додавання рядка даних в кінець аркуша.

**Acceptance Criteria**:
- [x] Метод `worksheet.append(data)`
- [x] Автоматичне визначення наступного порожнього рядка
- [x] Підтримка списків та кортежів

**Приклад API**:
```python
ws.append([1, 2, 3])
ws.append(['Alice', 25, 'Developer'])
```

**Файли для зміни**:
- `src/cytosheet/worksheet.pyx` - додати метод `append()`
- Потрібен атрибут `max_row` для визначення позиції

---

### TASK-006: max_row, max_column 🔴
**Пріоритет**: ВИСОКИЙ  
**Статус**: ✅ Виконано  
**Час оцінка**: 1 день

**Опис**: Атрибути для визначення меж даних.

**Acceptance Criteria**:
- [x] `worksheet.max_row` - номер останнього рядка з даними
- [x] `worksheet.max_column` - номер останньої колонки з даними
- [x] Ефективне обчислення (кешування)

**Файли для зміни**:
- `src/cytosheet/worksheet.pyx` - додати property `max_row`, `max_column`

---

## Milestone 2: Advanced Features (v0.3.0)

**Термін**: 2 тижні  
**Прогрес**: 🔴 10%

### TASK-010: Freeze Panes 🟡
**Пріоритет**: ВИСОКИЙ  
**Статус**: ❌ Не розпочато  
**Час оцінка**: 1 день

**Acceptance Criteria**:
- [ ] `worksheet.freeze_panes = 'B2'`
- [ ] `worksheet.freeze_panes = None` для відміни
- [ ] Збереження у sheetViews

**Файли для зміни**:
- `src/cytosheet/worksheet.pyx`

---

### TASK-011: Insert/Delete Rows/Columns 🟡
**Пріоритет**: ВИСОКИЙ  
**Статус**: ❌ Не розпочато  
**Час оцінка**: 2 дні

**Acceptance Criteria**:
- [ ] `worksheet.insert_rows(idx, amount)`
- [ ] `worksheet.delete_rows(idx, amount)`
- [ ] `worksheet.insert_cols(idx, amount)`
- [ ] `worksheet.delete_cols(idx, amount)`
- [ ] Оновлення формул при зсуві

**Файли для зміни**:
- `src/cytosheet/worksheet.pyx`

---

### TASK-012: Data Validation 🟡
**Пріоритет**: СЕРЕДНІЙ  
**Статус**: ❌ Не розпочато  
**Час оцінка**: 3 дні

**Acceptance Criteria**:
- [ ] Створення `DataValidation` об'єкта
- [ ] Типи: list, whole, decimal, date, time, textLength, custom
- [ ] Додавання до комірки/діапазону
- [ ] Збереження у XML

**Приклад API**:
```python
dv = DataValidation(type='list', formula1='"Item1,Item2,Item3"')
ws.add_data_validation(dv)
dv.add('A1:A10')
```

**Файли для зміни**:
- `src/cytosheet/validation.py` - новий модуль
- `src/cytosheet/worksheet.pyx`

---

### TASK-013: Conditional Formatting 🟡
**Пріоритет**: СЕРЕДНІЙ  
**Статус**: ❌ Не розпочато  
**Час оцінка**: 4 дні

**Acceptance Criteria**:
- [ ] Правила на основі значень
- [ ] Кольорові шкали (ColorScale)
- [ ] Data bars
- [ ] Icon sets
- [ ] Збереження у XML

**Файли для зміни**:
- `src/cytosheet/formatting.py` - новий модуль
- `src/cytosheet/worksheet.pyx`

---

## Milestone 3: Visualization (v0.4.0)

**Термін**: 3 тижні  
**Прогрес**: 🔴 0%

### TASK-020: Charts (Діаграми) 🟡
**Пріоритет**: СЕРЕДНІЙ  
**Статус**: ❌ Не розпочато  
**Час оцінка**: 5 днів

**Acceptance Criteria**:
- [ ] BarChart
- [ ] LineChart
- [ ] PieChart
- [ ] ScatterChart
- [ ] Метод `worksheet.add_chart(chart, anchor)`
- [ ] Збереження у drawing.xml

**Приклад API**:
```python
from cytosheet.chart import BarChart, Reference

chart = BarChart()
data = Reference(ws, min_col=2, min_row=1, max_row=10)
cats = Reference(ws, min_col=1, min_row=2, max_row=10)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
ws.add_chart(chart, 'E5')
```

**Файли для зміни**:
- `src/cytosheet/chart/` - новий пакет
- `src/cytosheet/worksheet.pyx`

---

### TASK-021: Images 🟡
**Пріоритет**: НИЗЬКИЙ  
**Статус**: ❌ Не розпочато  
**Час оцінка**: 3 дні (перенесено у v1.0+)

**Acceptance Criteria (v1.0+)**:
- [ ] Клас `Image`
- [ ] Підтримка PNG, JPEG, GIF
- [ ] Метод `worksheet.add_image(img, anchor)`
- [ ] Масштабування

**Файли для зміни**:
- `src/cytosheet/drawing.py` - новий модуль
- `src/cytosheet/worksheet.pyx`

---

## Milestone 4: Advanced Workbook Features (v0.5.0)

**Термін**: 2 тижні  
**Прогрес**: 🔴 0%

### TASK-030: Named Ranges 🟢
**Пріоритет**: НИЗЬКИЙ  
**Статус**: ❌ Не розпочато  
**Час оцінка**: 2 дні

**Файли для зміни**:
- `src/cytosheet/workbook.pyx`

---

### TASK-031: Protection 🟢
**Пріоритет**: НИЗЬКИЙ  
**Статус**: ❌ Не розпочато  
**Час оцінка**: 2 дні

**Файли для зміни**:
- `src/cytosheet/worksheet.pyx`
- `src/cytosheet/workbook.pyx`

---

### TASK-032: Page Setup 🟢
**Пріоритет**: НИЗЬКИЙ  
**Статус**: ❌ Не розпочато  
**Час оцінка**: 2 дні

**Файли для зміни**:
- `src/cytosheet/worksheet.pyx`

---

### TASK-033: Comments 🟢
**Пріоритет**: НИЗЬКИЙ  
**Статус**: ❌ Не розпочато  
**Час оцінка**: 2 дні

**Файли для зміни**:
- `src/cytosheet/comments.py` - новий модуль
- `src/cytosheet/worksheet.pyx`

---

## Milestone 5: Testing & Documentation (v1.0.0)

**Термін**: 2 тижні  
**Прогрес**: 🟡 40%

### TASK-040: Comprehensive Test Suite 🔴
**Пріоритет**: КРИТИЧНИЙ  
**Статус**: ⏳ В розробці  
**Час оцінка**: 5 днів

**Acceptance Criteria**:
- [ ] Тести для всіх публічних API
- [ ] Compatibility tests (порівняння з openpyxl)
- [ ] Performance benchmarks
- [ ] Покриття коду >80%

**Файли**:
- `tests/test_compatibility.py` - новий файл
- `tests/test_formulas.py`
- `tests/test_charts.py`
- тощо

---

### TASK-041: Documentation 🔴
**Пріоритет**: ВИСОКИЙ  
**Статус**: ⏳ В розробці  
**Час оцінка**: 3 дні

**Acceptance Criteria**:
- [ ] Docstrings для всіх публічних API
- [ ] README з прикладами
- [ ] Migration guide від openpyxl
- [ ] API reference documentation

---

### TASK-042: Performance Optimization 🟡
**Пріоритет**: ВИСОКИЙ  
**Статус**: ⏳ В розробці  
**Час оцінка**: 3 дні

**Цілі**:
- [ ] Читання: 5x швидше openpyxl
- [ ] Запис: 3x швидше openpyxl
- [ ] Пам'ять: не більше 2x розміру файлу

### TASK-043: Табличні сценарії 🔴
**Пріоритет**: КРИТИЧНИЙ  
**Статус**: ⏳ В розробці  
**Час оцінка**: 3 дні

**Опис**: Забезпечити стабільну роботу основних сценаріїв компанії, повʼязаних з читанням/записом табличних даних у шаблони Excel.

**Acceptance Criteria**:
- [x] Підтримка сценарію «заповнення звіту з БД у шаблон» (приклад 1 з `examples.md`):
  - читання шаблону/створення книги `load_workbook(...)` / `Workbook()`
  - доступ до аркуша через `workbook.active` / `workbook.worksheets[0]` / `workbook[sheet_name]`
  - запис значень через `worksheet.cell(row=..., column=...).value = ...` та `worksheet[coord] = ...`
  - застосування `Alignment` до комірок
  - робота з `worksheet.column_dimensions[...]` (ширина колонок)
  - round‑trip через cytosheet і openpyxl (інтеграційний тест `test_example1_fill_template_from_db`).
- [x] Підтримка сценарію створення робочої книги з нуля з хедерами та стилями (приклад 2): еквівалентний функціонал через openpyxl-сумісне API (`Workbook`, `Worksheet`, `cell.style`, `Alignment`, `Font`, `row_dimensions[1].height`) + інтеграційний тест `test_example2_create_workbook_with_headers_and_styles`.
- [x] Підтримка сценарію копіювання аркуша-шаблону (приклад 3) через документований обхідний спосіб: створення нового листа, копіювання значень/стилей, використання `cell.font.color.rgb`, `column_dimensions`, `Workbook.remove(worksheet)`; інтеграційний тест `test_example3_copy_like_scenario_with_template_sheet`.
- [ ] Додаткові сценарії при необходимости (будуть добавлятися по мірі появи реальних кейсів).
