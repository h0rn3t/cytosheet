# Приклади використання Cytosheet

Цей документ описує ключові сценарії роботи з бібліотекою Cytosheet, які є критичними для компанії.

Основний фокус: **читання та запис табличних даних** у XLSX-шаблони та створювані файли, зі збереженням сумісності з openpyxl API.

---

## 1. Заповнення шаблону звіту з БД (приклад 1)

**Сценарій:**
- Є готовий XLSX-шаблон (із заголовками, стилями, формулами).
- З БД приходить табличний результат `db_result`.
- Потрібно:
  - записати дату/час формування звіту у фіксовану комірку;
  - построково записати табличні дані починаючи з певного рядка;
  - застосувати вирівнювання до всіх комірок;
  - автоматично підігнати ширину колонок.

```python
from cytosheet import load_workbook, Alignment
from openpyxl.utils import get_column_letter
import datetime

first_row_number = 5
date_cell = "G1"

workbook = load_workbook(template_path)
worksheet = workbook.worksheets[0]

worksheet[date_cell] = (
    f"Дата та час генерації звіту: {localize_datetime(datetime.datetime.utcnow()):%d.%m.%Yр. %Hг.%Mхв.}"
)

alignment_cell_style = Alignment(horizontal="left")
column_widths = []

for row_number, row_data in enumerate(db_result, first_row_number):
    for column_number, value in enumerate(row_data, 1):
        cell = worksheet.cell(row=row_number, column=column_number)
        cell.value = value
        cell.alignment = alignment_cell_style

        if len(column_widths) > column_number:
            if len(str(value)) > column_widths[column_number - 1]:
                column_widths[column_number - 1] = len(str(value))
        else:
            column_widths += [len(str(value))]

for column_number, column_width in enumerate(column_widths, 1):
    col_dim = worksheet.column_dimensions[get_column_letter(column_number)]
    if col_dim.width < (column_width + 3):
        col_dim.width = column_width + 3

workbook.save(output_path)
```

**Вимоги до Cytosheet для цього сценарію:**
- `load_workbook(path)` з підтримкою існуючих стилів/формул шаблону.
- `workbook.worksheets` (індексований доступ до аркушів).
- `worksheet.__getitem__` для прямого доступу (`worksheet["G1"] = ...`).
- `worksheet.cell(row=..., column=...)` (FR-011 / TASK-003+TASK-006).
- `Alignment` та застосування через `cell.alignment`.
- `worksheet.column_dimensions[col_letter].width` (Row/Column dimensions, TASK-003).

---

## 2. Створення книги з нуля з хедерами та стилями (аналог прикладу 2)

У прикладі 2 використовується кастомний API (`new_sheet`, `set_cell_value`, `set_col_style`, `set_row_style`).
У Cytosheet ціль — надати **еквівалентний функціонал** через API openpyxl-сумісного типу.

```python
from cytosheet import Workbook, Font, Alignment, Style

workbook = Workbook()
ws = workbook.create_sheet("Логіни користувачів")

headers = [
    "Логін користувача",
    "Організація",
    "Дата та час входу",
    "Успішність входу",
    "Метод входу",
]

# Стиль заголовків
header_font = Font(bold=True, size=9)
header_alignment = Alignment(horizontal="center", vertical="center")
header_style = Style(font=header_font, alignment=header_alignment)

# Запис заголовків у перший рядок
for col_idx, title in enumerate(headers, 1):
    cell = ws.cell(row=1, column=col_idx)
    cell.value = title
    cell.style = header_style

# Висота першого рядка (TASK-003 RowDimensions)
# ws.row_dimensions[1].height = 40

workbook.save(output_path)
```

**Вимоги до Cytosheet:**
- `Workbook.create_sheet(title)` (є).
- `worksheet.cell(row, column, value=None)` (TASK-003/FR-011).
- Робота зі стилями через `cell.style`, `Font`, `Alignment`, `Style` (вже базово реалізовано).
- `row_dimensions[1].height` (TASK-003).

---

## 3. Копіювання шаблонного аркуша і заповнення даними (приклад 3)

**Сценарій:**
- Є шаблонний аркуш у робочій книзі.
- Для кожної групи даних потрібно:
  - скопіювати аркуш-шаблон;
  - перейменувати його;
  - дописати службову інформацію у фіксовані комірки;
  - заповнити табличні дані з умовним форматуванням (наприклад, колір шрифту залежно від значення);
  - підлаштувати ширину колонок;
  - у кінці видалити початковий трафаретний аркуш.

```python
from cytosheet import load_workbook
from openpyxl.utils import get_column_letter
from datetime import datetime

workbook = load_workbook(template_path)

for worksheet_data in CountOfMPOfTSByGridArea.get_aggregation_data():
    # TODO: copy_worksheet API або еквівалентний обхідний варіант
    target_worksheet = workbook.copy_worksheet(workbook.worksheets[0])
    target_worksheet.title = worksheet_data[0]

    target_worksheet["E3"] = (
        f"Дата/час формування звіту: {localize_datetime(report_created_at):%d.%m.%Yр. %Hг.%Mхв.}"
    )
    target_worksheet["B4"] = "Дата за яку формується звіт: {:%d.%m.%Yр.}".format(
        datetime.strptime(worksheet_data[0], "%Y-%m-%d")
    )

    column_widths = []
    for row_number, row_data in enumerate(worksheet_data[1], 8):
        color_for_ts_cell = "FF00FF00" if row_data[3] == row_data[4] else "FFFF0000"
        for column_number, value in enumerate(row_data, 1):
            cell = target_worksheet.cell(row=row_number, column=column_number)
            cell.value = value
            if column_number in [5, 6]:
                cell.font.color.rgb = color_for_ts_cell

            if len(column_widths) > column_number:
                if len(str(value)) > column_widths[column_number - 1]:
                    column_widths[column_number - 1] = len(str(value))
            else:
                column_widths += [len(str(value))]

    for column_number, column_width in enumerate(column_widths, 1):
        col_dim = target_worksheet.column_dimensions[get_column_letter(column_number)]
        if col_dim.width < (column_width + 3):
            col_dim.width = column_width + 3

# Видаляємо лист-трафарет
workbook.remove(workbook.worksheets[0])
workbook.save(output_path)
```

**Вимоги до Cytosheet:**
- `load_workbook(template_path)`.
- `workbook.worksheets` (послідовність аркушів).
- API для копіювання аркуша:
  - Або `workbook.copy_worksheet(source_ws)` як в openpyxl (в ідеалі до v1.0),
  - Або документований обхідний варіант (створити новий аркуш і вручну копіювати комірки/стилі).
- `worksheet.cell(row, column)`.
- Робота з шрифтами: `cell.font.color.rgb`.
- `worksheet.column_dimensions[col].width`.
- Видалення аркуша: `workbook.remove(worksheet)` або `remove_sheet(title)`.

---

## 4. Узагальнений перелік потрібного функціоналу до v1.0

З урахуванням сценаріїв вище та вимоги до базової сумісності з openpyxl, для **v1.0** Cytosheet має підтримувати:

1. **Робоча книга (Workbook):**
   - Створення: `Workbook()`.
   - Завантаження: `load_workbook(path, read_only=False, data_only=False)` (мінімум: `filename`, `read_only`/`lazy`).
   - Доступ до аркушів: `workbook.active`, `workbook.worksheets`, `workbook[sheet_name]`.
   - Створення/видалення аркушів: `create_sheet(title)`, `remove(worksheet)` / `remove_sheet(title)`.
   - (Бажано до v1.0) `copy_worksheet(source_ws)` – або чітко задокументований обхідний спосіб.
   - Збереження: `workbook.save(path)`, `save_virtual_workbook()`.

2. **Аркуші (Worksheet):**
   - Індексований доступ до комірок: `ws['A1']`, `ws['A1:C3']` (мінімум — одинична комірка).
   - Координатний доступ: `ws.cell(row=..., column=..., value=None)`.
   - Ітерація: `iter_rows(..., values_only=True/False)`, `iter_cols(..., values_only=True/False)`.
   - Додавання рядків: `append(iterable)`.
   - Метадані діапазону: `max_row`, `max_column`, `dimensions`.
   - Об'єднання: `merge_cells(range_string)`, `unmerge_cells(range_string)`, `merged_cells`.
   - Розміри: `row_dimensions[row].height`, `column_dimensions[col].width`, `hidden`.
   - Freeze panes: `freeze_panes` (бажано до v1.0, але не критично для табличних сценаріїв).

3. **Комірки (Cell):**
   - `cell.value` для типів: int, float, str, bool, datetime, formula (рядок із '=...').
   - `cell.data_type` (принаймні: 'n', 's', 'b', 'f', 'd').
   - Координати: `cell.coordinate`, `cell.row`, `cell.column`, `cell.column_letter`.
   - Стилі: `cell.style`, а також окремі властивості через `Style`:
     - `font` (name, size, bold, italic, color.rgb).
     - `alignment` (horizontal, vertical, wrap_text тощо).
     - `border`, `fill`, `protection` – базово.
     - `number_format` – як мінімум для дат і чисел (%, десяткові).

4. **Формули:**
   - Запис/читання рядків формул: `cell.value = "=SUM(A1:A10)"`.
   - Збереження у XML через `<f>`.
   - Повернення формули як тексту при читанні (`data_only=False`).
   - Не обовʼязково мати власний обчислювальний движок до v1.0 (можна покладатися на Excel для обчислень).

5. **Продуктивність:**
   - Lazy/largе-file режим (`read_only=True` / `lazy=True`), `iter_rows(values_only=True)` для великих файлів.
   - Відсутність істотних регресій у порівнянні з поточними бенчмарками.

6. **Що можна відкласти на v1.0+ (після базової сумісності):**
   - Повноцінний модуль формул (`formula.pyx`) із парсером/движком.
   - Діаграми (Charts API).
   - Зображення (Images API, повністю перенесено на v1.0+).
   - Data validation, conditional formatting, comments, named ranges, pivot tables, VBA.

Цей список синхронізовано з `requirements.md` та `tasks.md` і фокусується на реальних робочих сценаріях.

