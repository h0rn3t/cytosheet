# Cytosheet

Cytosheet is a PoC library for working with XLSX files, written in Cython.

## Roadmap

1. [] High performance due to Cython and libxml2.
2. [] Support for reading large XLSX files using SAX parsing, allowing files to be processed in chunks without loading the entire file into memory.
3. [] Support for creating and writing XLSX files.
4. [] Lightweight and user-friendly API, similar to openpyxl.

## API совместимость с openpyxl (кратко)

Оценка текущего состояния по сравнению с openpyxl:

- Общая совместимость базового API (Workbook / Worksheet / Cell / стили): **~55–60%**.
- Совместимость в рамках реальных табличных сценариев из `.agent/specs/examples.md` (генерация отчётов, работа с шаблонами): **~90–95%**.

### Что уже совместимо (основное)

- `Workbook` / `load_workbook`:
  - Создание: `Workbook()`.
  - Загрузка: `load_workbook(path, read_only=False, data_only=False, ...)` (лишние аргументы принимаются и игнорируются).
  - Доступ к листам: `active`, `sheetnames`, `__getitem__`, `create_sheet(title)`, `remove(worksheet)`, `remove_sheet(title)`.
  - `save(path)`, `save_virtual_workbook()`.
- `Worksheet`:
  - Доступ к ячейкам: `ws['A1']`, диапазоны `ws['A1:C3']`.
  - `cell(row, column, value=None)`.
  - `append(iterable)`.
  - Итерация: `iter_rows`, `iter_cols`, свойства `rows`, `columns`, `values`.
  - Границы: `max_row`, `max_column`, `dimensions`.
  - Объединение: `merge_cells`, `unmerge_cells`.
  - Размеры строк/колонок: `row_dimensions[row].height/hidden`, `column_dimensions[col].width/hidden` (round‑trip с openpyxl).
- `Cell`:
  - `value`, `coordinate`, `row`, `column`, `column_letter`.
  - Формулы как строки с `=` и `<f>` в XML, `data_type='f'`.
  - `number_format` (proxy к `Style.numberFormat`), запись/чтение `styles.xml` для числовых форматов.
  - Прокси к стилям: `cell.font`, `cell.alignment`, `cell.border`, `cell.fill`, `cell.protection`.
- `Styles`:
  - `Font`, `Color`, `Border/Side`, `PatternFill`, `Alignment`, `Protection`, `Style`, `DEFAULT_STYLE`.
  - Минимальный `styles.xml`, который без ошибок читается openpyxl.

### Что пока не покрыто

- Диаграммы, изображения, data validation, conditional formatting, comments, named ranges, filters, pivot tables, VBA.
- Продвинутый модуль формул (`formula.pyx`) и собственный движок вычислений.
- Расширенные настройки страницы (page setup), freeze panes, protection на уровне листа/книги.

### Практическая совместимость (сценарии отчётов)

Сценарии из `.agent/specs/examples.md`:

1. **Заполнение отчёта из БД в шаблон**:
   - Создание/загрузка книги, доступ к `worksheet[coord]` и `worksheet.cell`.
   - Применение `Alignment` через `cell.alignment`.
   - Управление шириной колонок через `worksheet.column_dimensions[col].width`.
   - Round‑trip через cytosheet и openpyxl (файлы читаются обеими библиотеками).

2. **Создание книги с нуля с заголовками и стилями**:
   - `Workbook()`, `create_sheet(title)`.
   - Установка стилей заголовков через `Font`, `Alignment` и `cell.font` / `cell.alignment`.
   - `row_dimensions[1].height` c корректным round‑trip через cytosheet и openpyxl.

3. **Сценарий “копирования” шаблонного листа и заполнения отчёта**:
   - Поддержано через создание нового листа и ручное копирование содержимого/стилей.
   - Используется `cell.font.color.rgb`, `column_dimensions`, `remove(worksheet)`.
   - Файл без ошибок открывается openpyxl, значения и базовые стили совпадают.

Для типовых задач генерации отчётов (чтение/запись табличных данных, стили ячеек, ширины колонок, базовые формулы) Cytosheet уже может использоваться как практический аналог openpyxl, с учётом указанных ограничений по продвинутым возможностям Excel.

## Installation

### Requirements

- Python 3.7+
- Cython
- lxml
### Install from source

   pip install .
   python -m build
   python setup.py build_ext --inplace
