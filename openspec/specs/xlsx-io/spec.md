# XLSX Input/Output Specification

## Purpose

Визначає ядро читання та запису формату XLSX (OOXML): функцію `load_workbook`,
стратегії парсингу листів за розміром, lazy-завантаження, збереження у файл /
file-like / байти, збереження оригінальних частин шаблону та взаємну сумісність
round-trip з openpyxl. Це «суть» бібліотеки — заради чого вона існує.

## Requirements

### Requirement: Завантаження книги з файлу

Бібліотека SHALL відкривати XLSX-архів і повертати ініціалізований `Workbook`
через `load_workbook(filename, read_only=False, keep_vba=False, data_only=False,
keep_links=True, rich_text=False, lazy=False)`. Сигнатура SHALL бути сумісною з
openpyxl: невикористовувані аргументи (`keep_vba`, `data_only`, `keep_links`,
`rich_text`) приймаються й ігноруються без помилки. `filename` SHALL приймати як
рядок, так і обʼєкт, що підтримує `os.PathLike` (через `__fspath__`, напр.
`pathlib.Path`).

#### Scenario: Завантаження через pathlib.Path

- **GIVEN** збережений файл і шлях `test_file = Path(...)`
- **WHEN** викликано `load_workbook(filename=test_file)`
- **THEN** книга завантажується і `wb.active['A1'].value` повертає збережене значення

#### Scenario: Зайві openpyxl-аргументи не спричиняють помилку

- **WHEN** викликано `load_workbook(path, data_only=True, keep_vba=False)`
- **THEN** книга завантажується успішно, аргументи проігноровано

### Requirement: Імена листів та shared strings із архіву

Під час завантаження бібліотека SHALL читати реальні назви листів із
`xl/workbook.xml` (елементи `sheet name=...`) і відображати їх на відповідні
`xl/worksheets/sheetN.xml` за порядком `sheetId`. За відсутності назви SHALL
застосовуватися fallback `"sheet{N}"`. Таблиця `xl/sharedStrings.xml`, якщо є,
SHALL парситися для відновлення рядкових значень комірок.

#### Scenario: Назва листа читається з workbook.xml

- **GIVEN** файл, створений із листом `"Sheet1"`
- **WHEN** виконано `wb = load_workbook(path)`
- **THEN** `wb.get_sheet_by_name('Sheet1').title == 'Sheet1'`

### Requirement: Вибір стратегії парсингу за розміром листа

Для оптимальної продуктивності бібліотека SHALL обирати парсер залежно від
розміру XML листа: string-парсинг для малих листів (`< 50 KB`), `lxml.iterparse`
для середніх (`< 100 KB`) та батчевий `iterparse` з `recover=True` для великих
(`> 100 KB`). Усі стратегії SHALL давати ідентичну модель комірок.

#### Scenario: Числові й текстові значення коректно типізуються при читанні

- **GIVEN** збережена книга з `A1=42`, `B1=3.14`, `C1='text'`
- **WHEN** файл перечитано через `load_workbook`
- **THEN** `ws['A1'].value == 42` (int), `ws['B1'].value == 3.14` (float), `ws['C1'].value == 'text'` (str)

### Requirement: Lazy-завантаження великих файлів

У режимі `read_only=True` або `lazy=True` листи SHALL створюватися без негайного
парсингу даних, а дані SHALL підвантажуватися при першому виклику ітератора
(`iter_rows`/`iter_cols`). Це дозволяє потокове читання без завантаження листа
цілком у памʼять.

#### Scenario: Потокове читання рядків у lazy-режимі

- **GIVEN** великий файл, відкритий через `load_workbook(path, lazy=True)`
- **WHEN** взято `rows = ws.iter_rows(values_only=True)` і прочитано два перші рядки через `next`
- **THEN** другий рядок містить очікувані дані (напр. `data[0] == 1`)

### Requirement: Збереження у файл, file-like обʼєкт та байти

`Workbook.save(target)` SHALL зберігати книгу у файл за шляхом (створюючи проміжні
каталоги) або у будь-який file-like обʼєкт, що має метод `write` (напр. `BytesIO`).
`Workbook.save_virtual_workbook()` SHALL повертати вміст XLSX як `bytes`. Порожній
шлях SHALL відхилятися через `ValueError`.

#### Scenario: Збереження у BytesIO

- **GIVEN** книга з даними в комірках
- **WHEN** виконано `wb.save(BytesIO())`
- **THEN** буфер містить непорожній XLSX-вміст (`len(getvalue()) > 0`)

### Requirement: Збереження зі збереженням оригінальних частин

Для книги, завантаженої з файлу, `save` SHALL копіювати незмінені частини архіву
без модифікації та перезаписувати лише змінені листи й службові частини
(`workbook.xml`, `[Content_Types].xml`, rels, `styles.xml`, `sharedStrings.xml`).
Лист, що не змінювався, SHALL серіалізуватися зі свого оригінального XML.

#### Scenario: Незмінений лист повертає оригінальний XML

- **GIVEN** лист, завантажений із файлу й не модифікований
- **WHEN** генерується його XML для збереження
- **THEN** повертається байтовий оригінал листа без перегенерації

### Requirement: Взаємна сумісність round-trip з openpyxl

Файли, записані cytosheet, SHALL відкриватися openpyxl без помилок із коректними
значеннями та базовими стилями. Файли, створені openpyxl, SHALL відкриватися
cytosheet із коректними значеннями.

#### Scenario: openpyxl читає файл, записаний cytosheet

- **GIVEN** книга cytosheet зі значенням і стилями в `A1`, збережена у файл
- **WHEN** файл відкрито через `openpyxl.load_workbook`
- **THEN** `openpyxl_ws['A1'].value` дорівнює записаному значенню, а обʼєкти стилів доступні
