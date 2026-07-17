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

Під час завантаження бібліотека SHALL визначати листи за порядком у
`xl/workbook.xml` (`<sheet name r:id>`) і резолвити кожен `r:id` у цільовий
`xl/worksheets/sheetN.xml` через `xl/_rels/workbook.xml.rels`. Бібліотека SHALL
NOT покладатися на лексикографічний порядок імен файлів чи на позиційне
присвоєння `sheetId`. Таблиця `xl/sharedStrings.xml`, якщо є, SHALL парситися для
відновлення рядкових значень.

#### Scenario: Дані кожного листа відповідають його імені для 10+ листів

- **GIVEN** файл з 11 листами (`First`, `S2`…`S11`), де у кожного в `A1` записано його власне імʼя
- **WHEN** файл відкрито через cytosheet
- **THEN** для кожного імені `wb[name]['A1'].value == name` (жодного перемішування даних)

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

### Requirement: Збереження змін і стилів завантаженої книги при записі

`save` SHALL відображати всі зміни, внесені у завантажену книгу (значення через
індексатор чи `cell()`, стилі, обʼєднання, розміри), і одночасно зберігати стилі
комірок, які не змінювалися.

#### Scenario: Зміна значень застосовується, стилі інших комірок збережені

- **GIVEN** fixture `tests/fixtures/atko_extended.xlsx`, відкритий через cytosheet
- **WHEN** змінено `ws['A1']` (індексатор) та ще одну комірку через `ws.cell(...)`, книгу збережено й відкрито через openpyxl
- **THEN** openpyxl бачить нові значення у змінених комірках
- **AND** нечіпані стильовані комірки (напр. `H1`) зберігають `font.bold` та `fill.fgColor.rgb`

### Requirement: Lazy-режим не читає XML листа при відкритті

У режимі `read_only=True`/`lazy=True` конструктор листа SHALL NOT читати XML листа
з архіву — ані для парсингу, ані для збереження оригіналу. Той самий запис архіву
SHALL NOT читатися двічі за одне відкриття.

#### Scenario: Відкриття в lazy-режимі не звертається до XML листа

- **GIVEN** архів, у якого метод `read` обгорнуто лічильником звернень
- **WHEN** виконано `load_workbook(path, read_only=True)` без подальших операцій
- **THEN** `read` жодного разу не викликано для `xl/worksheets/sheet1.xml`

#### Scenario: Звичайне відкриття читає XML листа рівно один раз

- **GIVEN** архів, у якого метод `read` обгорнуто лічильником звернень
- **WHEN** виконано `load_workbook(path)` (без `read_only`)
- **THEN** `read` викликано для `xl/worksheets/sheet1.xml` рівно один раз

### Requirement: Збереження книги, відкритої в read-only/lazy режимі

`save` на книзі, відкритій з `read_only=True`/`lazy=True`, SHALL створювати
коректний файл (на відміну від openpyxl, який зберігати read-only книгу не
дозволяє). Оригінальний XML незмінених листів SHALL читатися на вимогу під час
`save` і SHALL NOT кешуватися у полі листа, щоб пікова памʼять збереження
лишалася в межах одного найбільшого листа, а не суми всіх листів книги.

#### Scenario: lazy без змін → save зберігає дані всіх листів

- **GIVEN** файл із трьома листами з даними, відкритий через `load_workbook(path, read_only=True)`
- **WHEN** виконано `wb.save(out)` без жодних змін
- **THEN** openpyxl читає з `out` ті самі значення всіх трьох листів, що й з оригіналу

#### Scenario: lazy зі зміною → save застосовує зміну й зберігає решту

- **GIVEN** файл із кількома листами, відкритий через `load_workbook(path, read_only=True)`
- **WHEN** виконано `ws['A1'] = 'CHANGED'` на першому листі та `wb.save(out)`
- **THEN** openpyxl бачить `'CHANGED'` у `A1` першого листа
- **AND** дані інших листів збігаються з оригіналом

### Requirement: Збереження після закриття архіву

`Workbook.save` після `Workbook.close` SHALL підіймати `ValueError`, а не
створювати файл із порожніми листами.

#### Scenario: save після close підіймає ValueError

- **GIVEN** книга, завантажена через `load_workbook(path)`, на якій викликано `wb.close()`
- **WHEN** виконано `wb.save(out)`
- **THEN** підіймається `ValueError`

### Requirement: Емісія `<dimension>` у згенерованому XML листа

Згенерований XML листа SHALL містити елемент `<dimension ref="...">` з фактичними
межами даних перед `<sheetData>`. Це вимагається Excel і робить згенеровані
cytosheet файли придатними до потокового читання без фолбеку.

#### Scenario: Записаний cytosheet файл стримиться без фолбеку

- **GIVEN** нова книга cytosheet з даними в діапазоні `A1:C3`, збережена у файл
- **WHEN** файл відкрито через `load_workbook(out, read_only=True)` і проітеровано `ws.iter_rows(values_only=True)`
- **THEN** XML листа містить `<dimension ref="A1:C3"/>`
- **AND** усі отримані кортежі мають ширину 3

