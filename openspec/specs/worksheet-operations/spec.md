# Worksheet Operations Specification

## Purpose

Визначає доступ до даних листа та операції над сіткою комірок: індексований
доступ (`ws['A1']`, діапазони, доступ за номером рядка), координатний доступ
(`ws.cell`), додавання рядків (`append`), межі використаного діапазону
(`max_row`/`max_column`/`dimensions`) та ітерацію (`iter_rows`/`iter_cols` і
властивості `rows`/`columns`/`values`). API повторює openpyxl.

## Requirements

### Requirement: Індексований доступ до однієї комірки

`ws[coord]` SHALL повертати комірку за координатою `"A1"`, створюючи її за потреби.
`ws[coord] = value` SHALL встановлювати значення комірки; якщо значення — рядок,
що починається з `'='`, тип даних комірки SHALL стати `'f'` (формула). Доступ і
запис SHALL оновлювати межі використаного діапазону.

#### Scenario: Створення та читання комірки за координатою

- **WHEN** виконано `ws['A1'] = 'Hello'` і `ws['B2'] = 'World'`
- **THEN** `ws['A1'].value == 'Hello'` і `ws['B2'].value == 'World'`

### Requirement: Доступ до діапазону комірок

`ws[range_string]` для діапазону виду `"A1:C3"` SHALL повертати кортеж кортежів
комірок (рядки × колонки), створюючи відсутні комірки.

#### Scenario: ws['A1:C3'] повертає сітку 3×3

- **GIVEN** лист, заповнений значеннями `R{r}C{c}` у `A1:C3`
- **WHEN** виконано `rng = ws['A1:C3']`
- **THEN** `len(rng) == 3` і `len(rng[0]) == 3`
- **AND** `rng[0][0].value == 'R1C1'` і `rng[2][2].value == 'R3C3'`

### Requirement: Доступ до рядка за номером

`ws[n]`, де `n` — ціле число, SHALL повертати комірки рядка з номером `n`
(сумісно з openpyxl `ws[1]`), реалізовано через `iter_rows(min_row=n, max_row=n)`.

#### Scenario: Отримання заголовків через ws[1]

- **GIVEN** лист зі значеннями в `A1`, `B1`, `C1`
- **WHEN** виконано `[cell.value for cell in ws[1]]`
- **THEN** повертається список значень першого рядка

### Requirement: Координатний доступ через ws.cell

`ws.cell(row, column, value=None)` SHALL повертати комірку за 1-базованими
координатами, створюючи її за потреби, і встановлювати `value`, якщо він заданий
(позначаючи лист як змінений). `row < 1` або `column < 1` SHALL відхилятися через
`ValueError`. Виклик SHALL оновлювати `max_row`/`max_column`.

#### Scenario: Запис значення за координатами row/column

- **WHEN** виконано `ws.cell(row=1, column=1, value="A1")` і `ws.cell(row=5, column=2, value="B5")`
- **THEN** `ws['A1'].value == "A1"` і `ws['B5'].value == "B5"`
- **AND** `ws.max_row == 5` і `ws.max_column == 2`

### Requirement: Додавання рядка через append

`ws.append(iterable)` SHALL записувати значення ітерабельного обʼєкта в новий
рядок одразу після `max_row`, починаючи з першої колонки. Рядки, байти та
неітерабельні обʼєкти SHALL трактуватися як скаляр і записуватися в першу колонку.

#### Scenario: append додає рядок після max_row

- **GIVEN** лист, де `max_row == 5`
- **WHEN** виконано `ws.append([1, 2, 3])`
- **THEN** `ws.max_row == 6`
- **AND** `ws['A6'].value == 1`, `ws['B6'].value == 2`, `ws['C6'].value == 3`

#### Scenario: append зі скаляром пише в першу колонку

- **WHEN** для порожнього листа виконано `ws.append("single")`
- **THEN** `ws['A1'].value == "single"` і `ws.max_column == 1`

### Requirement: Межі використаного діапазону

`ws.max_row` та `ws.max_column` SHALL повертати найбільші номер рядка й колонки
серед комірок з даними. `ws.dimensions` SHALL повертати рядок діапазону виду
`"B2:D5"`, що охоплює використаний прямокутник; для порожнього листа — `"A1:A1"`.
Межі SHALL коректно враховувати змішане використання `ws[coord]` і `ws.cell`.

#### Scenario: dimensions відображає використаний прямокутник

- **GIVEN** лист, де `ws['B2'] = 1` і `ws['D5'] = 2`
- **WHEN** прочитано межі
- **THEN** `ws.max_row == 5`, `ws.max_column == 4`, `ws.dimensions == 'B2:D5'`

### Requirement: Ітерація по рядках і колонках

Лист SHALL надавати ітерацію по рядках і колонках. `ws.iter_rows(min_row,
max_row, min_col, max_col, values_only=False)` та `ws.iter_cols(...)` генерують
кортежі комірок (або значень за `values_only=True`) у межах діапазону.
Властивості `ws.rows`, `ws.columns` та `ws.values` SHALL бути еквівалентні
`iter_rows()`, `iter_cols()` та `iter_rows(values_only=True)` відповідно.

#### Scenario: iter_rows та .values повертають дані по рядках

- **GIVEN** лист зі значеннями `["ID","Name"]`, `[1,"Alice"]`, `[2,"Bob"]`
- **WHEN** виконано `list(ws.values)`
- **THEN** результат `[("ID","Name"), (1,"Alice"), (2,"Bob")]`

#### Scenario: iter_cols повертає дані по колонках

- **GIVEN** той самий лист
- **WHEN** виконано `list(ws.iter_cols(min_col=1, max_col=2, min_row=1, max_row=3, values_only=True))`
- **THEN** перша колонка `("ID", 1, 2)`, друга `("Name", "Alice", "Bob")`
