# Row and Column Dimensions Specification

## Purpose

Визначає роботу з розмірами рядків і колонок: висота й видимість рядків
(`row_dimensions`), ширина й видимість колонок (`column_dimensions`), їх читання
з файлу, запис у XML та взаємний round-trip із openpyxl. API повторює openpyxl.

## Requirements

### Requirement: Доступ до розмірів колонок

`ws.column_dimensions[letter]` SHALL повертати обʼєкт `ColumnDimension` для
літери колонки (`'A'`), створюючи його за потреби, з полями `width` та `hidden`,
доступними для читання й запису. Контейнер SHALL поводитися як відображення
(`MutableMapping`).

#### Scenario: Встановлення й читання ширини колонки

- **WHEN** виконано `ws.column_dimensions['A'].width = 20`
- **THEN** `ws.column_dimensions['A'].width == 20`

### Requirement: Доступ до розмірів рядків

`ws.row_dimensions[index]` SHALL повертати обʼєкт `RowDimension` для номера рядка,
створюючи його за потреби, з полями `height` та `hidden`, доступними для читання
й запису.

#### Scenario: Встановлення висоти й прихованості рядка

- **WHEN** виконано `ws.row_dimensions[1].height = 25.0` і `ws.row_dimensions[2].hidden = True`
- **THEN** `ws.row_dimensions[1].height == 25.0` і `ws.row_dimensions[2].hidden is True`

### Requirement: Round-trip розмірів через cytosheet

Записані ширини колонок та висоти/прихованість рядків SHALL серіалізуватися у
`sheetN.xml` (`<col ... width customWidth>`, `<row ... ht customHeight hidden>`)
і відновлюватися при повторному завантаженні тим самим API.

#### Scenario: Ширина колонок переживає збереження й завантаження

- **GIVEN** книгу, де задано ширини колонок `A`, `B`, `C`, збережену у файл
- **WHEN** файл перечитано через `load_workbook`
- **THEN** `ws.column_dimensions[letter].width` повертає ті самі значення

#### Scenario: Висота й hidden рядків переживають round-trip

- **GIVEN** книгу, де `ws.row_dimensions[1].height = 25.0` і `ws.row_dimensions[2].hidden = True`, збережену у файл
- **WHEN** файл перечитано через `load_workbook`
- **THEN** `ws.row_dimensions[1].height == 25.0` і `ws.row_dimensions[2].hidden is True`

### Requirement: Взаємна сумісність розмірів із openpyxl

Розміри, записані cytosheet, SHALL коректно читатися openpyxl, а розміри,
записані openpyxl, SHALL коректно читатися cytosheet (в обох напрямках, для
ширини колонок та висоти/прихованості рядків).

#### Scenario: openpyxl читає ширини колонок, записані cytosheet

- **GIVEN** книгу cytosheet, де `ws.column_dimensions['A'].width = 20` і `ws.column_dimensions['B'].width = 40`, збережену у файл
- **WHEN** файл відкрито через `openpyxl.load_workbook`
- **THEN** `openpyxl_ws.column_dimensions['A'].width == 20` і `openpyxl_ws.column_dimensions['B'].width == 40`

#### Scenario: cytosheet читає розміри, записані openpyxl

- **GIVEN** файл, створений openpyxl, де `column_dimensions['A'].width = 15` і `row_dimensions[1].height = 22`, `row_dimensions[2].hidden = True`
- **WHEN** файл відкрито через cytosheet `load_workbook`
- **THEN** `ws.column_dimensions['A'].width == 15`, `ws.row_dimensions[1].height == 22`, `ws.row_dimensions[2].hidden is True`
