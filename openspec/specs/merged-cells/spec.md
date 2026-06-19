# Merged Cells Specification

## Purpose

Визначає поведінку обʼєднання та розʼєднання діапазонів комірок (`merge_cells` /
`unmerge_cells`), сумісну з openpyxl, включно зі збереженням стану в `mergeCells`
секції `sheetN.xml` та зчитуванням його назад.

## Requirements

### Requirement: Обʼєднання діапазону комірок

`Worksheet.merge_cells(range_string)` SHALL обʼєднувати прямокутний діапазон у
форматі `"A1:B2"`, повертати верхню-ліву (anchor) комірку, помічати всі комірки
діапазону прапорцем `is_merged_cell = True` із заповненим `merged_range`, та
очищати значення всіх комірок діапазону, окрім anchor.

#### Scenario: Обʼєднання A1:B2 зберігає значення лише в anchor

- **GIVEN** новий лист, де `ws['A1'] = 'Merged Cell'`
- **WHEN** викликано `ws.merge_cells('A1:B2')`
- **THEN** повернена комірка має `is_merged_cell == True` і `merged_range == 'A1:B2'`
- **AND** `ws['A1'].value == 'Merged Cell'`, тоді як `ws['A2']`, `ws['B1']`, `ws['B2']` мають `value is None` і `is_merged_cell == True`

#### Scenario: Некоректний формат діапазону відхиляється

- **WHEN** викликано `ws.merge_cells('A1')` без двокрапки
- **THEN** піднімається `ValueError`

### Requirement: Розʼєднання діапазону комірок

`Worksheet.unmerge_cells(range_string)` SHALL знімати прапорці `is_merged_cell` та
`merged_range` з усіх комірок діапазону й видаляти діапазон зі списку обʼєднаних.
Виклик для діапазону, якого немає серед обʼєднаних, SHALL бути безпечним (no-op).

#### Scenario: Розʼєднання знімає прапорці й зберігає значення anchor

- **GIVEN** лист, де `ws['A1'] = 'Merged Cell'` і виконано `ws.merge_cells('A1:B2')`
- **WHEN** викликано `ws.unmerge_cells('A1:B2')`
- **THEN** усі з `A1`, `A2`, `B1`, `B2` мають `is_merged_cell == False`
- **AND** `ws['A1'].value == 'Merged Cell'`

### Requirement: Збереження обʼєднань при записі та читанні

Обʼєднані діапазони SHALL серіалізуватися у секцію `<mergeCells>` файлу
`sheetN.xml` і відновлюватися при повторному завантаженні через `load_workbook`,
забезпечуючи round-trip значень anchor-комірки.

#### Scenario: Round-trip обʼєднаної комірки через файл

- **GIVEN** збережена книга, де `ws['A1'] = 'Merged Cell'` і виконано `ws.merge_cells('A1:B2')`
- **WHEN** файл повторно відкрито через `load_workbook`
- **THEN** `ws['A1'].value == 'Merged Cell'`
