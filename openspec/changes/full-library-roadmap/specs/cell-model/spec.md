## MODIFIED Requirements

### Requirement: Стиль комірки за замовчуванням

Новостворена `Cell` SHALL мати власний незалежний стиль за замовчуванням (не
спільний module-level singleton). Зміна стилю однієї комірки SHALL NOT впливати
на інші комірки чи на інші книги в тому ж процесі.

#### Scenario: Стиль не протікає між комірками й книгами

- **GIVEN** нову книгу й комірки `A1` та `B1`
- **WHEN** встановлено `A1.font = Font(bold=True)`
- **THEN** `B1.font.bold is False`
- **AND** у свіжоствореній книзі в тому ж процесі будь-яка комірка має `font.bold is False`

## ADDED Requirements

### Requirement: Вірність типів даних при round-trip

Бібліотека SHALL зберігати типи даних при round-trip: `bool` ↔ `bool` (а не
`int`), дата/час ↔ `datetime` (за датовим `number_format`), а `inlineStr` SHALL
читатися в усіх стратегіях парсингу (включно з файлами ≥50 KB).

#### Scenario: Boolean round-trip зберігає тип

- **GIVEN** нову книгу, де `ws['A1'] = True`
- **WHEN** книгу збережено й перечитано
- **THEN** `ws['A1'].value is True` (тип `bool`, не `int`)

#### Scenario: Дата round-trip повертає datetime

- **GIVEN** нову книгу, де `ws['A1']` — `datetime(2025, 1, 2)` з датовим форматом
- **WHEN** книгу збережено й перечитано
- **THEN** `ws['A1'].value` має тип `datetime` зі значенням `2025-01-02`
