## ADDED Requirements

### Requirement: Відстеження модифікацій листа

Будь-яка мутація завантаженого листа SHALL позначати його як змінений, щоб
наступний `save` відображав зміну: присвоєння через `ws[coord] = value`,
`merge_cells`/`unmerge_cells`, зміна стилів/`number_format`, зміна
`row_dimensions`/`column_dimensions`. Зміна значення комірки, що походить із
shared strings, SHALL скидати її посилання на shared string.

#### Scenario: Зміна через індексатор на завантаженій книзі зберігається

- **GIVEN** fixture `tests/fixtures/atko_extended.xlsx`, відкритий через cytosheet
- **WHEN** виконано `ws['A1'] = 'CHANGED'`, книгу збережено й перечитано через cytosheet
- **THEN** `ws['A1'].value == 'CHANGED'`

### Requirement: Потокове читання без повної матеріалізації

У режимі read-only/lazy `iter_rows`/`iter_cols` SHALL віддавати значення
генератором поверх потокового парсингу, не створюючи обʼєкти `Cell` для порожніх
координат і не заповнюючи внутрішнє сховище комірок цілою сіткою.

#### Scenario: Порожні координати не матеріалізуються при потоковому читанні

- **GIVEN** великий файл, відкритий через `load_workbook(path, read_only=True)`
- **WHEN** повністю проітеровано `ws.iter_rows(values_only=True)`
- **THEN** ніколи не звернена порожня координата (напр. `ZZ9999`) відсутня у внутрішньому сховищі комірок листа
