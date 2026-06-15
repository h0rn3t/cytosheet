## ADDED Requirements

### Requirement: Повне читання стилів комірки зі styles.xml

Бібліотека SHALL читати зі `styles.xml` повний набір стилів (`fonts`, `fills`,
`borders`, `alignment`, `protection`, `numFmts`) і експонувати їх на кожній
комірці через її `s=` (xfId), а не лише `number_format`. Прочитані значення SHALL
збігатися з тим, що для тих самих комірок повертає openpyxl.

#### Scenario: Стилі реального файлу читаються як в openpyxl

- **GIVEN** fixture `tests/fixtures/atko_extended.xlsx`, відкритий і через cytosheet, і через openpyxl
- **WHEN** прочитано стильовані комірки (напр. `A1`, `H1`)
- **THEN** `cyto.font.name`, `cyto.font.bold`, `cyto.font.sz` збігаються з openpyxl
- **AND** `cyto.alignment.horizontal` збігається з openpyxl
- **AND** для комірок із rgb-заливкою `cyto.fill.fgColor.rgb` збігається з openpyxl (напр. `H1` → `FFBDD7EE`)

#### Scenario: Комірка без явного стилю успадковує дефолтний стиль книги

- **GIVEN** завантажений файл, у якому комірка не має атрибута `s=`
- **WHEN** прочитано `cell.font` цієї комірки
- **THEN** повертається шрифт за `fontId=0` (дефолтний стиль книги), як в openpyxl
- **AND** для нової книги або файлу без `styles.xml` дефолтний шрифт збігається з openpyxl (`Calibri`, `11`)

### Requirement: Повна дедуплікована генерація styles.xml

Під час `save` бібліотека SHALL серіалізувати всі унікальні компоненти стилів
(`fonts`/`fills`/`borders`/`numFmts`) у дедуплікований `styles.xml` і будувати
`cellXfs` з реальними посиланнями, щоб значення стилів (bold, імʼя шрифту, колір
заливки, вирівнювання) round-trip-ились, а не лише `number_format`. Згенерований
`styles.xml` SHALL читатися openpyxl без помилок.

#### Scenario: Візуальні стилі нової книги читаються openpyxl

- **GIVEN** нову книгу, де `ws['A1']` має `Font(name='Arial', bold=True)`, `PatternFill(patternType='solid', fgColor=Color(rgb='FFFFFF00'))` та `Alignment(horizontal='center')`
- **WHEN** книгу збережено й відкрито через openpyxl
- **THEN** openpyxl читає `font.name == 'Arial'`, `font.bold is True`, жовту заливку та `alignment.horizontal == 'center'`
