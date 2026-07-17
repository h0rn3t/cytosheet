# Cell Styling Specification

## Purpose

Визначає модель стилів та форматування комірок: обʼєкти стилів (`Color`, `Side`,
`Border`, `Font`, `PatternFill`, `Alignment`, `Protection`, `Style`,
`DEFAULT_STYLE`), доступ до них через openpyxl-сумісні проксі комірки
(`cell.font` тощо) та числовий формат (`number_format`) з його round-trip через
`styles.xml`. API повторює openpyxl.
## Requirements
### Requirement: Обʼєкти моделі стилів

Бібліотека SHALL надавати обʼєкти стилів із полями, сумісними з openpyxl:
`Color` (`rgb`, `indexed`, `auto`, `theme`, `tint`); `Side` (`style`, `color`);
`Border` (`left`/`right`/`top`/`bottom`/`diagonal` + прапорці), де відсутні
сторони ініціалізуються порожнім `Side`; `Font` (`name`, `size`, `bold`,
`italic`, `underline`, `strike`, `color`); `PatternFill` (`patternType`,
`fgColor`, `bgColor`); `Alignment` (`horizontal`, `vertical`, `text_rotation`,
`wrap_text`, `shrink_to_fit`, `indent`); `Protection` (`locked`, `hidden`); та
складений `Style`, що обʼєднує їх плюс `numberFormat`. `DEFAULT_STYLE` SHALL бути
готовим екземпляром `Style` зі значеннями за замовчуванням.

#### Scenario: Складений Style зберігає вкладені атрибути

- **GIVEN** `Style`, зібраний із `Font(name='Arial', size=12, bold=True)`, `Border(left=Side(style='thin', color=Color(rgb='FF0000')))`, `PatternFill(patternType='solid', fgColor=Color(rgb='FFFF00'))` та `Alignment(horizontal='center', vertical='center')`
- **WHEN** його застосовано як `ws['A1'].style`
- **THEN** `ws['A1'].style.font.name == 'Arial'`, `ws['A1'].style.border.left.color.rgb == 'FF0000'`, `ws['A1'].style.fill.fgColor.rgb == 'FFFF00'`, `ws['A1'].style.alignment.horizontal == 'center'`

### Requirement: Комбінація рамок через додавання

`Border.__add__` SHALL повертати нову рамку, у якій кожна сторона береться з
правого операнда, якщо в нього задано `style` для цієї сторони, інакше — з лівого.

#### Scenario: Права рамка має пріоритет для заданих сторін

- **GIVEN** ліву рамку зі стилем `left` і праву рамку зі стилем `right`
- **WHEN** обчислено `left_border + right_border`
- **THEN** результат містить `left` сторону з лівого операнда та `right` сторону з правого

### Requirement: Проксі стилів комірки

`Cell` SHALL надавати openpyxl-сумісні властивості-проксі `font`, `border`,
`fill`, `alignment`, `protection`, що читають і записують відповідні поля
`cell.style`. Присвоєння через проксі SHALL зберігати переданий екземпляр
(`cell.font is font` після `cell.font = font`).

#### Scenario: Присвоєння font через проксі

- **GIVEN** нову комірку `ws['A1']` зі стилем за замовчуванням
- **WHEN** виконано `cell.font = Font(name='Arial', size=12, bold=True)`
- **THEN** `cell.font is font` і `cell.style.font is font`

#### Scenario: Стилі через проксі переживають round-trip у файл

- **GIVEN** комірку зі значенням і застосованими `font`/`alignment`/`border`/`fill`/`protection`, збережену у файл
- **WHEN** файл відкрито через openpyxl та через cytosheet
- **THEN** обидві бібліотеки читають значення комірки й бачать ненульові обʼєкти стилів

### Requirement: Числовий формат комірки

`Cell.number_format` SHALL бути проксі до `cell.style.numberFormat`: читання
повертає поточний рядок формату, запис створює/оновлює `Style` за потреби.

#### Scenario: number_format узгоджений зі style.numberFormat

- **WHEN** виконано `cell.number_format = '0.00%'`
- **THEN** `cell.number_format == '0.00%'` і `cell.style.numberFormat == '0.00%'`

### Requirement: Round-trip числового формату через styles.xml

При збереженні книги, створеної з нуля, користувацькі числові формати SHALL
реєструватися в `numFmts`/`cellXfs` файлу `styles.xml` (користувацькі id від 164)
так, щоб openpyxl читав той самий рядок формату. При читанні файлу формати SHALL
відновлюватися з `styles.xml` (включно з вбудованими форматами на кшталт
`'0.00'`, `'0.00%'`, `'#,##0'`, `'m/d/yy'`, `'m/d/yy h:mm'`) і призначатися
комірці як `number_format`. Згенерований `styles.xml` SHALL читатися openpyxl без
помилок.

#### Scenario: openpyxl бачить формат, записаний cytosheet

- **GIVEN** книгу cytosheet, де `ws['A1'].value = 0.5` і `ws['A1'].number_format = '0.00%'`, збережену у файл
- **WHEN** файл відкрито через `openpyxl.load_workbook`
- **THEN** `openpyxl_ws['A1'].number_format == '0.00%'`

#### Scenario: cytosheet читає формат, записаний openpyxl

- **GIVEN** файл, створений openpyxl, де `A1 = 0.25` з форматом `'0.00%'`
- **WHEN** файл відкрито через cytosheet `load_workbook`
- **THEN** `ws['A1'].value == 0.25` і `ws['A1'].number_format == '0.00%'`

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

#### Scenario: Нові стилі зливаються у завантажену книгу без втрати старих

- **GIVEN** завантажений стильований файл, до якого додано нові комірки з власними стилями (напр. червона заливка на рядку 6, зокрема й на комірці, що вже мала `s=`)
- **WHEN** книгу збережено
- **THEN** нові компоненти стилів дописуються в кінець таблиць `styles.xml`, а всі оригінальні елементи та їхні індекси зберігаються
- **AND** openpyxl читає нові комірки з їхніми стилями, а стилі нечіпаних комірок незмінні (`H1` → `bold`, `FFBDD7EE`)
- **AND** якщо змінено лише значення (без нових стилів), `styles.xml` лишається байт-у-байт ідентичним оригіналу

