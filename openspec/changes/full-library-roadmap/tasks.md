## 1. Підготовка та фікстури

- [x] 1.1 Fixture `tests/fixtures/atko_extended.xlsx` (скопійовано з кореня; стабільна, тільки читання)
- [x] 1.2 Характеризаційний `xfail`-тест Phase 1 (D-1, `test_atko_cell_styles_read_matches_openpyxl`); решта D-* покрито позитивними тестами Phase 0
- [x] 1.3 Перезбірка: `python setup.py build_ext --inplace`
- [x] 1.4 Перевстановлено пакет як editable (`pip install -e .`): `import cytosheet` тепер бере `src/`, а не застарілу копію в site-packages (через неї тести бачили старий білд)

## 2. Фаза 0 — Коректність (блокери) ✅

- [x] 2.1 (D-3) `worksheet.pyx`: `__setitem__`, `merge_cells`/`unmerge_cells`, контейнери `row/column_dimensions` ставлять `_modified=True`; у `cell.pyx` style-проксі та `number_format` позначають батьківський лист зміненим
- [x] 2.2 (D-4) `cell.pyx::set_value` і `worksheet.__setitem__` скидають `_shared_string_index = -1` при зміні значення
- [x] 2.3 (D-5) `workbook.pyx::_load_from_archive`: мапінг через `xl/workbook.xml` + `xl/_rels/workbook.xml.rels` (r:id→target); прибрано лексикографічний `sort()`; fallback — числове сортування + імена за позицією
- [x] 2.4 (D-6) `cell.pyx`: кожна комірка отримує власний `Style()` замість спільного `DEFAULT_STYLE`
- [x] 2.5 Тести `test_atko_modify_persists_and_keeps_other_styles` + `test_atko_styles_preserved_in_styles_xml_on_save` (відкрити→змінити→зберегти→перечитати; зміни застосовані, стилі інших комірок збережені)
- [x] 2.6 Тест `test_multisheet_data_mapping` (11 листів, `wb[name]['A1'] == name`)
- [x] 2.7 Тест `test_cell_style_isolation_within_and_across_workbooks`

## 3. Фаза 1 — Повний рушій стилів (read + write) ✅

- [x] 3.1 `styles.pyx`/`workbook.pyx::_parse_styles`: парсити `fonts/fills/borders/cellXfs/alignment/protection` у індексовані таблиці. У `styles.pyx` додано `*_from_element()` фабрики + `_to_xml()`/`copy()` на класах стилів; `_parse_styles` будує `_xf_style_map` (xfId → повний `Style`)
- [x] 3.2 `worksheet.pyx` парсери: для кожної комірки збирати повний `Style` за її `s=` (xfId). Додано helper-и `_style_for_xf()`/`_apply_style_simple()`; усі три парсери (simple/standard/chunked) застосовують копію `Style` замість лише `numberFormat`
- [x] 3.3 `workbook.pyx::_collect_styles`/`_get_styles_xml`: дедуплікований реєстр font/fill/border/numFmt (ключ = XML-фрагмент) + `cellXfs` з реальними посиланнями та apply*-прапорцями; зарезервовано fillId 0=none/1=gray125
- [x] 3.3a Merge нових стилів у завантажену книгу (`_merge_loaded_styles_xml`/`_collect_new_styles`): оригінальний `styles.xml` зберігається байт-у-байт, нові/змінені стилі дописуються в кінець таблиць; зміна стилю комірки скидає `_style_id` (cell.pyx `_mark_modified`), стиль xf0 трактується як дефолт, щоб успадковані комірки не плодили дублі
- [x] 3.4 Тест `test_atko_cell_styles_read_matches_openpyxl` (font.name/bold/sz, fill.fgColor.rgb, alignment звірені з openpyxl на fixture для A1/H1/A2) — закриває D-1
- [x] 3.5 Тест `test_styles_xml_roundtrip_values` (нова книга: bold/Arial/fill/alignment → openpyxl читає ті самі значення) — закриває D-2
- [x] 3.6 Знято `xfail` з кроку 1.2 (тест тепер проходить як звичайний позитивний)

## 4. Фаза 2 — Вірність типів даних

- [x] 4.1 (D-7) `worksheet.pyx` парсери: `t="b"` → `bool` — через спільну `_value_from_t()`, яку використовують усі три парсери + потоковий шлях (диспетчеризація за `t=` була розповзлась по трьох копіях)
- [x] 4.2 (D-8) Дати: `serial → datetime` за датовим `number_format` (`_is_date_format`/`_from_excel`/`_to_excel` дзеркалять openpyxl, включно з міфічним 1900-02-29); `date1904` читається з `<workbookPr>` і передається в лист; запис `datetime → serial`, формат ставить `Cell.set_value` (`_ensure_date_format`). Заразом заповнено `BUILTIN_NUMFMTS` до повної таблиці ECMA-376 — без id 21 (`h:mm:ss`) час від openpyxl читався числом
- [x] 4.3 (D-9) `inlineStr` обробляти у `_parse_sheet_standard` та `_parse_sheet_chunked` — спільний helper `_cell_from_element()` замість двох копій блоку вилучення значення (CS-7); `itertext()` склеює rich-text runs; `_stream_cell_value` вирівняно на ту саму семантику. Закриває `test_stream_no_dimension_falls_back` + `test_inlinestr_read_in_standard_parser`
- [x] 4.4 Значення-помилки `t="e"` зберігати з `data_type='e'` — гілка в `_value_from_t`; значення й `data_type` збігаються з openpyxl
- [x] 4.5 Тести — `tests/test_type_fidelity.py` (11 кейсів): bool, дати (параметризовано, з межею serial<60), час з builtin-формату, `t="e"`, XML-сутності; усі звірені з openpyxl як ground truth
- [x] 4.6 (поза початковим планом) Простий парсер не знімав XML-сутності: `'a & b < c'` читався як `'a &amp; b &lt; c'`, кирилиця — як `&#1090;...`. Тиха порча даних на дефолтному шляху малих файлів (COMPAT-1). Фікс: `html.unescape` на трьох місцях вилучення підрядка (lxml-шляхи декодують самі, тож їх НЕ чіпаємо). Знайдено крос-тестом 4.5

## 5. Фаза 3 — Памʼять і продуктивність

**Основа:** зовнішній патч `~/Downloads/cytosheet_stream_values_patch.pyx` (iterparse-ядро
з `clear()` + видаленням попередніх сиблінгів) — перевірений емпірично, але на щільному
файлі, тож розріджені входи в ньому зламані. Ядро беремо, форму — ні (див. design D6.1).

**Ескалація scope (2026-07-17):** тести 5.1 засвітили два дефекти поза межами streaming,
які його блокують: `ws['A1'].value = x` не позначає лист зміненим (D-3/D-4 крізь
публічний C-атрибут → task 5.1.13, design D8) і `inlineStr` у standard-парсері
(D-9 → task 4.3, design D9). Обидва втягнуто в цю зміну.

- [ ] 5.1 (D-10) Справжній streaming read-only
  - [x] 5.1.1 `worksheet.pyx:77`: читання `_original_xml` перевести під `_preload`; усунути подвійний `archive.read` (рядки 79 і 84 читають той самий entry)
  - [x] 5.1.2 `get_xml_data`: ліниве читання `_original_xml` з архіву **без memoize в поле** (інакше пік = сума всіх листів, бо `workbook.pyx:675` кличе `get_xml_data` на кожному листі) — `_read_original_xml()`
  - [x] 5.1.3 `workbook.pyx::save`/`save_virtual_workbook`: `ValueError`, якщо `_archive is None` після `close()` — прапорець `_closed` + `_check_open()` (лише для завантажених книг; нова книга без архіву зберігається як раніше)
  - [x] 5.1.4 `<dimension>`: парсити `ref` під час `iterparse` (передує `<sheetData>`, вартість O(1)) для `max_col` — `_peek_dimension_max_col()` на подіях `start`
  - [x] 5.1.5 `<dimension ref>` емітити в `get_xml_data` перед `<sheetData>` — зараз не пишемо взагалі, через що власні файли не стримляться
  - [x] 5.1.6 `iter_rows`/`iter_cols`: гілка `not _preloaded` → `iterparse`-генератор; позиція комірки з `r=` (через наявний `_col_to_num`), падінг пропущених комірок і рядків до `None`; `values_only=False` створює Cell на льоту без запису в `_cells` — `_stream_rows`/`_make_row`/`_make_stream_cell`
  - [x] 5.1.7 Фолбек на матеріалізуючий шлях, якщо `<dimension>` відсутній і `max_col` не передано — зелений після 5.1.9 + 4.3
  - [x] 5.1.8 Eager-шлях недоторканий: `_preloaded=True` → поточна bbox-ітерація (identity комірок = семантика запису) — `_iter_rows_materialized`; зелений після 5.1.13
  - [x] 5.1.9 `_is_small_file` рахується з `archive.getinfo(path).file_size` в обох режимах, до й поза `if _preload` — без читання entry (design D9)
  - [ ] 5.1.10 (опційно) `stream_values()` як однорядковий аліас `return self.iter_rows(values_only=True)` — тільки якщо ім'я потрібне зовні
  - [x] 5.1.11 Тести: `tests/test_streaming_readonly.py` — 10 з 10 зелені
  - [x] 5.1.12 Бенчмарк сталої памʼяті: `test_stream_peak_memory_constant_in_rows` — пік RSS у окремому процесі (ru_maxrss монотонний, тож два заміри в одному процесі непридатні), фікстури зібрані сирим XML (openpyxl на 200k рядків сам стає вузьким місцем). **Результат: 25.5 → 26.8 MB за 20× даних (+5.2%)**. Фальсифікація: той самий обхід eager-шляхом дає 86.6 → 1292 MB (+1392%), тобто поріг 20% справді дискримінує; стримінг = 48× менше памʼяті на 200k рядків
  - [x] 5.1.13 (D-3/D-4, design D8) `cell.pyx`: `value` → `@property` над полем `_value`; `set_value` — єдина реалізація запису (оновлює `data_type`, скидає `_shared_string_index=-1`, ставить `parent._modified=True`), сеттер делегує їй. Конструктор пише `self._value` напряму; парсери створюють Cell лише через конструктор, тож завантаження не позначає лист зміненим
  - [x] 5.1.14 Тести до 5.1.13: `test_cell_value_attr_marks_modified`, `test_load_and_save_unmodified_is_byte_identical` (regression на ризик D8 — байтова звірка XML усіх листів), `test_inlinestr_read_in_standard_parser` (закриває 4.3)
  - [x] 5.1.15 PERF-1 до/після 5.1.13: виміряно на `file_example_XLSX_5000.xlsx` (40 007 комірок, best-of-7): 31.8 ns/читання до → 31.4 ns після, просідання немає (`cdef public` і так дескриптор). Ризик із design D8 знято
- [ ] 5.2 Write-only режим для запису дуже великих файлів (аналог openpyxl `WriteOnlyWorksheet`)
- [ ] 5.3 Прибрати `print()`/`except Exception: pass` з парсерів (D-12): логування або контрольоване підняття
- [ ] 5.4 Відтворюваний бенчмарк-набір (час + RSS) vs openpyxl/xlsxwriter на матриці розмірів/щільності; автоматизація

## 6. Фаза 4 — Широта можливостей (окремі delta-specs у наступних змінах)

- [ ] 6.1 `freeze_panes`; `auto_filter`
- [ ] 6.2 `insert_rows`/`delete_rows`/`insert_cols`/`delete_cols` зі зсувом формул
- [ ] 6.3 `copy_worksheet`
- [ ] 6.4 data validation; conditional formatting; comments; named ranges; page setup; protection
- [ ] 6.5 charts; images (`drawing.xml`)

## 7. Фаза 5 — Інженерія та якість (окрема зміна)

- [ ] 7.1 Покриття тестами >90%; corpus реальних `.xlsx` (Excel/LibreOffice/Google Sheets/openpyxl) для round-trip і fuzz
- [ ] 7.2 CI (GitHub Actions): `cibuildwheel` (Linux/macOS/Windows, py3.8–3.13), `ruff`/`mypy`, публікація на PyPI
- [ ] 7.3 `.pxd` + `.pyi` стаби; документація (Sphinx) + migration guide від openpyxl
- [ ] 7.4 Автогенерована формальна метрика drop-in-сумісності з openpyxl

## 8. Наукова новизна (PhD)

- [ ] 8.1 Адаптивний вибір стратегії парсингу за профілем файлу (розмір/щільність/shared strings) замість фіксованих порогів 50/100 KB; обґрунтування моделлю
- [ ] 8.2 Кількісна модель «памʼять × час» Cython+libxml2 vs pure-Python (де й чому виграш)
- [ ] 8.3 Методологія вимірювання сумісності та продуктивності як відтворюваний дослідницький артефакт

## 9. Закриття

- [ ] 9.1 `openspec validate --strict full-library-roadmap` зелений
- [ ] 9.2 Усі тести фаз 0–2 проходять; немає регресій у наявних тестах
- [ ] 9.3 Review summary; `openspec archive full-library-roadmap` (оновлює living specs)
