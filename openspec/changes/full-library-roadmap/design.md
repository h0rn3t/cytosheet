## Context

cytosheet (0.1.0, PoC) — Cython+lxml бібліотека для XLSX з openpyxl-сумісним API.
Поточний стан і архітектура — у `openspec/project.md`; підтверджені дефекти —
у `CLAUDE.md §10` (D-1…D-12). Ключова проблема: рушій стилів реалізований лише
для `number_format`, а збереження завантажених книг втрачає зміни. Ця зміна
задає курс до повноцінної бібліотеки, починаючи з коректності й вірності стилів.

## Goals / Non-Goals

**Goals:**
- Відкриття реального стильованого файлу (`atko_extended.xlsx`) зберігає й
  експонує стилі через API так само, як openpyxl.
- `Відкрити → змінити → зберегти` не втрачає ні змін, ні стилів інших комірок.
- Коректний мапінг листів для будь-якої кількості листів і порядку.
- Вірність типів (`bool`/дати/`inlineStr`), ізоляція стилю комірки.
- Кожен фікс підтверджено тестом, що відтворює відповідний дефект (D-x).

**Non-Goals (у межах цієї зміни / delta-specs):**
- Власний рушій обчислення формул.
- Charts/images/data-validation/conditional-formatting (Фаза 4 — окремі зміни).
- CI/wheels/docs (Фаза 5 — окрема зміна).

## Decisions

### D1. Повний реєстр стилів (read + write) — `styles.pyx`/`workbook.pyx`
Замість лише `numFmt` парсити та генерувати повний `styles.xml`:
- **Read:** з `styles.xml` будувати індексовані таблиці `fonts`, `fills`, `borders`,
  `numFmts` і `cellXfs` (xf → fontId/fillId/borderId/numFmtId/alignment). Для кожної
  комірки за її `s=` (xfId) збирати повний `Style` (font/fill/border/alignment/protection/numberFormat).
- **Write:** під час `save` обходити комірки, дедуплікувати кожен компонент у
  відповідний реєстр (font/fill/border/numFmt), формувати `cellXfs` з реальними
  посиланнями; `cell._style_id` = індекс xf. Зберегти мінімальний коректний
  `styles.xml`, який читає openpyxl.
- **Trade-off:** складніший серіалізатор; обираємо повну модель замість «passthrough
  лише оригіналу», бо інакше неможливо стилізувати нові/змінені комірки.

### D2. Збереження змін завантаженої книги — `worksheet.pyx`/`cell.pyx`
- Усі мутатори (`__setitem__`, `merge_cells`/`unmerge_cells`, проксі-сеттери стилів,
  `number_format`, контейнери `row/column_dimensions`) ставлять `_modified=True`.
- `Cell.set_value`/`__setitem__` скидають `_shared_string_index = -1`, щоб
  серіалізатор не відновлював старий рядок (D-4).
- `get_xml_data` повертає оригінальний XML лише коли лист справді не змінювався.
- **Альтернатива (відхилено):** завжди регенерувати XML — простіше, але втрачає
  байтову точність незмінених листів і частини OOXML, які ми не моделюємо.

### D3. Мапінг листів через relationships — `workbook.pyx::_load_from_archive`
- Читати порядок листів із `xl/workbook.xml` (`<sheet name r:id>`), резолвити
  `r:id` → target через `xl/_rels/workbook.xml.rels`, відкривати саме цей
  `worksheets/sheetN.xml`. Прибрати лексикографічний `sort()` і присвоєння імені
  за позицією. Це усуває D-5 для будь-якої кількості/порядку листів.

### D4. Ізоляція стилю комірки — `cell.pyx`
- Нова комірка отримує **власний** `Style()` (copy-on-write від `DEFAULT_STYLE`),
  а не спільний singleton. Усуває глобальне «протікання» (D-6). Зважити вартість:
  власний `Style` на кожну комірку дорожчий за памʼяттю — для read-only/streaming
  застосовувати lazy-створення стилю лише за наявності `s=` або звернення до проксі.

### D5. Вірність типів — `worksheet.pyx` парсери
- `t="b"` → `bool`; `t="inlineStr"` обробляти у standard/chunked парсерах також;
  дати: якщо `numberFormat` є датовим — конвертувати serial → `datetime`
  (системи 1900/1904), а запис `datetime` → serial + формат.

### D6. Справжній streaming — `worksheet.pyx::iter_rows/iter_cols`
- У read-only режимі ітерувати генератором поверх `iterparse`, не створюючи Cell
  для порожніх координат і не заповнюючи `_cells` цілком (усуває D-10).

## Test strategy (ключова вимога)

Перевірка вірності стилів — **fixture-тестами на реальному файлі**
`tests/fixtures/atko_extended.xlsx` (перенести з кореня репозиторію). openpyxl —
джерело ground truth.

1. **`test_atko_styles_read_matches_openpyxl`** — відкрити fixture через cytosheet
   та openpyxl; для набору стильованих комірок (напр. `A1`, `H1`) звірити
   `cell.font.name`, `font.bold`, `font.sz`, `fill.fgColor.rgb` (де колір — rgb),
   `alignment`. Доводить виправлення D-1.
2. **`test_atko_styles_preserved_on_modify_save`** — відкрити fixture, змінити
   `ws['A1']` (індексатор) і ще одну комірку через `ws.cell(...)`, зберегти,
   перечитати через openpyxl: нові значення застосовані (D-3, D-4), а стилі
   нечіпаних комірок (`H1` font.bold, `fill.fgColor.rgb`) збережені.
3. **`test_styles_xml_roundtrip_values`** — нова книга: задати `Font(bold, name)`,
   `PatternFill`, `Alignment`; зберегти; openpyxl читає саме ці значення (D-2).
4. **`test_multisheet_data_mapping`** — 11 листів з маркером у `A1`; cytosheet
   `wb[name]['A1'] == name` для всіх (D-5).
5. **`test_cell_style_isolation`** — `A1.font=bold` не впливає на `B1` і на свіжу
   книгу в тому ж процесі (D-6).
6. **`test_type_fidelity_roundtrip`** — `bool`↔`bool`, `datetime`↔`datetime` (D-7, D-8).

Усі тести — у дискретних файлах у `tests/`, з `tmp_path` для виходів; cross-перевірка
обома бібліотеками. Кожен тест явно прив'язаний до дефекту, який він закриває.

## Risks / Trade-offs

- **C-рівневий layout `cdef class`:** зміна полів `Cell`/`Style` потребує перезбірки
  (`build_ext --inplace`) і уваги до сумісності pickling/атрибутів.
- **Формат `styles.xml`:** нова генерація має лишатися сумісною з openpyxl/Excel —
  тому cross-тести openpyxl обовʼязкові.
- **Памʼять при ізоляції стилів:** власний `Style` на комірку дорожчий — мітигуємо
  lazy-створенням і дедуплікацією при записі.
- **Обсяг:** зміна велика; реалізовувати фазами (0→1→2→…), кожна фаза зелена в
  тестах перед наступною. Фази 3–5 виносяться в окремі зміни зі своїми specs.
