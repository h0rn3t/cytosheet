## Why

cytosheet позиціонується як drop-in заміна openpyxl, але емпірична діагностика
(cytosheet проти openpyxl на реальному `atko_extended.xlsx` та синтетичних файлах)
виявила критичні дефекти, що роблять неможливим основний сценарій «відкрити
існуючий стильований файл → змінити → зберегти»:

- стилі (шрифти/заливки/рамки/вирівнювання) **не читаються** зі `styles.xml` (D-1)
  і **не пишуться** для нових книг (D-2) — парситься/серіалізується лише `number_format`;
- зміни на завантаженій книзі **мовчки губляться**: `ws['A1']=x` не позначає лист
  зміненим (D-3), а редагування завантаженої текстової комірки повторно емітить
  старий `sharedString` індекс (D-4);
- файли з **≥10 листами** повертають дані не тих листів (D-5) через лексикографічне
  сортування й позиційний `sheetId` замість мапінгу через relationships;
- стилі **«протікають»** між комірками й навіть книгами через спільний
  module-level `DEFAULT_STYLE` (D-6);
- типи даних спотворюються: `bool`→`int` (D-7), дати лишаються числами (D-8),
  `inlineStr` губиться у файлах ≥50 KB (D-9);
- «streaming/lazy» несправжній — `iter_rows` матеріалізує Cell на кожну координату (D-10).

Щоб cytosheet став повноцінною бібліотекою, придатною до впровадження в
Python-екосистемі та як практичний внесок PhD-дисертації, потрібен послідовний
курс: спершу **коректність і вірність стилів/даних**, далі продуктивність і
широта можливостей, зрештою інженерна якість.

## What Changes

Зміна охоплює весь курс розвитку (фази 0–5 + наукова новизна). Поведінкові
контракти, що **специфікуються формально в цій зміні** (delta-specs нижче),
покривають фази 0–2 — найкритичніше й одразу тестоване. Фази 3–5 та наукова
новизна включені до плану (`design.md`, `tasks.md`) і отримають власні
delta-specs у наступних змінах, коли будуть детально проскоповані.

- **Фаза 0 — Коректність (блокери):** усі мутатори позначають лист зміненим;
  `set_value`/`__setitem__` скидають `_shared_string_index`; мапінг листів через
  `workbook.xml` + `workbook.xml.rels`; кожна комірка отримує власний `Style`.
- **Фаза 1 — Повний рушій стилів:** читання `fonts/fills/borders/alignment/protection`
  зі `styles.xml` у `Style` на комірку; генерація повного дедуплікованого
  `styles.xml` при записі; round-trip **значень** стилів обома напрямками.
- **Фаза 2 — Вірність типів:** `bool`, дати (`serial↔datetime`), `inlineStr` у всіх
  парсерах, значення-помилки.
- **Фаза 3 — Памʼять/продуктивність (теза дисертації):** справжній streaming
  read-only без матеріалізації; write-only режим; відтворювані бенчмарки.
- **Фаза 4 — Широта можливостей:** `freeze_panes`, insert/delete rows/cols,
  `copy_worksheet`, data validation, conditional formatting, comments, named ranges,
  charts, images.
- **Фаза 5 — Інженерія:** покриття >90%, corpus реальних `.xlsx`, CI + wheels,
  `.pyi`/`.pxd`, документація, формальна метрика сумісності.

**Тестова стратегія (ключова вимога):** перевірка вірності стилів виконується
fixture-тестами на реальному файлі `tests/fixtures/atko_extended.xlsx` — відкриття
через cytosheet, порівняння прочитаних стилів зі станом у `styles.xml` (ground
truth — openpyxl), та round-trip `відкрити → змінити → зберегти → перечитати` з
перевіркою, що стилі інших комірок збережені, а зміни застосовані. Деталі — у `design.md`.

## Capabilities

### New Capabilities

(У цій зміні нових capability-специфікацій немає. Capability-и фаз 3–5 — charts,
images, data-validation, conditional-formatting, comments, named-ranges,
freeze-panes, row-column-insertion, performance-benchmarking — будуть
проспецифіковані окремими змінами під час відповідних фаз.)

### Modified Capabilities

- `cell-styling`: повне читання стилів зі `styles.xml` та повна дедуплікована генерація `styles.xml` з round-trip значень.
- `xlsx-io`: коректний мапінг листів через relationships; збереження змін і збереження стилів завантаженої книги при `save`.
- `cell-model`: вірність типів даних (`bool`/дати/`inlineStr`) та ізоляція стилю кожної комірки.
- `worksheet-operations`: відстеження модифікацій листа й справжнє потокове читання без матеріалізації.

## Impact

- **Код:** `src/cytosheet/styles.pyx` (повна модель + парсинг/серіалізація реєстру стилів), `workbook.pyx` (`_parse_styles`, `_collect_styles`, `_get_styles_xml`, `_load_from_archive` мапінг), `worksheet.pyx` (`__setitem__`/`merge_cells`/контейнери — `_modified`; парсери — типи/`inlineStr`; `iter_rows/iter_cols` — streaming), `cell.pyx` (ізоляція `Style`, скидання `_shared_string_index`).
- **Тести:** `tests/fixtures/atko_extended.xlsx` (перенести з кореня) + нові round-trip і cross-openpyxl тести.
- **Specs:** оновлення living-specs `cell-styling`, `xlsx-io`, `cell-model`, `worksheet-operations` після archive.
- **Збірка/CI (Фаза 5):** `cibuildwheel`, `ruff`/`mypy`, публікація.
- **Ризики:** зміна C-рівневого layout `cdef class Cell`/`Style` (потребує перезбірки й уваги до сумісності); зміна формату згенерованого `styles.xml` (перевіряти openpyxl-сумісність).
