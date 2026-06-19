# Cytosheet — контекст проєкту

## Суть проєкту

**Cytosheet** — високопродуктивна бібліотека для читання та запису файлів `XLSX`
(Office Open XML / OOXML), написана на **Cython** поверх **lxml** (libxml2).

Головна мета — бути **drop-in заміною openpyxl** для типових табличних сценаріїв
(звіти, шаблони, ETL), але швидшою та економнішою до памʼяті за рахунок:

- компіляції критичного коду в C через Cython (`*.pyx`);
- парсингу XML через libxml2 (`lxml`);
- вибору стратегії парсингу залежно від розміру листа;
- lazy-режиму (потокове читання великих файлів без завантаження листа цілком).

Публічний API навмисно повторює openpyxl: `Workbook`, `load_workbook`, `Worksheet`,
`Cell`, `Font`, `Border`, `Side`, `Color`, `PatternFill`, `Alignment`, `Protection`,
`Style`. Це дозволяє переносити наявний openpyxl-код з мінімальними правками.

Статус: **PoC / pre-1.0** (версія 0.1.0). Базові операції читання/запису, стилі,
обʼєднання комірок, розміри рядків/колонок та формули-як-рядки працюють і мають
round-trip із openpyxl. Складні можливості Excel (діаграми, зображення, валідація
даних, умовне форматування, рушій обчислення формул) ще не реалізовані.

## Цільові сценарії (для кого і навіщо)

1. **Заповнення XLSX-шаблону з БД** — відкрити готовий шаблон зі стилями/формулами,
   построково записати дані, підлаштувати ширину колонок, зберегти.
2. **Створення книги з нуля** — заголовки, стилі (`Font`/`Alignment`), висота рядків.
3. **«Копіювання» шаблонного листа** — створити новий лист, скопіювати комірки/стилі
   вручну, заповнити, видалити трафаретний лист.
4. **Потокове читання великих файлів** — `load_workbook(..., read_only=True/lazy=True)`
   + `iter_rows(values_only=True)`.

## Технологічний стек

- **Мова**: Python 3.7+ та Cython `>= 3.0.11` (модулі `*.pyx`, `language_level=3`).
- **XML**: `lxml >= 4.9` (libxml2) — `etree.fromstring`, `etree.iterparse`.
- **Контейнер XLSX**: стандартний `zipfile.ZipFile` (XLSX — це ZIP з XML-частинами).
- **Збірка**: `setuptools` + `Cython.Build.cythonize` (`setup.py`, `pyproject.toml`).
- **Тести**: `pytest`; частина тестів порівнює поведінку з реальним `openpyxl`.

## Архітектура модулів

```
src/cytosheet/
├── __init__.py     # Публічне API + pyximport.install() для роботи без компіляції
├── cell.pyx        # cdef class Cell — значення, координати, тип, проксі до стилів
├── worksheet.pyx   # cdef class Worksheet + Row/ColumnDimension(+Container) — грід комірок, парсинг/генерація sheetN.xml
├── workbook.pyx    # cdef class Workbook — листи, shared strings, styles.xml, збереження ZIP
├── excel.pyx       # load_workbook(...) — відкриття архіву й ініціалізація Workbook
└── styles.pyx      # Color/Side/Border/Font/PatternFill/Alignment/Protection/Style/DEFAULT_STYLE
```

Залежності між модулями: `excel → workbook → worksheet → cell → styles`.

## Модель XLSX (OOXML)

XLSX — ZIP-архів із XML-частинами. Cytosheet працює з таким набором:

```
[Content_Types].xml          # MIME-типи частин
_rels/.rels                  # кореневі звʼязки → xl/workbook.xml
xl/workbook.xml              # список листів (name, sheetId, r:id)
xl/_rels/workbook.xml.rels   # звʼязки книги → styles.xml, worksheets/sheetN.xml
xl/styles.xml                # numFmts / fonts / fills / borders / cellXfs / cellStyles
xl/sharedStrings.xml         # таблиця рядкових значень (si)
xl/worksheets/sheetN.xml     # дані листа: <cols>, <sheetData><row><c>, <mergeCells>
```

Ключові поведінкові інваріанти:

- `sheetId` починається з 1; `rId1` зарезервовано під `styles.xml`, листи йдуть з `rId2`.
- Імена листів читаються з `xl/workbook.xml`; якщо їх нема — fallback `sheet{N}`.
- При завантаженні зберігаються **оригінальні** `[Content_Types].xml`, `workbook.xml`,
  `workbook.xml.rels`, `styles.xml`, `sharedStrings.xml` та оригінальний XML кожного
  листа; при `save()` незмінені частини копіюються байт-у-байт, перезаписуються лише
  змінені листи й службові частини. Це зберігає сумісність і не «псує» шаблон.

## Стратегії парсингу листа (за розміром XML)

| Розмір `sheetN.xml` | Метод | Підхід |
|---|---|---|
| `< 50 KB` | `_parse_sheet_simple` | швидкий string-парсинг (`str.find`) |
| `50–100 KB` | `_parse_sheet_standard` | `lxml.etree.iterparse` (по тегу `c`) |
| `> 100 KB` | `_parse_sheet_chunked` | `iterparse` + батчі по 1000 комірок, `recover=True` |

Lazy-режим (`read_only=True`/`lazy=True`): листи створюються без парсингу; дані
підвантажуються при першому `iter_rows`/`iter_cols`. Генерація XML використовує
конкатенацію рядків (f-strings), а не XML-білдери — це швидше.

## Сумісність з openpyxl (станом на 0.1.0)

- **Core API** ~65%, **реальні табличні сценарії** ~90–95%, **повний API** ~45%.
- Реалізовано: Workbook/Worksheet/Cell базовий API, стилі, `number_format`,
  `merge/unmerge`, `row/column_dimensions`, формули як рядки (`<f>`), Path/BytesIO,
  доступ до рядка `ws[1]`, round-trip із openpyxl в обидва боки.
- **Не реалізовано**: діаграми, зображення, data validation, conditional formatting,
  коментарі, named ranges, фільтри, pivot, VBA, freeze panes, page setup,
  `insert/delete_rows/cols`, `copy_worksheet`, рушій обчислення формул.

Детальні живі специфікації поведінки — у `openspec/specs/<capability>/spec.md`.

## Конвенції

- **Мова**: прозовий текст у документації/специфікаціях/коментарях — українська;
  структурні ключові слова OpenSpec (`SHALL`, `WHEN`, `THEN`, `Requirement`,
  `Scenario`) — англійською (їх вимагає валідатор). Ідентифікатори в коді — англійською.
- **Cython-стиль**: публічні класи — `cdef class`; гарячі методи — `cpdef`;
  типізовані локальні змінні через `cdef`. Зовнішні сигнатури повторюють openpyxl.
- **Сумісність важливіша за чистоту**: зайві аргументи openpyxl приймаються та
  ігноруються (напр. `keep_vba`, `data_only`, `rich_text`), щоб код-міграції не падали.

## Команди

```bash
# Збірка Cython-розширень in-place
python setup.py build_ext --inplace
pip install -e .

# Тести (потрібен встановлений openpyxl для тестів сумісності)
pytest
pytest tests/test_worksheet_api.py -v
```

`pyximport.install()` у `__init__.py` дозволяє імпортувати бібліотеку й без
попередньої компіляції (`.pyx` компілюються на льоту при першому імпорті).
