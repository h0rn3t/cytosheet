# Cytosheet - Quick Start для LLM агентів

## Місія проекту

Створити високопродуктивну Cython-бібліотеку для роботи з XLSX файлами, яка є **100% сумісною з openpyxl API** і може використовуватись як drop-in replacement без змін у коді користувачів.

## Ключові цілі

1. ✅ **API сумісність**: Повна відповідність openpyxl API
2. ⚡ **Продуктивність**: 2-10x швидше ніж openpyxl
3. 🔧 **Drop-in replacement**: `from cytosheet import Workbook` замість `from openpyxl import Workbook`
4. 🧪 **Тестування**: Всі тести openpyxl мають проходити

## Поточний стан (v0.1.0)

### 📊 Совместимость API с openpyxl

**Дата оценки**: 2025-12-05

- **Базовый API (Core)**: ~65% совместимость
- **Реальные табличные сценарии**: ~90-95% совместимость
- **Полный API openpyxl**: ~45% совместимость

Детальный анализ: см. `.agent/api_compatibility_analysis.md`

### ✅ Реалізовано

#### Core API (полная поддержка)
- Базова структура (Workbook, Worksheet, Cell) ✅
- Читання XLSX файлів з підтримкою lazy loading ✅
- Запис XLSX файлів (включая save_virtual_workbook) ✅
- Базові стилі (Font, Border, PatternFill, Alignment, Protection, Color, Side) ✅
- Об'єднання ячейок (merge_cells/unmerge_cells) ✅
- Shared strings підтримка ✅
- Оптимізований парсинг (chunked для великих файлів) ✅
- **Формули** - запис/читання формул (data_type='f', XML <f> теги) ✅
- **NumberFormat** - форматування чисел через cell.number_format ✅
- **Row/Column dimensions** - висота/ширина/приховування рядків і колонок ✅
- **Ітератори** - iter_rows, iter_cols, rows, columns, values ✅
- **Діапазони** - доступ до діапазонів ws['A1:C3'] ✅
- **append()** метод для додавання рядків ✅
- **max_row, max_column** - визначення меж даних ✅

#### Совместимость по категориям
- Workbook API: 52% (11/21 полностью + 3/21 частично)
- Worksheet API: 48% (23/48 полностью + 5/48 частично)
- Cell API: 82% (14/17 полностью + 2/17 частично)
- Styles API: 80% (8/10 полностью + 1/10 частично)

### ❌ Відсутнє (критичне для повної сумісності)

#### High Priority (Milestone 2-3)
- Freeze panes - закріплення областей ❌
- Insert/Delete rows/cols - вставка/удаление строк/колонок ❌
- Auto filter - автофільтри ❌
- Data validation - перевірка даних ❌
- Conditional formatting - умовне форматування ❌
- Copy worksheet - копіювання аркушів ❌

#### Advanced Features (Milestone 4+)
- Діаграми (Charts) - всі типи ❌
- Зображення (Images) ❌
- Коментарі (Comments) ❌
- Захист аркушів (Protection) - частково 🟡
- Іменовані діапазони (Named ranges) ❌
- Page setup - налаштування друку ❌
- Фільтри та сортування ❌
- Pivot tables ❌
- Gradient fills ❌

## Швидкий старт для розробки

### 1. Встановлення залежностей

```bash
pip install Cython>=3.0.11 lxml>=4.9 openpyxl pytest
```

### 2. Компіляція Cython модулів

```bash
cd src
python setup.py build_ext --inplace
```

### 3. Запуск тестів

```bash
pytest tests/tests.py -v
pytest tests/benchmark_tests.py -v
```

## Структура проекту

```
cytosheet/
├── src/cytosheet/
│   ├── __init__.py          # Публічне API
│   ├── cell.pyx             # Cell клас
│   ├── workbook.pyx         # Workbook клас
│   ├── worksheet.pyx        # Worksheet клас
│   ├── excel.pyx            # load_workbook функція
│   └── styles.pyx           # Стилі (Font, Border, etc.)
├── tests/
│   ├── tests.py             # Базові тести
│   └── benchmark_tests.py   # Бенчмарки
└── .agent/                  # Документація для агентів
```

## Пріоритетні завдання

### Фаза 1: Розширення базового API (поточна)

1. **Формули** - додати підтримку `cell.value = "=SUM(A1:A10)"`
2. **NumberFormat** - форматування чисел як у openpyxl
3. **Row/Column dimensions** - висота/ширина рядків і стовпців
4. **Freeze panes** - закріплення областей
5. **Page setup** - налаштування друку

### Фаза 2: Розширені функції

1. Charts (діаграми)
2. Images (зображення)
3. Data validation
4. Conditional formatting
5. Comments

### Фаза 3: Повна сумісність

1. Named ranges
2. Protection
3. Filters
4. Pivot tables
5. VBA macros (readonly)

## Правила розробки

### API сумісність

```python
# Цей код має працювати однаково в openpyxl і cytosheet
from cytosheet import Workbook  # або from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws['A1'] = 'Hello'
ws['A1'].font = Font(bold=True)
wb.save('test.xlsx')
```

### Тестування сумісності

Кожна нова функція має тест, який:
1. Виконується для openpyxl
2. Виконується для cytosheet
3. Порівнює результати (ідентичні)

```python
def test_feature_compatibility():
    # Test with openpyxl
    from openpyxl import Workbook as OpenpyxlWorkbook
    wb1 = OpenpyxlWorkbook()
    # ... test logic ...
    
    # Test with cytosheet
    from cytosheet import Workbook
    wb2 = Workbook()
    # ... same test logic ...
    
    # Assert results are identical
    assert result1 == result2
```

### Продуктивність

- Використовуйте Cython типізацію (`cdef`, `cpdef`)
- Мінімізуйте Python об'єкти у критичних місцях
- Оптимізуйте парсинг XML (SAX для великих файлів)
- Батчева обробка даних

## Корисні команди

```bash
# Компіляція з детальним виводом
python setup.py build_ext --inplace --verbose

# Запуск конкретного тесту
pytest tests/tests.py::test_write_cells -v

# Бенчмарки
python tests/benchmark_tests.py

# Перевірка сумісності API
python -c "from cytosheet import Workbook; print(dir(Workbook))"
```

## Наступні кроки

1. Прочитайте `specs/requirements.md` для деталей
2. Ознайомтесь з `specs/api-reference.md` для повного API
3. Виберіть завдання з `specs/tasks.md`
4. Вивчіть `specs/design.md` для архітектурних рішень

## Джерела

- [openpyxl документація](https://openpyxl.readthedocs.io/)
- [OOXML стандарт](http://www.ecma-international.org/publications/standards/Ecma-376.htm)
- [Cython документація](https://cython.readthedocs.io/)