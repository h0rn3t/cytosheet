# Cytosheet v1.0 Checklist

Этот файл фиксирует прогресс по ключевым задачам, необходимым для v1.0 и совместимости с openpyxl.

## 1. Обязательные для v1.0

### 1.1 Worksheet API (диапазоны, размеры, итерация)

- [x] WS-001 `Worksheet.max_row`
- [x] WS-002 `Worksheet.max_column`
- [x] WS-003 `Worksheet.dimensions`
- [x] WS-004 `Worksheet.append(iterable)`
- [x] WS-005 `Worksheet.iter_cols(...)`
- [x] WS-006 `Worksheet.rows` (proxy к `iter_rows`)
- [x] WS-007 `Worksheet.columns` (proxy к `iter_cols`)
- [x] WS-008 `Worksheet.values` (proxy к `iter_rows(values_only=True)`)
- [x] WS-009 `Worksheet['A1:C3']` (диапазоны ячеек)

### 1.2 Формулы

- [x] FM-001 `cell.data_type = 'f'` при установке формулы через API
- [x] FM-002 Корректная запись `<f>` при сохранении XLSX
- [ ] FM-003 Опциональная запись кэшированного `<v>` значения для формул (перенесено за рамки v1.0)

### 1.3 NumberFormat / стили

- [x] NF-001 Свойство `cell.number_format` (proxy к `Style.numberFormat`)
- [x] NF-002 Запись number formats в `styles.xml`
- [x] NF-003 Чтение number formats из `styles.xml`

### 1.4 Row / Column dimensions

- [x] DM-001 `worksheet.row_dimensions[row].height`
- [x] DM-002 `worksheet.row_dimensions[row].hidden`
- [x] DM-003 `worksheet.column_dimensions[col].width`
- [x] DM-004 `worksheet.column_dimensions[col].hidden`
- [x] DM-005 Парсинг `<row>` тегов из `sheet.xml`
- [x] DM-006 Парсинг `<col>` тегов из `sheet.xml`
- [x] DM-007 Генерация `<row>` / `<col>` при сохранении

### 1.5 API load_workbook (совместимость с openpyxl)

- [ ] LW-001 Сигнатура `load_workbook(filename, read_only=False, keep_vba=False, data_only=False, keep_links=True, rich_text=False)`
- [ ] LW-002 Аргумент `read_only` маппится на внутренний `lazy`
- [ ] LW-003 Остальные аргументы принимаются и игнорируются без ошибок

### 1.6 Cell proxies к стилям

- [ ] CP-001 `cell.font` (proxy к `cell.style.font`)
- [ ] CP-002 `cell.alignment` (proxy к `cell.style.alignment`)
- [ ] CP-003 `cell.border` (proxy к `cell.style.border`)
- [ ] CP-004 `cell.fill` (proxy к `cell.style.fill`)
- [ ] CP-005 `cell.protection` (proxy к `cell.style.protection`)

## 2. Очень желательно для 1.0 (можно перенести в 1.1)

- [ ] FZ-001 `Worksheet.freeze_panes`
- [ ] MD-001 `worksheet.insert_rows(idx, amount=1)`
- [ ] MD-002 `worksheet.delete_rows(idx, amount=1)`
- [ ] MD-003 `worksheet.insert_cols(idx, amount=1)`
- [ ] MD-004 `worksheet.delete_cols(idx, amount=1)`
- [ ] SV-001 `Workbook.save_virtual_workbook()`
- [ ] CS-001 `Workbook.create_sheet(title=None, index=None)` с поддержкой `index`

---

## История изменений

- 2025-11-30: Создан базовый чеклист (GitHub Copilot).
- 2025-11-30: Реализованы WS-001, WS-002, WS-004, WS-006, WS-007, WS-008, WS-009 (Worksheet bounds, append, rows/cols/values, диапазоны).
- 2025-11-30: Реализованы WS-003 (dimensions) и WS-005 (iter_cols) + дополнительные тесты на итерацию и диапазоны.
- 2025-11-30: Реализованы DM-001..DM-007 (row_dimensions/column_dimensions API, парсинг/генерация <row>/<col>, совместимость с openpyxl).
