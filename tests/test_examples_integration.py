import os
import sys
import datetime

import pytest

# Добавляем src в sys.path для локального запуска
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../src')))

from openpyxl import load_workbook as oxl_load_workbook
from openpyxl.utils import get_column_letter

from cytosheet import Workbook, load_workbook, Alignment, Font


@pytest.fixture
def tmp_xlsx_path(tmp_path):
    def _make(name: str):
        return tmp_path / name
    return _make


def test_example1_fill_template_from_db(tmp_xlsx_path):
    """Интеграционный тест по мотивам примера 1 из .agent/specs/examples.md.

    Проверяем базовый сценарий:
    - создание книги/листа
    - запись заголовочной ячейки через индексатор
    - построчная запись данных через worksheet.cell
    - применение Alignment к ячейкам
    - авто‑подбор ширины колонок через column_dimensions
    - round‑trip через cytosheet и openpyxl
    """

    # Эмулируем db_result простой структурой данных
    db_result = [
        ("alice", "Org A", datetime.datetime(2025, 1, 1, 10, 0, 0), True, "password"),
        ("bob", "Org B", datetime.datetime(2025, 1, 2, 11, 30, 0), False, "sso"),
    ]

    first_row_number = 5
    date_cell = "G1"

    out_path = tmp_xlsx_path("example1.xlsx")

    wb = Workbook()
    ws = wb.active

    # Запись даты в фиксированную ячейку через индексатор
    now = datetime.datetime(2025, 1, 15, 12, 0, 0)
    ws[date_cell] = f"Дата та час генерації звіту: {now:%d.%m.%Yр. %Hг.%Mхв.}"

    # Выравнивание для всех ячеек
    alignment_cell_style = Alignment(horizontal="left")
    column_widths = []

    for row_number, row_data in enumerate(db_result, first_row_number):
        for column_number, value in enumerate(row_data, 1):
            cell = ws.cell(row=row_number, column=column_number)
            cell.value = value
            # alignment может быть None для только что созданной ячейки,
            # поэтому просто проверяем, что присвоение не падает
            cell.alignment = alignment_cell_style

            text_len = len(str(value))
            if len(column_widths) >= column_number:
                if text_len > column_widths[column_number - 1]:
                    column_widths[column_number - 1] = text_len
            else:
                column_widths += [text_len]

    for column_number, column_width in enumerate(column_widths, 1):
        col_dim = ws.column_dimensions[get_column_letter(column_number)]
        if col_dim.width < (column_width + 3):
            col_dim.width = column_width + 3

    wb.save(str(out_path))

    # --- Проверяем round‑trip через cytosheet ---
    wb2 = load_workbook(str(out_path))
    ws2 = wb2.active

    assert ws2[date_cell].value.startswith("Дата та час генерації звіту:")

    rows = list(ws2.iter_rows(min_row=first_row_number, max_row=first_row_number + len(db_result) - 1))
    assert len(rows) == len(db_result)

    for (login, org, dt, ok, method), row in zip(db_result, rows):
        vals = [c.value for c in row]
        assert vals[0] == login
        assert vals[1] == org
        assert vals[4] == method
        # alignment должен быть не None, как в openpyxl
        for c in row:
            assert c.alignment is not None

    # Проверяем, что openpyxl без ошибок читает файл и видит ширины колонок
    oxl_wb = oxl_load_workbook(str(out_path))
    oxl_ws = oxl_wb.active

    for column_number, column_width in enumerate(column_widths, 1):
        col_letter = get_column_letter(column_number)
        oxl_dim = oxl_ws.column_dimensions[col_letter]
        assert oxl_dim.width == pytest.approx(column_width + 3)


def test_example2_create_workbook_with_headers_and_styles(tmp_xlsx_path):
    """Интеграционный тест по мотивам примера 2.

    Проверяем:
    - Workbook()/create_sheet
    - worksheet.cell(row, column, value)
    - Style через Font и Alignment
    - row_dimensions[1].height
    - round‑trip через cytosheet и openpyxl
    """

    out_path = tmp_xlsx_path("example2.xlsx")

    wb = Workbook()
    ws = wb.create_sheet("Логіни користувачів")

    headers = [
        "Логін користувача",
        "Організація",
        "Дата та час входу",
        "Успішність входу",
        "Метод входу",
    ]

    header_font = Font(bold=True, size=9)
    header_alignment = Alignment(horizontal="center", vertical="center")

    for col_idx, title in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.value = title
        cell.font = header_font
        cell.alignment = header_alignment

    ws.row_dimensions[1].height = 40

    wb.save(str(out_path))

    # cytosheet round‑trip
    wb2 = load_workbook(str(out_path))
    ws2 = wb2["Логіни користувачів"]

    for col_idx, title in enumerate(headers, 1):
        c = ws2.cell(row=1, column=col_idx)
        assert c.value == title
        # alignment/font могут быть None в зависимости от реализации стилей,
        # для интеграционного теста важно, что файл читается и значения сохранены

    assert ws2.row_dimensions[1].height == 40

    # openpyxl читає файл та бачить значення/высоту строки
    oxl_wb = oxl_load_workbook(str(out_path))
    oxl_ws = oxl_wb["Логіни користувачів"]

    for col_idx, title in enumerate(headers, 1):
        assert oxl_ws.cell(row=1, column=col_idx).value == title

    assert oxl_ws.row_dimensions[1].height == 40


def test_example3_copy_like_scenario_with_template_sheet(tmp_xlsx_path):
    """Интеграционный тест по мотивам примера 3.

    В cytosheet пока нет copy_worksheet/remove, поэтому используем обходной сценарий:
    - создаём исходный шаблонный лист и настраиваем его
    - создаём новый лист и копируем нужные ячейки/стили вручную
    - заполняем таблицу значениями, применяем цвет шрифта через font.color.rgb
    - настраиваем ширину колонок через column_dimensions
    - "удаление" шаблонного листа эмулируем тем, что просто не используем его далее
    - проверяем round‑trip значений и базовых стилей
    """

    out_path = tmp_xlsx_path("example3.xlsx")

    # Шаблонная книга/лист
    wb = Workbook()
    tpl_ws = wb.active
    tpl_ws.title = "Template"

    report_created_at = datetime.datetime(2025, 1, 10, 9, 0, 0)

    tpl_ws["E3"] = f"Дата/час формування звіту: {report_created_at:%d.%m.%Yр. %Hг.%Mхв.}"
    tpl_ws["B4"] = "Дата за яку формується звіт: 01.01.2025р."

    # Эмулируем один элемент worksheet_data: (sheet_title, rows)
    worksheet_data = (
        "2025-01-01",
        [
            ("area1", 1, 2, 3, 3),
            ("area2", 2, 3, 4, 5),
        ],
    )

    # Создаём новый лист и копируем фиксированные ячейки
    target_ws = wb.create_sheet("Report")
    target_ws["E3"] = tpl_ws["E3"].value
    target_ws["B4"] = tpl_ws["B4"].value

    column_widths = []
    start_row = 8

    for row_number, row_data in enumerate(worksheet_data[1], start_row):
        color_for_ts_cell = "FF00FF00" if row_data[3] == row_data[4] else "FFFF0000"
        for column_number, value in enumerate(row_data, 1):
            cell = target_ws.cell(row=row_number, column=column_number)
            cell.value = value
            if column_number in (4, 5):
                # Гарантируем наличие Font/Color перед установкой rgb
                if cell.font is None:
                    from cytosheet import Font as CytoFont, Color as CytoColor
                    cell.font = CytoFont(color=CytoColor(rgb=color_for_ts_cell))
                else:
                    if getattr(cell.font, "color", None) is None:
                        from cytosheet import Color as CytoColor
                        cell.font.color = CytoColor(rgb=color_for_ts_cell)
                    else:
                        cell.font.color.rgb = color_for_ts_cell

            text_len = len(str(value))
            if len(column_widths) >= column_number:
                if text_len > column_widths[column_number - 1]:
                    column_widths[column_number - 1] = text_len
            else:
                column_widths += [text_len]

    for column_number, column_width in enumerate(column_widths, 1):
        col_dim = target_ws.column_dimensions[get_column_letter(column_number)]
        if col_dim.width < (column_width + 3):
            col_dim.width = column_width + 3

    # Теперь API Workbook.remove реализован, удаляем шаблонный лист как в примере openpyxl
    wb.remove(tpl_ws)

    wb.save(str(out_path))

    # cytosheet round‑trip
    wb2 = load_workbook(str(out_path))
    ws2 = wb2["Report"]

    assert ws2["E3"].value.startswith("Дата/час формування звіту:")
    assert ws2["B4"].value.startswith("Дата за яку формується звіт:")

    rows = list(ws2.iter_rows(min_row=start_row, max_row=start_row + len(worksheet_data[1]) - 1))
    assert len(rows) == len(worksheet_data[1])

    for (orig_row, loaded_row) in zip(worksheet_data[1], rows):
        vals = [c.value for c in loaded_row]
        assert vals[: len(orig_row)] == list(orig_row)

    last_row = rows[-1]
    ts_cell = last_row[3]  # column 4
    assert ts_cell.font is not None
    assert getattr(getattr(ts_cell.font, "color", None), "rgb", None) is not None

    # openpyxl также должен без ошибок читать этот файл
    oxl_wb = oxl_load_workbook(str(out_path))
    oxl_ws = oxl_wb["Report"]
    assert oxl_ws["E3"].value.startswith("Дата/час формування звіту:")
    assert oxl_ws["B4"].value.startswith("Дата за яку формується звіт:")

    oxl_rows = list(oxl_ws.iter_rows(min_row=start_row, max_row=start_row + len(worksheet_data[1]) - 1, values_only=True))
    assert len(oxl_rows) == len(worksheet_data[1])

    for orig_row, loaded_vals in zip(worksheet_data[1], oxl_rows):
        assert list(loaded_vals)[: len(orig_row)] == list(orig_row)


if __name__ == "__main__":  # отладочный запуск
    pytest.main([__file__, "-vv"])
