import os
import sys

# Добавляем src в sys.path для локального запуска
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../src')))

from openpyxl import Workbook as OpenpyxlWorkbook, load_workbook as oxl_load_workbook
from openpyxl.utils import get_column_letter

from cytosheet import Workbook, load_workbook


def test_column_dimensions_api_width_compatible_with_openpyxl(tmp_path):
    """Проверяем, что ws.column_dimensions[get_column_letter(idx)].width доступен и изменяемый.

    Сценарий максимально близок к коду пользователя: обходим ширины колонок и увеличиваем их.
    """
    wb = Workbook()
    ws = wb.active

    column_widths = [5, 10, 15]
    for column_number, column_width in enumerate(column_widths, 1):
        col_letter = get_column_letter(column_number)
        col_dim = ws.column_dimensions[col_letter]
        # в cytosheet width по умолчанию 0.0, в openpyxl — None; нам важно, что мы можем увеличить
        if col_dim.width < (column_width + 3):
            col_dim.width = column_width + 3

    # сохраняем и перечитываем нашей же библиотекой
    file_path = tmp_path / "col_dims_roundtrip.xlsx"
    wb.save(str(file_path))

    wb2 = load_workbook(str(file_path))
    ws2 = wb2.active

    for column_number, column_width in enumerate(column_widths, 1):
        col_letter = get_column_letter(column_number)
        col_dim2 = ws2.column_dimensions[col_letter]
        # после загрузки должны видеть ту же ширину
        assert col_dim2.width == column_width + 3


def test_column_dimensions_written_readable_by_openpyxl(tmp_path):
    """Пишем ширины колонок в cytosheet, читаем их через openpyxl."""
    wb = Workbook()
    ws = wb.active

    ws['A1'] = 1
    ws['B1'] = 2

    ws.column_dimensions['A'].width = 20
    ws.column_dimensions['B'].width = 40

    file_path = tmp_path / "col_dims_openpyxl.xlsx"
    wb.save(str(file_path))

    oxl_wb = oxl_load_workbook(str(file_path))
    oxl_ws = oxl_wb.active

    assert oxl_ws.column_dimensions['A'].width == 20
    assert oxl_ws.column_dimensions['B'].width == 40


def test_row_dimensions_height_and_hidden_roundtrip(tmp_path):
    """Пишем высоту/hidden для строк, проверяем roundtrip через cytosheet."""
    wb = Workbook()
    ws = wb.active

    ws['A1'] = 'x'
    ws['A2'] = 'y'
    ws['A3'] = 'z'

    ws.row_dimensions[1].height = 25.0
    ws.row_dimensions[2].hidden = True

    file_path = tmp_path / "row_dims_roundtrip.xlsx"
    wb.save(str(file_path))

    wb2 = load_workbook(str(file_path))
    ws2 = wb2.active

    rd1 = ws2.row_dimensions[1]
    rd2 = ws2.row_dimensions[2]

    assert rd1.height == 25.0
    assert rd2.hidden is True


def test_row_dimensions_written_readable_by_openpyxl(tmp_path):
    """Проверяем, что openpyxl видит высоты/hidden строк, записанные cytosheet."""
    wb = Workbook()
    ws = wb.active

    ws['A1'] = 'foo'
    ws['A2'] = 'bar'

    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].hidden = True

    file_path = tmp_path / "row_dims_openpyxl.xlsx"
    wb.save(str(file_path))

    oxl_wb = oxl_load_workbook(str(file_path))
    oxl_ws = oxl_wb.active

    assert oxl_ws.row_dimensions[1].height == 30
    assert oxl_ws.row_dimensions[2].hidden is True


def test_load_existing_openpyxl_file_with_column_dimensions(tmp_path):
    """Создаём файл через openpyxl с ширинами колонок и проверяем, что cytosheet их видит."""
    oxl_wb = OpenpyxlWorkbook()
    oxl_ws = oxl_wb.active

    oxl_ws['A1'] = 'a'
    oxl_ws['B1'] = 'b'

    oxl_ws.column_dimensions['A'].width = 15
    oxl_ws.column_dimensions['B'].width = 25

    file_path = tmp_path / "from_openpyxl_cols.xlsx"
    oxl_wb.save(str(file_path))

    wb = load_workbook(str(file_path))
    ws = wb.active

    cd_a = ws.column_dimensions['A']
    cd_b = ws.column_dimensions['B']

    assert cd_a.width == 15
    assert cd_b.width == 25


def test_load_existing_openpyxl_file_with_row_dimensions(tmp_path):
    """Создаём файл через openpyxl с высотами/hidden строк и проверяем, что cytosheet их видит."""
    oxl_wb = OpenpyxlWorkbook()
    oxl_ws = oxl_wb.active

    oxl_ws['A1'] = '1'
    oxl_ws['A2'] = '2'

    oxl_ws.row_dimensions[1].height = 22
    oxl_ws.row_dimensions[2].hidden = True

    file_path = tmp_path / "from_openpyxl_rows.xlsx"
    oxl_wb.save(str(file_path))

    wb = load_workbook(str(file_path))
    ws = wb.active

    rd1 = ws.row_dimensions[1]
    rd2 = ws.row_dimensions[2]

    assert rd1.height == 22
    assert rd2.hidden is True
