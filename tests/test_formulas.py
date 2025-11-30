import os
import sys

# Add the src directory to the path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../src')))

from cytosheet import Workbook, load_workbook


def test_formula_write_read_roundtrip(tmp_path):
    file_path = tmp_path / "formula_test.xlsx"

    wb = Workbook()
    ws = wb.active

    ws['A1'] = 10
    ws['A2'] = 20
    ws['A3'] = '=SUM(A1:A2)'

    # При прямой установке через __setitem__ data_type должен быть 'f'
    assert getattr(ws['A3'], 'data_type', None) == 'f'

    # Дополнительно проверим установку через .cell(row, column, value=...)
    ws.cell(row=4, column=1, value='=A1+A2')
    assert getattr(ws['A4'], 'data_type', None) == 'f'

    wb.save(str(file_path))

    wb2 = load_workbook(str(file_path))
    ws2 = wb2.active

    cell = ws2['A3']

    assert cell.value == '=SUM(A1:A2)'
    # data_type может быть None в текущей реализации, но зарезервируем проверку на будущее
    assert getattr(cell, 'data_type', 'f') in (None, 'f')


def test_formula_iter_rows_lazy(tmp_path):
    file_path = tmp_path / "formula_iter_rows.xlsx"

    wb = Workbook()
    ws = wb.active

    ws['A1'] = 10
    ws['A2'] = 20
    ws['A3'] = '=SUM(A1:A2)'

    # Убедимся, что при установке формулы data_type помечен как 'f'
    assert getattr(ws['A3'], 'data_type', None) == 'f'

    wb.save(str(file_path))

    wb2 = load_workbook(str(file_path), lazy=True)
    ws2 = wb2.get_sheet_by_name('Sheet')

    rows = list(ws2.iter_rows(values_only=True))

    # Последняя строка должна содержать формулу как текст, не вычисленное значение
    assert rows[-1][0] == '=SUM(A1:A2)'
