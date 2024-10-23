import os
import timeit
from openpyxl import load_workbook
import pytest
from src.cytosheet import Workbook


import sys
sys.path.append("src")

def test_parse_xlsx():
    """Тестирование парсинга существующего XLSX файла."""
    test_file = os.path.join(os.path.dirname(__file__), 'test.xlsx')

    wb = Workbook()
    wb.load_workbook(test_file)


    assert 'sheet1' == wb.get_sheet_by_name('Sheet1').title


def test_read_cells():
    """Тестирование чтения значений ячеек."""
    test_file = os.path.join(os.path.dirname(__file__), 'file_example_XLSX_5000.xlsx')

    # Тестирование cytosheet
    start = timeit.default_timer()
    wb = Workbook()
    wb.load_workbook(test_file)
    ws = wb.get_sheet_by_name('Sheet1')
    assert ws['A2'].value == '1'  # Приведение типов, если нужно
    stop = timeit.default_timer()
    cytosheet_time = stop - start
    print(f'\ncytosheet Time: {cytosheet_time:.4f} seconds')

    # Тестирование openpyxl
    start = timeit.default_timer()
    wb = load_workbook(test_file)
    ws = wb.get_sheet_by_name('Sheet1')
    assert str(ws['A2'].value) == '1'  # Приведение типов
    stop = timeit.default_timer()
    openpyxl_time = stop - start
    print(f'\nopenpyxl Time: {openpyxl_time:.4f} seconds')
