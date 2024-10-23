import os
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
    test_file = os.path.join(os.path.dirname(__file__), 'test.xlsx')

    wb = Workbook()
    wb.load_workbook(test_file)

    ws = wb.get_sheet_by_name('Sheet1')
    assert ws['A3'].value == '1234'

# if __name__ == '__main__':
#     test_read_cells()