#!/usr/bin/env python3
"""Тест совместимости с openpyxl API"""

import pytest
from pathlib import Path
from cytosheet import Workbook, load_workbook
from cytosheet.styles import Alignment
import tempfile

def test_path_object_in_load_workbook():
    """Path объект в load_workbook"""
    wb = Workbook()
    ws = wb.active
    ws['A1'] = 'Header 1'
    ws['B1'] = 'Header 2'
    ws['C1'] = 'Header 3'
    ws['A2'] = 'Data 1'
    ws['B2'] = 'Data 2'
    ws['C2'] = 'Data 3'

    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        test_file = Path(tmp.name)
    
    wb.save(str(test_file))

    # Загружаем через Path объект (как в вашем коде)
    wb2 = load_workbook(filename=test_file)
    assert wb2.active['A1'].value == 'Header 1'

def test_access_row_by_number():
    """Доступ к строке по номеру"""
    wb = Workbook()
    ws = wb.active
    ws['A1'] = 'Header 1'
    ws['B1'] = 'Header 2'
    ws['C1'] = 'Header 3'

    # Получаем заголовки через ws[1] (первая строка)
    headers = [cell.value for cell in ws[1]]
    assert headers == ['Header 1', 'Header 2', 'Header 3']

def test_full_scenario():
    """Полный сценарий как в коде пользователя"""
    wb = Workbook()
    ws = wb.active
    ws['A1'] = 'Header 1'
    ws['B1'] = 'Header 2'
    ws['C1'] = 'Header 3'
    
    index_headers = 1
    start_row = 2

    # получаем хидеры из эксельки (как в вашем коде)
    headers = [cell.value for cell in ws[index_headers]]
    assert headers == ['Header 1', 'Header 2', 'Header 3']

    # Симулируем данные из JSON
    data = [
        {'Header 1': 'Row2-A', 'Header 2': 'Row2-B', 'Header 3': 'Row2-C'},
        {'Header 1': 'Row3-A', 'Header 2': 'Row3-B', 'Header 3': 'Row3-C'},
    ]

    for row_index, row in enumerate(data, start=start_row):
        for col_index, header in enumerate(headers):
            value = row.get(str(header), "")
            ws.cell(row=row_index, column=col_index + 1, value=value)

    # Проверяем что данные записались
    assert ws['A2'].value == 'Row2-A'
    assert ws['B3'].value == 'Row3-B'

def test_alignment():
    """Alignment как в коде пользователя"""
    wb = Workbook()
    ws = wb.active
    ws['B1'] = 'Test ID'
    ws['B1'].alignment = Alignment(horizontal="left")
    
    assert ws['B1'].alignment.horizontal == "left"

