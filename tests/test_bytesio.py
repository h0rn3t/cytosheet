#!/usr/bin/env python3
"""Тест сохранения в BytesIO"""

import pytest
from io import BytesIO
from cytosheet import Workbook, load_workbook
from cytosheet.styles import Alignment

def test_save_to_bytesio():
    """Сохранение в BytesIO"""
    # Создаем книгу
    wb = Workbook()
    ws = wb.active
    ws['A1'] = 'Header 1'
    ws['B1'] = 'Header 2'
    ws['C1'] = 'Header 3'
    ws['A2'] = 'Data 1'
    ws['B2'] = 'Data 2'
    ws['C2'] = 'Data 3'

    # Применяем стили
    ws['B1'].alignment = Alignment(horizontal="left")

    # Сохраняем в BytesIO (как в вашем коде)
    file_stream = BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)

    assert len(file_stream.getvalue()) > 0

def test_load_from_bytesio():
    """Загрузка файла из BytesIO"""
    # Создаем книгу
    wb = Workbook()
    ws = wb.active
    ws['A1'] = 'Header 1'
    ws['B2'] = 'Data 2'

    # Сохраняем в BytesIO
    file_stream = BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)

    # Проверяем что можем загрузить обратно
    import tempfile
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        f.write(file_stream.getvalue())
        tmp_path = f.name

    wb2 = load_workbook(tmp_path)
    ws2 = wb2.active

    assert ws2['A1'].value == 'Header 1'
    assert ws2['B2'].value == 'Data 2'

