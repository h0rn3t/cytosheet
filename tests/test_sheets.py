#!/usr/bin/env python3
"""Тесты для работы с листами (sheets)"""

import pytest
from cytosheet import Workbook, load_workbook
from pathlib import Path
import tempfile

def test_create_multiple_sheets():
    """Создание нескольких листов"""
    wb = Workbook()
    
    # Первый лист (active по умолчанию)
    assert wb.active.title == "Sheet"
    
    # Создаем дополнительные листы
    ws2 = wb.create_sheet("Лист2")
    ws3 = wb.create_sheet("Лист3")
    
    assert len(wb.sheetnames) == 3
    assert wb.sheetnames == ["Sheet", "Лист2", "Лист3"]

def test_access_sheets_by_name():
    """Доступ к листам по имени"""
    wb = Workbook()
    wb.create_sheet("Данные")
    wb.create_sheet("Отчет")
    
    # Доступ через __getitem__
    ws1 = wb["Данные"]
    ws2 = wb["Отчет"]
    
    assert ws1.title == "Данные"
    assert ws2.title == "Отчет"

def test_access_sheets_by_index():
    """Доступ к листам по индексу"""
    wb = Workbook()
    wb.create_sheet("Второй")
    wb.create_sheet("Третий")
    
    # Доступ через sheetnames и индекс
    ws1 = wb[wb.sheetnames[0]]
    ws2 = wb[wb.sheetnames[1]]
    ws3 = wb[wb.sheetnames[2]]
    
    assert ws1.title == "Sheet"
    assert ws2.title == "Второй"
    assert ws3.title == "Третий"

def test_remove_sheet():
    """Удаление листа"""
    wb = Workbook()
    ws2 = wb.create_sheet("Удалить")
    ws3 = wb.create_sheet("Оставить")
    
    assert len(wb.sheetnames) == 3
    
    # Удаляем лист
    wb.remove(ws2)
    
    assert len(wb.sheetnames) == 2
    assert "Удалить" not in wb.sheetnames
    assert "Оставить" in wb.sheetnames

def test_write_to_different_sheets():
    """Запись данных в разные листы"""
    wb = Workbook()
    ws1 = wb.active
    ws2 = wb.create_sheet("Продажи")
    ws3 = wb.create_sheet("Расходы")
    
    # Записываем данные в разные листы
    ws1['A1'] = "Главная"
    ws2['A1'] = "Продажи 2024"
    ws3['A1'] = "Расходы 2024"
    
    assert ws1['A1'].value == "Главная"
    assert ws2['A1'].value == "Продажи 2024"
    assert ws3['A1'].value == "Расходы 2024"

def test_save_and_load_multiple_sheets():
    """Сохранение и загрузка файла с несколькими листами"""
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        tmp_path = tmp.name
    
    # Создаем и сохраняем
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Первый"
    ws1['A1'] = "Данные 1"
    
    ws2 = wb.create_sheet("Второй")
    ws2['A1'] = "Данные 2"
    
    ws3 = wb.create_sheet("Третий")
    ws3['A1'] = "Данные 3"
    
    wb.save(tmp_path)
    
    # Загружаем обратно
    wb2 = load_workbook(tmp_path)
    
    assert len(wb2.sheetnames) == 3
    assert wb2.sheetnames == ["Первый", "Второй", "Третий"]
    
    assert wb2["Первый"]['A1'].value == "Данные 1"
    assert wb2["Второй"]['A1'].value == "Данные 2"
    assert wb2["Третий"]['A1'].value == "Данные 3"

def test_active_sheet():
    """Работа с активным листом"""
    wb = Workbook()
    
    # По умолчанию активен первый лист
    assert wb.active.title == "Sheet"
    
    ws2 = wb.create_sheet("Новый")
    
    # active все еще указывает на первый лист
    assert wb.active.title == "Sheet"
    
    # Можем менять данные через active
    wb.active['A1'] = "Активный лист"
    assert wb["Sheet"]['A1'].value == "Активный лист"

def test_sheet_with_cyrillic_names():
    """Листы с кириллическими именами"""
    wb = Workbook()
    ws1 = wb.create_sheet("Продажи")
    ws2 = wb.create_sheet("Расходы")
    ws3 = wb.create_sheet("Прибыль")
    
    ws1['A1'] = "Январь"
    ws2['A1'] = "Февраль"
    ws3['A1'] = "Март"
    
    assert wb["Продажи"]['A1'].value == "Январь"
    assert wb["Расходы"]['A1'].value == "Февраль"
    assert wb["Прибыль"]['A1'].value == "Март"

def test_sheet_iteration():
    """Итерация по листам"""
    wb = Workbook()
    wb.create_sheet("Лист2")
    wb.create_sheet("Лист3")
    
    titles = []
    for sheet in wb.sheets.values():
        titles.append(sheet.title)
    
    assert titles == ["Sheet", "Лист2", "Лист3"]

