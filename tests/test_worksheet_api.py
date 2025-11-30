import os
import sys

import pytest

# Add the src directory to the path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../src')))

from cytosheet import Workbook, load_workbook


def test_cell_row_column_and_max_bounds_basic(tmp_path):
    """Перевіряє ws.cell(), max_row, max_column та append() на простому кейсі."""
    wb = Workbook()
    ws = wb.active

    # Пишемо кілька значень по координатах
    ws.cell(row=1, column=1, value="A1")
    ws.cell(row=1, column=3, value="C1")
    ws.cell(row=5, column=2, value="B5")

    # Перевіряємо доступ через A1/B5/C1
    assert ws['A1'].value == "A1"
    assert ws['C1'].value == "C1"
    assert ws['B5'].value == "B5"

    # Межі
    assert ws.max_row == 5
    assert ws.max_column == 3

    # append додає рядок після max_row
    ws.append([1, 2, 3])
    assert ws.max_row == 6
    assert ws['A6'].value == 1
    assert ws['B6'].value == 2
    assert ws['C6'].value == 3


def test_cell_and_append_roundtrip_with_save(tmp_path):
    """Перевірка, що cell()/append() коректно працюють після save/load_workbook."""
    file_path = tmp_path / "cell_append_roundtrip.xlsx"

    wb = Workbook()
    ws = wb.active

    # Заповнюємо перший рядок через cell()
    ws.cell(row=1, column=1, value="ID")
    ws.cell(row=1, column=2, value="Name")

    # Другий і третій рядок через append()
    ws.append([1, "Alice"])
    ws.append([2, "Bob"])

    assert ws.max_row == 3
    assert ws.max_column == 2

    wb.save(str(file_path))

    # Читаємо назад
    wb2 = load_workbook(str(file_path))
    ws2 = wb2.active

    assert ws2['A1'].value == "ID"
    assert ws2['B1'].value == "Name"
    assert ws2['A2'].value == 1
    assert ws2['B2'].value == "Alice"
    assert ws2['A3'].value == 2
    assert ws2['B3'].value == "Bob"


def test_max_row_max_column_with_indexing_and_cell(tmp_path):
    """Змішане використання ws['A1'] і ws.cell() не ламає max_row/max_column."""
    wb = Workbook()
    ws = wb.active

    # Через індексацію
    ws['A1'] = 10
    ws['D4'] = 40

    # Через cell()
    ws.cell(row=2, column=3, value=30)  # C2

    assert ws.max_row == 4
    # D -> 4
    assert ws.max_column == 4


def test_append_non_iterable_puts_value_in_first_column():
    """Якщо в append() передати неітерабельний обʼєкт — він пишеться у перший стовпець."""
    wb = Workbook()
    ws = wb.active

    ws.append("single")

    assert ws.max_row == 1
    assert ws.max_column == 1
    assert ws['A1'].value == "single"


def test_cell_formulas_and_max_bounds(tmp_path):
    """Перевірка, що формули через cell() зберігаються й не ламають max_row/max_column."""
    file_path = tmp_path / "cell_formulas.xlsx"

    wb = Workbook()
    ws = wb.active

    ws.cell(row=1, column=1, value=10)
    ws.cell(row=2, column=1, value=20)
    ws.cell(row=3, column=1, value="=SUM(A1:A2)")

    assert ws.max_row == 3
    assert ws.max_column == 1

    wb.save(str(file_path))

    wb2 = load_workbook(str(file_path))
    ws2 = wb2.active

    assert ws2['A1'].value == 10
    assert ws2['A2'].value == 20
    assert ws2['A3'].value == "=SUM(A1:A2)"


def test_dimensions_basic():
    """Проверяем, что dimensions отражает используемый прямоугольник, как в openpyxl."""
    wb = Workbook()
    ws = wb.active

    ws['B2'] = 1
    ws['D5'] = 2

    assert ws.max_row == 5
    assert ws.max_column == 4
    assert ws.dimensions == 'B2:D5'


def test_iter_rows_and_values_basic():
    """Базова перевірка iter_rows і .values."""
    wb = Workbook()
    ws = wb.active

    data = [
        ["ID", "Name"],
        [1, "Alice"],
        [2, "Bob"],
    ]

    for r_idx, row in enumerate(data, 1):
        for c_idx, value in enumerate(row, 1):
            ws.cell(row=r_idx, column=c_idx, value=value)

    # iter_rows с values_only=False
    rows = list(ws.iter_rows(min_row=1, max_row=3, min_col=1, max_col=2, values_only=False))
    assert len(rows) == 3
    assert [cell.value for cell in rows[0]] == ["ID", "Name"]

    # .values як proxy к iter_rows(values_only=True)
    values_rows = list(ws.values)
    assert values_rows[0] == ("ID", "Name")
    assert values_rows[1] == (1, "Alice")
    assert values_rows[2] == (2, "Bob")


def test_iter_cols_and_columns_basic():
    """Базова перевірка iter_cols і свойства .columns."""
    wb = Workbook()
    ws = wb.active

    data = [
        ["ID", "Name"],
        [1, "Alice"],
        [2, "Bob"],
    ]

    for r_idx, row in enumerate(data, 1):
        for c_idx, value in enumerate(row, 1):
            ws.cell(row=r_idx, column=c_idx, value=value)

    cols = list(ws.iter_cols(min_col=1, max_col=2, min_row=1, max_row=3, values_only=True))
    assert cols[0] == ("ID", 1, 2)
    assert cols[1] == ("Name", "Alice", "Bob")

    # .columns як proxy к iter_cols(values_only=False)
    cols_cells = list(ws.columns)
    assert [cell.value for cell in cols_cells[0]] == ["ID", 1, 2]
    assert [cell.value for cell in cols_cells[1]] == ["Name", "Alice", "Bob"]


def test_range_indexing_A1_C3():
    """Проверяем, что ws['A1:C3'] возвращает tuple(tuple(Cell))."""
    wb = Workbook()
    ws = wb.active

    for r in range(1, 4):
        for c in range(1, 4):
            ws.cell(row=r, column=c, value=f"R{r}C{c}")

    rng = ws['A1:C3']
    assert len(rng) == 3
    assert len(rng[0]) == 3
    assert rng[0][0].value == 'R1C1'
    assert rng[2][2].value == 'R3C3'
