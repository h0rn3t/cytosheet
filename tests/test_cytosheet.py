import os
import timeit

import sys
import pytest

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "src")))

try:
    from openpyxl import load_workbook as openpyxl_load_workbook
    HAS_OPENPYXL = True
except Exception:  # noqa: PIE786
    HAS_OPENPYXL = False

from cytosheet import Workbook, load_workbook


def test_parse_xlsx():
    test_file = os.path.join(os.path.dirname(__file__), "test.xlsx")
    wb = load_workbook(test_file)
    assert "Sheet1" in wb.sheetnames


def test_read_cells():
    test_file = os.path.join(os.path.dirname(__file__), "file_example_XLSX_5000.xlsx")

    start = timeit.default_timer()
    wb = load_workbook(test_file)
    ws = wb["Sheet1"]
    assert ws["A2"].value == "1"
    cytosheet_time = timeit.default_timer() - start
    print(f"cytosheet Time: {cytosheet_time:.4f} seconds")

    if HAS_OPENPYXL:
        start = timeit.default_timer()
        wb = openpyxl_load_workbook(test_file)
        ws = wb["Sheet1"]
        assert str(ws["A2"].value) == "1"
        openpyxl_time = timeit.default_timer() - start
        print(f"openpyxl Time: {openpyxl_time:.4f} seconds")


def test_write_cells(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws["A1"].value = "Hello"
    ws["B2"].value = "World"

    assert ws["A1"].value == "Hello"
    assert ws["B2"].value == "World"

    file_path = tmp_path / "cytosheet_create_file_test.xlsx"
    wb.save(str(file_path))

    wb2 = load_workbook(str(file_path))
    ws2 = wb2.active
    assert ws2["A1"].value == "Hello"
    assert ws2["B2"].value == "World"
