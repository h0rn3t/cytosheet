"""
Упрощенные бенчмарки для локального тестирования cytosheet vs openpyxl
"""

import os
import time
import random
import gc
import sys
import traceback
import pytest
from typing import List, Dict, Any, Tuple


sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

try:
    from openpyxl import load_workbook as openpyxl_load_workbook, Workbook as OpenpyxlWorkbook
    from cytosheet import (
        Workbook, load_workbook, 
        Color, Side, Border, Font, PatternFill, 
        Alignment, Protection, Style
    )
except ImportError as e:
    print(f"Ошибка импорта: {e}")
    print("Убедитесь, что все зависимости установлены и cytosheet скомпилирован")
    sys.exit(1)


class SimpleBenchmark:
    def __init__(self, test_size: int = 1000):
        """
        Простой бенчмарк для тестирования производительности

        Args:
            test_size: Количество строк в тестовом файле
        """
        self.test_size = test_size
        self.test_file = f"simple_test_{test_size}.xlsx"
        print(f"🚀 Инициализация бенчмарка для {test_size} строк")

    def generate_simple_data(self) -> List[List[Any]]:
        """Генерация простых тестовых данных"""
        print(f"📊 Генерация {self.test_size} строк тестовых данных...")

        data = [["ID", "Name", "Age", "Score"]]  # Заголовки

        names = ["Alice", "Bob", "Charlie", "Diana", "Eve", "Frank", "Grace", "Henry"]

        for i in range(1, self.test_size + 1):
            row = [
                i,
                random.choice(names),
                random.randint(20, 60),
                round(random.uniform(0, 100), 2)
            ]
            data.append(row)

        return data

    def create_test_file_cytosheet(self, data: List[List[Any]]) -> None:
        """Создание тестового файла с помощью cytosheet"""
        print(f"📝 Создание файла {self.test_file} (cytosheet)...")

        try:
            wb = Workbook()
            ws = wb.active

            for row_idx, row_data in enumerate(data, 1):
                for col_idx, value in enumerate(row_data, 1):
                    # Простое преобразование индекса в букву колонки
                    if col_idx <= 26:
                        col_letter = chr(64 + col_idx)  # A=65, B=66, etc.
                        cell_addr = f"{col_letter}{row_idx}"
                        ws[cell_addr].value = value

            wb.save(self.test_file)
            print(f"✅ Файл {self.test_file} создан успешно!")

        except Exception as e:
            print(f"❌ Ошибка создания файла: {e}")
            traceback.print_exc()
            raise

    def time_function(self, func, *args, **kwargs) -> Tuple[Any, float]:
        """Измерение времени выполнения функции"""
        gc.collect()  # Очистка мусора перед тестом

        start_time = time.perf_counter()
        try:
            result = func(*args, **kwargs)
            end_time = time.perf_counter()
            return result, end_time - start_time
        except Exception as e:
            end_time = time.perf_counter()
            print(f"❌ Ошибка в функции: {e}")
            return None, end_time - start_time

    # ================== ТЕСТ 1: ЧТЕНИЕ ФАЙЛА ==================

    def test_read_cytosheet(self) -> Tuple[int, float]:
        """Тест чтения файла через cytosheet"""

        def read_operation():
            wb = load_workbook(self.test_file)
            ws = wb.get_sheet_by_name('Sheet')

            count = 0
            for cell_addr, cell in ws._cells.items():
                if cell.value is not None:
                    count += 1

            wb.close() if hasattr(wb, 'close') else None
            return count

        result, exec_time = self.time_function(read_operation)
        return result or 0, exec_time

    def test_read_openpyxl(self) -> Tuple[int, float]:
        """Тест чтения файла через openpyxl"""

        def read_operation():
            wb = openpyxl_load_workbook(self.test_file)
            ws = wb.active

            count = 0
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value is not None:
                        count += 1

            wb.close()
            return count

        result, exec_time = self.time_function(read_operation)
        return result or 0, exec_time

    # ================== ТЕСТ 2: ПОИСК ДАННЫХ ==================

    def test_search_cytosheet(self, search_name: str = "Alice") -> Tuple[int, float]:
        """Тест поиска данных через cytosheet"""

        def search_operation():
            wb = load_workbook(self.test_file, lazy=True)
            ws = wb.get_sheet_by_name('Sheet')

            found_count = 0
            try:
                rows = ws.iter_rows(values_only=True)
                next(rows)  # Пропускаем заголовки

                for row_data in rows:
                    if len(row_data) > 1 and row_data[1] == search_name:
                        found_count += 1
            except Exception as e:
                print(f"Ошибка в поиске cytosheet: {e}")
            finally:
                wb.close() if hasattr(wb, 'close') else None

            return found_count

        result, exec_time = self.time_function(search_operation)
        return result or 0, exec_time

    def test_search_openpyxl(self, search_name: str = "Alice") -> Tuple[int, float]:
        """Тест поиска данных через openpyxl"""

        def search_operation():
            wb = openpyxl_load_workbook(self.test_file, read_only=True)
            ws = wb.active

            found_count = 0
            rows = ws.iter_rows(values_only=True)
            next(rows)  # Пропускаем заголовки

            for row_data in rows:
                if len(row_data) > 1 and row_data[1] == search_name:
                    found_count += 1

            wb.close()
            return found_count

        result, exec_time = self.time_function(search_operation)
        return result or 0, exec_time

    # ================== ТЕСТ 3: СОЗДАНИЕ ФАЙЛА ==================

    def test_create_cytosheet(self, output_file: str) -> float:
        """Тест создания файла через cytosheet"""

        def create_operation():
            wb = Workbook()
            ws = wb.active

            # Простые данные
            data = [
                ["New_ID", "New_Name", "New_Value"],
                [1, "Test1", 100],
                [2, "Test2", 200],
                [3, "Test3", 300]
            ]

            for row_idx, row_data in enumerate(data, 1):
                for col_idx, value in enumerate(row_data, 1):
                    if col_idx <= 26:
                        col_letter = chr(64 + col_idx)
                        cell_addr = f"{col_letter}{row_idx}"
                        ws[cell_addr].value = value

            wb.save(output_file)

        _, exec_time = self.time_function(create_operation)
        return exec_time

    def test_create_openpyxl(self, output_file: str) -> float:
        """Тест создания файла через openpyxl"""

        def create_operation():
            wb = OpenpyxlWorkbook()
            ws = wb.active

            # Простые данные
            data = [
                ["New_ID", "New_Name", "New_Value"],
                [1, "Test1", 100],
                [2, "Test2", 200],
                [3, "Test3", 300]
            ]

            for row_data in data:
                ws.append(row_data)

            wb.save(output_file)

        _, exec_time = self.time_function(create_operation)
        return exec_time

    # ================== ЗАПУСК ВСЕХ ТЕСТОВ ==================

    def run_all_tests(self) -> Dict[str, Any]:
        """Запуск всех тестов"""
        print(f"\n{'=' * 60}")
        print(f"🎯 ЗАПУСК ТЕСТОВ ({self.test_size} строк)")
        print(f"{'=' * 60}")

        results = {}

        try:
            # Генерируем и создаем тестовый файл
            test_data = self.generate_simple_data()
            self.create_test_file_cytosheet(test_data)

            # ТЕСТ 1: Чтение файла
            print(f"\n📖 ТЕСТ 1: Чтение файла")
            print("-" * 30)

            print("  Cytosheet...")
            cyto_read_count, cyto_read_time = self.test_read_cytosheet()
            print(f"    Время: {cyto_read_time:.4f}s, Ячеек: {cyto_read_count}")

            print("  OpenPyXL...")
            openpyxl_read_count, openpyxl_read_time = self.test_read_openpyxl()
            print(f"    Время: {openpyxl_read_time:.4f}s, Ячеек: {openpyxl_read_count}")

            results['read_test'] = {
                'cytosheet': {'time': cyto_read_time, 'cells': cyto_read_count},
                'openpyxl': {'time': openpyxl_read_time, 'cells': openpyxl_read_count}
            }

            # ТЕСТ 2: Поиск данных
            print(f"\n🔍 ТЕСТ 2: Поиск данных")
            print("-" * 30)

            search_name = "Alice"
            print(f"  Ищем: '{search_name}'")

            print("  Cytosheet...")
            cyto_search_count, cyto_search_time = self.test_search_cytosheet(search_name)
            print(f"    Время: {cyto_search_time:.4f}s, Найдено: {cyto_search_count}")

            print("  OpenPyXL...")
            openpyxl_search_count, openpyxl_search_time = self.test_search_openpyxl(search_name)
            print(f"    Время: {openpyxl_search_time:.4f}s, Найдено: {openpyxl_search_count}")

            results['search_test'] = {
                'cytosheet': {'time': cyto_search_time, 'found': cyto_search_count},
                'openpyxl': {'time': openpyxl_search_time, 'found': openpyxl_search_count}
            }

            # ТЕСТ 3: Создание файла
            print(f"\n📝 ТЕСТ 3: Создание файла")
            print("-" * 30)

            print("  Cytosheet...")
            cyto_create_time = self.test_create_cytosheet("test_create_cyto.xlsx")
            print(f"    Время: {cyto_create_time:.4f}s")

            print("  OpenPyXL...")
            openpyxl_create_time = self.test_create_openpyxl("test_create_openpyxl.xlsx")
            print(f"    Время: {openpyxl_create_time:.4f}s")

            results['create_test'] = {
                'cytosheet': {'time': cyto_create_time},
                'openpyxl': {'time': openpyxl_create_time}
            }

            # Показываем сравнение
            self.print_comparison(results)

            return results

        except Exception as e:
            print(f"❌ Ошибка в тестах: {e}")
            traceback.print_exc()
            return {}
        finally:
            # Очистка временных файлов
            self.cleanup_files()

    def print_comparison(self, results: Dict[str, Any]) -> None:
        """Вывод сравнения результатов"""
        print(f"\n{'=' * 60}")
        print("📊 СРАВНЕНИЕ РЕЗУЛЬТАТОВ")
        print(f"{'=' * 60}")

        for test_name, data in results.items():
            if 'cytosheet' in data and 'openpyxl' in data:
                cyto_time = data['cytosheet']['time']
                openpyxl_time = data['openpyxl']['time']

                if cyto_time > 0 and openpyxl_time > 0:
                    speedup = openpyxl_time / cyto_time
                    faster = "Cytosheet" if speedup > 1 else "OpenPyXL"

                    print(f"\n{test_name.replace('_', ' ').title()}:")
                    print(f"  Cytosheet: {cyto_time:.4f}s")
                    print(f"  OpenPyXL:  {openpyxl_time:.4f}s")
                    print(f"  Быстрее:   {faster} в {abs(speedup):.2f}x раз")

    def cleanup_files(self) -> None:
        """Очистка временных файлов"""
        files_to_remove = [
            self.test_file,
            "test_create_cyto.xlsx",
            "test_create_openpyxl.xlsx"
        ]

        for file in files_to_remove:
            try:
                if os.path.exists(file):
                    os.remove(file)
                    print(f"🗑 Удален: {file}")
            except Exception as e:
                print(f"⚠ Не удалось удалить {file}: {e}")


def main():
    """Главная функция"""
    print("🚀 ПРОСТЫЕ БЕНЧМАРКИ CYTOSHEET vs OPENPYXL")
    print("=" * 60)

    # Различные размеры для тестирования
    test_sizes = [100, 500, 1000, 10000, 100000]

    all_results = {}

    for size in test_sizes:
        print(f"\n{'=' * 80}")
        print(f"🎯 ТЕСТИРОВАНИЕ С {size} СТРОКАМИ")
        print(f"{'=' * 80}")

        try:
            benchmark = SimpleBenchmark(test_size=size)
            results = benchmark.run_all_tests()
            all_results[size] = results

            print(f"✅ Тесты с {size} строками завершены")

        except Exception as e:
            print(f"❌ Ошибка в тестах с {size} строками: {e}")
            traceback.print_exc()

    # Общая сводка
    print(f"\n{'=' * 80}")
    print("🏆 ОБЩАЯ СВОДКА")
    print(f"{'=' * 80}")

    for size, results in all_results.items():
        if results:
            print(f"\n📊 {size} строк:")
            for test_name, data in results.items():
                if 'cytosheet' in data and 'openpyxl' in data:
                    cyto_time = data['cytosheet']['time']
                    openpyxl_time = data['openpyxl']['time']

                    if cyto_time > 0 and openpyxl_time > 0:
                        speedup = openpyxl_time / cyto_time
                        status = "быстрее" if speedup > 1 else "медленнее"
                        print(f"  {test_name}: Cytosheet {status} в {abs(speedup):.2f}x")


# Performance tests from test_performance.py

def test_merge_cells_performance():
    """Test the performance of merge_cells method."""
    # Create a new workbook
    wb = Workbook()
    ws = wb.active

    # Set values in cells
    ws['A1'] = 'Merged Cell'

    # Measure the time to merge cells
    start_time = time.time()
    for i in range(1, 101):
        range_string = f'A{i}:C{i+2}'
        merged_cell = ws.merge_cells(range_string)
    end_time = time.time()

    merge_time = end_time - start_time
    print(f"Time to merge 100 cell ranges: {merge_time:.4f} seconds")

    # Measure the time to unmerge cells
    start_time = time.time()
    for i in range(1, 101):
        range_string = f'A{i}:C{i+2}'
        ws.unmerge_cells(range_string)
    end_time = time.time()

    unmerge_time = end_time - start_time
    print(f"Time to unmerge 100 cell ranges: {unmerge_time:.4f} seconds")

    # Assert that the operations completed in a reasonable time
    assert merge_time < 1.0, f"Merge operation took too long: {merge_time:.4f} seconds"
    assert unmerge_time < 1.0, f"Unmerge operation took too long: {unmerge_time:.4f} seconds"

def test_style_performance():
    """Test the performance of style-related methods."""
    # Create a new workbook
    wb = Workbook()
    ws = wb.active

    # Create a style
    font = Font(name='Arial', size=12, bold=True, italic=True)
    border = Border(
        left=Side(style='thin', color=Color(rgb='FF0000')),
        right=Side(style='thin', color=Color(rgb='FF0000')),
        top=Side(style='thin', color=Color(rgb='FF0000')),
        bottom=Side(style='thin', color=Color(rgb='FF0000'))
    )
    fill = PatternFill(patternType='solid', fgColor=Color(rgb='FFFF00'))
    alignment = Alignment(horizontal='center', vertical='center')
    protection = Protection(locked=True, hidden=False)
    style = Style(font=font, border=border, fill=fill, alignment=alignment, protection=protection)

    # Measure the time to apply styles to cells
    start_time = time.time()
    for i in range(1, 1001):
        cell_ref = f'A{i}'
        ws[cell_ref] = f'Cell {i}'
        ws[cell_ref].style = style
    end_time = time.time()

    style_time = end_time - start_time
    print(f"Time to apply styles to 1000 cells: {style_time:.4f} seconds")

    # Measure the time to access style properties
    start_time = time.time()
    for i in range(1, 1001):
        cell_ref = f'A{i}'
        font = ws[cell_ref].style.font
        border = ws[cell_ref].style.border
        fill = ws[cell_ref].style.fill
        alignment = ws[cell_ref].style.alignment
        protection = ws[cell_ref].style.protection
    end_time = time.time()

    access_time = end_time - start_time
    print(f"Time to access style properties of 1000 cells: {access_time:.4f} seconds")

    # Assert that the operations completed in a reasonable time
    assert style_time < 2.0, f"Style application took too long: {style_time:.4f} seconds"
    assert access_time < 2.0, f"Style access took too long: {access_time:.4f} seconds"

def test_border_add_performance():
    """Test the performance of the Border.__add__ method."""
    # Create borders
    border1 = Border(
        left=Side(style='thin', color=Color(rgb='FF0000')),
        right=Side(style='thin', color=Color(rgb='FF0000'))
    )
    border2 = Border(
        top=Side(style='thin', color=Color(rgb='0000FF')),
        bottom=Side(style='thin', color=Color(rgb='0000FF'))
    )

    # Measure the time to add borders
    start_time = time.time()
    for _ in range(10000):
        result = border1 + border2
    end_time = time.time()

    add_time = end_time - start_time
    print(f"Time to add borders 10000 times: {add_time:.4f} seconds")

    # Assert that the operation completed in a reasonable time
    assert add_time < 1.0, f"Border addition took too long: {add_time:.4f} seconds"

if __name__ == "__main__":
    main()
    # Uncomment to run performance tests
    # test_merge_cells_performance()
    # test_style_performance()
    # test_border_add_performance()
