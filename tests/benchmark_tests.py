"""
Спрощені бенчмарки для локального тестування cytosheet vs openpyxl
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
    print(f"Помилка імпорту: {e}")
    print("Переконайтеся, що всі залежності встановлені та cytosheet скомпільований")
    sys.exit(1)


class SimpleBenchmark:
    def __init__(self, test_size: int = 1000):
        """
        Простий бенчмарк для тестування продуктивності

        Args:
            test_size: Кількість рядків у тестовому файлі
        """
        self.test_size = test_size
        self.test_file = f"simple_test_{test_size}.xlsx"
        print(f"🚀 Ініціалізація бенчмарка для {test_size} рядків")

    def generate_simple_data(self) -> List[List[Any]]:
        """Генерація простих тестових даних"""
        print(f"📊 Генерація {self.test_size} рядків тестових даних...")

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
        """Створення тестового файлу за допомогою cytosheet"""
        print(f"📝 Створення файлу {self.test_file} (cytosheet)...")

        try:
            wb = Workbook()
            ws = wb.active

            for row_idx, row_data in enumerate(data, 1):
                for col_idx, value in enumerate(row_data, 1):
                    # Використовуємо openpyxl-сумісне API .cell(row, column)
                    ws.cell(row=row_idx, column=col_idx, value=value)

            wb.save(self.test_file)
            print(f"✅ Файл {self.test_file} створено успішно!")

        except Exception as e:
            print(f"❌ Помилка створення файлу: {e}")
            traceback.print_exc()
            raise

    def time_function(self, func, *args, **kwargs) -> Tuple[Any, float]:
        """Вимірювання часу виконання функції"""
        gc.collect()  # Очищення сміття перед тестом

        start_time = time.perf_counter()
        try:
            result = func(*args, **kwargs)
            end_time = time.perf_counter()
            return result, end_time - start_time
        except Exception as e:
            end_time = time.perf_counter()
            print(f"❌ Помилка у функції: {e}")
            return None, end_time - start_time

    # ================== ТЕСТ 1: ЧИТАННЯ ФАЙЛУ ==================

    def test_read_cytosheet(self) -> Tuple[int, float]:
        """Тест читання файлу через cytosheet"""

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
        """Тест читання файлу через openpyxl"""

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

    # ================== ТЕСТ 2: ПОШУК ДАНИХ ==================

    def test_search_cytosheet(self, search_name: str = "Alice") -> Tuple[int, float]:
        """Тест пошуку даних через cytosheet"""

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
                print(f"Помилка у пошуку cytosheet: {e}")
            finally:
                wb.close() if hasattr(wb, 'close') else None

            return found_count

        result, exec_time = self.time_function(search_operation)
        return result or 0, exec_time

    def test_search_openpyxl(self, search_name: str = "Alice") -> Tuple[int, float]:
        """Тест пошуку даних через openpyxl"""

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

    # ================== ТЕСТ 3: СТВОРЕННЯ ФАЙЛУ ==================

    def test_create_cytosheet(self, output_file: str) -> float:
        """Тест створення файлу через cytosheet"""

        def create_operation():
            wb = Workbook()
            ws = wb.active

            # Прості дані
            data = [
                ["New_ID", "New_Name", "New_Value"],
                [1, "Test1", 100],
                [2, "Test2", 200],
                [3, "Test3", 300]
            ]

            for row_idx, row_data in enumerate(data, 1):
                for col_idx, value in enumerate(row_data, 1):
                    ws.cell(row=row_idx, column=col_idx, value=value)

            wb.save(output_file)

        _, exec_time = self.time_function(create_operation)
        return exec_time

    def test_create_openpyxl(self, output_file: str) -> float:
        """Тест створення файлу через openpyxl"""

        def create_operation():
            wb = OpenpyxlWorkbook()
            ws = wb.active

            # Прості дані
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

    # ================== ЗАПУСК УСІХ ТЕСТІВ ==================

    def run_all_tests(self) -> Dict[str, Any]:
        """Запуск усіх тестів"""
        print(f"\n{'=' * 60}")
        print(f"🎯 ЗАПУСК ТЕСТІВ ({self.test_size} рядків)")
        print(f"{'=' * 60}")

        results = {}

        try:
            # Генеруємо та створюємо тестовий файл
            test_data = self.generate_simple_data()
            self.create_test_file_cytosheet(test_data)

            # # ТЕСТ 1: Читання файлу
            # print(f"\n📖 ТЕСТ 1: Читання файлу")
            # print("-" * 30)

            # print("  Cytosheet...")
            # cyto_read_count, cyto_read_time = self.test_read_cytosheet()
            # print(f"    Час: {cyto_read_time:.4f}s, Комірок: {cyto_read_count}")

            # print("  OpenPyXL...")
            # openpyxl_read_count, openpyxl_read_time = self.test_read_openpyxl()
            # print(f"    Час: {openpyxl_read_time:.4f}s, Комірок: {openpyxl_read_count}")

            # results['read_test'] = {
            #     'cytosheet': {'time': cyto_read_time, 'cells': cyto_read_count},
            #     'openpyxl': {'time': openpyxl_read_time, 'cells': openpyxl_read_count}
            # }

            # # ТЕСТ 2: Пошук даних
            # print(f"\n🔍 ТЕСТ 2: Пошук даних")
            # print("-" * 30)

            # search_name = "Alice"
            # print(f"  Шукаємо: '{search_name}'")

            # print("  Cytosheet...")
            # cyto_search_count, cyto_search_time = self.test_search_cytosheet(search_name)
            # print(f"    Час: {cyto_search_time:.4f}s, Знайдено: {cyto_search_count}")

            # print("  OpenPyXL...")
            # openpyxl_search_count, openpyxl_search_time = self.test_search_openpyxl(search_name)
            # print(f"    Час: {openpyxl_search_time:.4f}s, Знайдено: {openpyxl_search_count}")

            # results['search_test'] = {
            #     'cytosheet': {'time': cyto_search_time, 'found': cyto_search_count},
            #     'openpyxl': {'time': openpyxl_search_time, 'found': openpyxl_search_count}
            # }

            # ТЕСТ 3: Створення файлу
            print(f"\n📝 ТЕСТ 3: Створення файлу")
            print("-" * 30)

            print("  Cytosheet...")
            cyto_create_time = self.test_create_cytosheet("test_create_cyto.xlsx")
            print(f"    Час: {cyto_create_time:.4f}s")

            print("  OpenPyXL...")
            openpyxl_create_time = self.test_create_openpyxl("test_create_openpyxl.xlsx")
            print(f"    Час: {openpyxl_create_time:.4f}s")

            results['create_test'] = {
                'cytosheet': {'time': cyto_create_time},
                'openpyxl': {'time': openpyxl_create_time}
            }

            # Показуємо порівняння
            self.print_comparison(results)

            return results

        except Exception as e:
            print(f"❌ Помилка у тестах: {e}")
            traceback.print_exc()
            return {}
        finally:
            # Очищення тимчасових файлів
            self.cleanup_files()

    def print_comparison(self, results: Dict[str, Any]) -> None:
        """Виведення порівняння результатів"""
        print(f"\n{'=' * 60}")
        print("📊 ПОРІВНЯННЯ РЕЗУЛЬТАТІВ")
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
                    print(f"  Швидше:   {faster} у {abs(speedup):.2f}x разів")

    def cleanup_files(self) -> None:
        """Очищення тимчасових файлів"""
        files_to_remove = [
            self.test_file,
            "test_create_cyto.xlsx",
            "test_create_openpyxl.xlsx"
        ]

        for file in files_to_remove:
            try:
                if os.path.exists(file):
                    os.remove(file)
                    print(f"🗑 Видалено: {file}")
            except Exception as e:
                print(f"⚠ Не вдалося видалити {file}: {e}")


def main():
    """Головна функція"""
    print("🚀 ПРОСТІ БЕНЧМАРКИ CYTOSHEET vs OPENPYXL")
    print("=" * 60)

    # Різні розміри для тестування
    test_sizes = [100, 500, 1000, 10000, 100000]

    all_results = {}

    for size in test_sizes:
        print(f"\n{'=' * 80}")
        print(f"🎯 ТЕСТУВАННЯ З {size} РЯДКАМИ")
        print(f"{'=' * 80}")

        try:
            benchmark = SimpleBenchmark(test_size=size)
            results = benchmark.run_all_tests()
            all_results[size] = results

            print(f"✅ Тести з {size} рядками завершено")

        except Exception as e:
            print(f"❌ Помилка у тестах з {size} рядками: {e}")
            traceback.print_exc()

    # Загальна зведена інформація
    print(f"\n{'=' * 80}")
    print("🏆 ЗАГАЛЬНА ЗВЕДЕНА ІНФОРМАЦІЯ")
    print(f"{'=' * 80}")

    for size, results in all_results.items():
        if results:
            print(f"\n📊 {size} рядків:")
            for test_name, data in results.items():
                if 'cytosheet' in data and 'openpyxl' in data:
                    cyto_time = data['cytosheet']['time']
                    openpyxl_time = data['openpyxl']['time']

                    if cyto_time > 0 and openpyxl_time > 0:
                        speedup = openpyxl_time / cyto_time
                        status = "швидше" if speedup > 1 else "повільніше"
                        print(f"  {test_name}: Cytosheet {status} у {abs(speedup):.2f}x")


# Тести продуктивності з test_performance.py
#
# def test_merge_cells_performance():
#     """Тест продуктивності методу merge_cells."""
#     # Створюємо нову книгу
#     wb = Workbook()
#     ws = wb.active
#
#     # Встановлюємо значення у комірки
#     ws['A1'] = 'Merged Cell'
#
#     # Вимірюємо час об'єднання комірок
#     start_time = time.time()
#     for i in range(1, 101):
#         range_string = f'A{i}:C{i+2}'
#         merged_cell = ws.merge_cells(range_string)
#     end_time = time.time()
#
#     merge_time = end_time - start_time
#     print(f"Час об'єднання 100 діапазонів комірок: {merge_time:.4f} секунд")
#
#     # Вимірюємо час роз'єднання комірок
#     start_time = time.time()
#     for i in range(1, 101):
#         range_string = f'A{i}:C{i+2}'
#         ws.unmerge_cells(range_string)
#     end_time = time.time()
#
#     unmerge_time = end_time - start_time
#     print(f"Час роз'єднання 100 діапазонів комірок: {unmerge_time:.4f} секунд")
#
#     # Перевіряємо, що операції завершилися за розумний час
#     assert merge_time < 1.0, f"Операція об'єднання зайняла занадто багато часу: {merge_time:.4f} секунд"
#     assert unmerge_time < 1.0, f"Операція роз'єднання зайняла занадто багато часу: {unmerge_time:.4f} секунд"
#
# def test_style_performance():
#     """Тест продуктивності методів, пов'язаних зі стилями."""
#     # Створюємо нову книгу
#     wb = Workbook()
#     ws = wb.active
#
#     # Створюємо стиль
#     font = Font(name='Arial', size=12, bold=True, italic=True)
#     border = Border(
#         left=Side(style='thin', color=Color(rgb='FF0000')),
#         right=Side(style='thin', color=Color(rgb='FF0000')),
#         top=Side(style='thin', color=Color(rgb='FF0000')),
#         bottom=Side(style='thin', color=Color(rgb='FF0000'))
#     )
#     fill = PatternFill(patternType='solid', fgColor=Color(rgb='FFFF00'))
#     alignment = Alignment(horizontal='center', vertical='center')
#     protection = Protection(locked=True, hidden=False)
#     style = Style(font=font, border=border, fill=fill, alignment=alignment, protection=protection)
#
#     # Вимірюємо час застосування стилів до комірок
#     start_time = time.time()
#     for i in range(1, 1001):
#         cell_ref = f'A{i}'
#         ws[cell_ref] = f'Cell {i}'
#         ws[cell_ref].style = style
#     end_time = time.time()
#
#     style_time = end_time - start_time
#     print(f"Час застосування стилів до 1000 комірок: {style_time:.4f} секунд")
#
#     # Вимірюємо час доступу до властивостей стилю
#     start_time = time.time()
#     for i in range(1, 1001):
#         cell_ref = f'A{i}'
#         font = ws[cell_ref].style.font
#         border = ws[cell_ref].style.border
#         fill = ws[cell_ref].style.fill
#         alignment = ws[cell_ref].style.alignment
#         protection = ws[cell_ref].style.protection
#     end_time = time.time()
#
#     access_time = end_time - start_time
#     print(f"Час доступу до властивостей стилю 1000 комірок: {access_time:.4f} секунд")
#
#     # Перевіряємо, що операції завершилися за розумний час
#     assert style_time < 2.0, f"Застосування стилю зайняло занадто багато часу: {style_time:.4f} секунд"
#     assert access_time < 2.0, f"Доступ до стилю зайняв занадто багато часу: {access_time:.4f} секунд"
#
# def test_border_add_performance():
#     """Тест продуктивності методу Border.__add__."""
#     # Створюємо рамки
#     border1 = Border(
#         left=Side(style='thin', color=Color(rgb='FF0000')),
#         right=Side(style='thin', color=Color(rgb='FF0000'))
#     )
#     border2 = Border(
#         top=Side(style='thin', color=Color(rgb='0000FF')),
#         bottom=Side(style='thin', color=Color(rgb='0000FF'))
#     )
#
#     # Вимірюємо час додавання рамок
#     start_time = time.time()
#     for _ in range(10000):
#         result = border1 + border2
#     end_time = time.time()
#
#     add_time = end_time - start_time
#     print(f"Час додавання рамок 10000 разів: {add_time:.4f} секунд")
#
#     # Перевіряємо, що операція завершилася за розумний час
#     assert add_time < 1.0, f"Додавання рамки зайняло занадто багато часу: {add_time:.4f} секунд"

if __name__ == "__main__":
    main()
    # Розкоментуйте для запуску тестів продуктивності
    # test_merge_cells_performance()
    # test_style_performance()
    # test_border_add_performance()
