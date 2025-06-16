from zipfile import ZipFile
from src.cytosheet import Workbook

def load_workbook(str filename, bint lazy = False):
    """
    Оптимизированная загрузка .xlsx файла.

    Args:
        filename: Путь к .xlsx файлу
        lazy: Если True, листы загружаются по требованию для экономии памяти

    Returns:
        Workbook: Объект рабочей книги
    """
    cdef object archive = ZipFile(filename, "r")
    return Workbook(_archive=archive, lazy=lazy)