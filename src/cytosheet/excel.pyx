from zipfile import ZipFile
from .workbook import Workbook

# Совместимый с openpyxl API: поддерживаем read_only, а также синоним lazy для внутренних тестов.
# Остальные аргументы принимаем и игнорируем без ошибки.
def load_workbook(
    str filename,
    bint read_only=False,
    bint keep_vba=False,
    bint data_only=False,
    bint keep_links=True,
    bint rich_text=False,
    bint lazy=False,
):
    """Загружаем XLSX и инициализируем Workbook с архивом.

    Параметры сумісні з openpyxl.load_workbook, плюс внутрішній синонім
    `lazy`, який маппится на internal `Workbook.lazy`.

    Зараз використовується тільки read_only / lazy. Інші аргументи приймаються
    і ігноруються без помилки.
    """
    cdef object archive = ZipFile(filename, 'r')
    # Поддерживаем оба способа задания lazy/read_only: если явно указан lazy=True,
    # он имеет приоритет над read_only для внутреннего флага.
    cdef bint use_lazy = lazy or read_only
    wb = Workbook(_archive=archive, lazy=use_lazy)
    # даём Worksheet доступ к _xf_numfmt_map через ссылку на книгу
    setattr(archive, '_workbook_ref', wb)
    return wb
