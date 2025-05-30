from Cython.Build import cythonize
from setuptools import setup

setup(
    ext_modules=cythonize(
        [
            "cytosheet/cell.pyx",
            "cytosheet/workbook.pyx",
            "cytosheet/worksheet.pyx",
            "cytosheet/excel.pyx",
            "openpyxl_c/workbook.pyx",
            "openpyxl_c/worksheet.pyx",
            "openpyxl_c/cell.pyx",
            "openpyxl_c/styles.pyx",
            "openpyxl_c/utils.pyx",
            "openpyxl_c/reader.pyx",
            "openpyxl_c/writer.pyx",
            "openpyxl_c/chart/__init__.pyx",
            "openpyxl_c/comments.pyx",
            "openpyxl_c/validation.pyx",
            "openpyxl_c/formatting.pyx",
        ],
        compiler_directives={"language_level": "3"},
    ),
    zip_safe=False,
)
