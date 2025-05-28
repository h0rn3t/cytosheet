from Cython.Build import cythonize
from setuptools import setup

setup(
    ext_modules=cythonize(
        [
            "cytosheet/cell.pyx",
            "cytosheet/workbook.pyx",
            "cytosheet/worksheet.pyx",
            "cytosheet/excel.pyx",
        ],
        compiler_directives={"language_level": "3"},
    ),
    zip_safe=False,
)
