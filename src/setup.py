from setuptools import setup
from Cython.Build import cythonize

setup(
    ext_modules=cythonize(
        ["cytosheet/cell.pyx", "cytosheet/workbook.pyx", "cytosheet/worksheet.pyx"],
        compiler_directives={'language_level': "3"}
    ),
    zip_safe=False,
)
