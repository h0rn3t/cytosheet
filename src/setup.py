from Cython.Build import cythonize
from setuptools import setup

setup(
    ext_modules=cythonize(
        [
            "cytosheet/*.pyx",
            "openpyxl_c/*.pyx",
            "openpyxl_c/chart/*.pyx",
        ],
        compiler_directives={"language_level": "3"},
    ),
    zip_safe=False,
)
