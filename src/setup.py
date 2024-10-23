from setuptools import setup
from Cython.Build import cythonize

setup(
    ext_modules=cythonize(["src/cyxls.pyx", "src/parser.pyx", "src/writer.pyx"]),
    # include_dirs=['/usr/include/libxml2'],
    zip_safe=False,
)
