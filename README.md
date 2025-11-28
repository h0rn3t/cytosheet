# Cytosheet

Cytosheet is a PoC library for working with XLSX files, written in Cython.

## Roadmap

1. [] High performance due to Cython and libxml2.
2. [] Support for reading large XLSX files using SAX parsing, allowing files to be processed in chunks without loading the entire file into memory.
3. [] Support for creating and writing XLSX files.
4. [] Lightweight and user-friendly API, similar to openpyxl.

## Installation

### Requirements

- Python 3.7+
- Cython
- lxml
### Install from source

   pip install .
   python -m build
   python setup.py build_ext --inplace
