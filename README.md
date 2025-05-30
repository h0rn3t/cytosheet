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

First install the build requirements and compile the Cython extensions:

```
pip install Cython lxml
python src/setup.py build_ext --inplace
```

Optionally you can build a wheel distribution:

```
python -m build
```

### Running tests


After building the Cython extensions you can execute the test suite. Add the
`src` directory to `PYTHONPATH` or install the package in editable mode:

```
export PYTHONPATH=$(pwd)/src
```

Install `pytest` and (optionally) `openpyxl` for the performance comparison:

```bash
pip install pytest
# Optional dependency for performance comparison
pip install openpyxl
```

