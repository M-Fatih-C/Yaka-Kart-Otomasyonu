# Medar Employee Badge Generator

A Tkinter desktop tool that arranges front/back PDF badge templates into an A4
DOCX document for duplex printing. It supports previews, ordering, profiles,
layout/DPI settings, statistics and PDF extraction from ZIP/RAR inputs.
This checkout does not implement an Excel-driven employee-data merge workflow.

## Install and run

Use Python 3.10+ with Tk support, preferably in a virtual environment:

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
python -m pip install -e .
python -m medar_yakakart
```

Installing only `requirements.txt` does not install the `src/medar_yakakart`
package. The editable install above installs the package and its declared
PyMuPDF, python-docx and Pillow dependencies. The installed `medar-yakakart`
command is an alternative entry point. Optional RAR/drag-and-drop support needs
the corresponding libraries and, for RAR, a supported external extraction tool.

## Printing workflow

1. Select authorized PDF templates or a trusted archive containing them.
2. Order the files and choose card size, DPI, margins and cards per page.
3. Generate the DOCX and inspect front/back alignment before printing.
4. Print a synthetic test sheet at A4 / 100% scale / duplex long-edge first.

Printer drivers and Word versions can change layout. Real duplex alignment and
Windows executable packaging require manual platform acceptance; they were not
verified by the audit. [Windows build instructions](docs/BUILD.md) use a real
launcher and do not depend on a missing icon asset.

## Layout and configuration

`src/medar_yakakart/app.py` contains the existing GUI/layout logic; `__main__.py`
is its module entry. `config.example.json`, `profiles.example.json` and
`stats.example.json` document runtime JSON files created in the working directory.
Keep a writable working directory and avoid committing real settings or outputs.

The CI performs packaging/import/syntax checks, not a GUI or print-quality test.
No automated functional print suite currently exists. Refactoring the large GUI
module is deferred to preserve the existing layout and import behavior.

## Privacy and license

Badge PDFs, photographs, names, generated Word files and logs can identify
employees. Use synthetic examples for development and keep real data outside Git.
Process only trusted archives/documents. See [LICENSE](LICENSE) for the existing
MIT license. Author: [Muhammet Fatih Çetintaş](https://github.com/M-Fatih-C).
