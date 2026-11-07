# Windows packaging

Build on Windows using Python with Tk support and a clean virtual environment.

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -e . pyinstaller
pyinstaller --onefile --windowed --name MedarYakaKart --paths src scripts/launch_app.py
```

The output is `dist/MedarYakaKart.exe`. `scripts/launch_app.py` imports the package
by its absolute name; the package's relative `__main__` import is not used as a
standalone script. No icon flag is supplied because the previously documented
`assets/medar.ico` is not present. Add an icon only after providing a licensed asset.

If enabling RAR or drag/drop, install and package the optional dependencies and
review the external extractor's distribution terms. Test the executable on a
clean Windows account with synthetic PDFs, writable settings/output paths and
an actual duplex printer before distributing it. The audit did not build a
Windows executable or validate printer alignment.
