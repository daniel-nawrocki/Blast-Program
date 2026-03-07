# Blast-Program

Section 1 scaffold for a downloadable desktop app using Python + PySide6.

Note: `Gassing Calculator` now runs natively in-app (no Excel dependency).

## Included in this section

- Start Screen with navigation
- Vibration Tool menu
- Vibration Calculator
- Site Factor Calibrator
- References/Cheat Sheets
- Gassing Calculator (native workbook-style layout)
- Placeholders for:
  - Empirical Formula Calculation
  - Diagram Maker

## Local run

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
$env:PYTHONPATH="src"
python main.py
```

## Portable file locations

The app now loads external files from project-relative paths so it works on other machines.

- Workbook: `assets/workbooks/Diff Gassing Titan XL Blends, 02 October 2012, Final.xlsm`
- References:
  - `assets/references/Empirical Cheat Sheet.docx`
  - `assets/references/Open Pit Book.pdf`

Optional override:

```powershell
$env:BLAST_WORKBOOK_PATH="C:\path\to\your\workbook.xlsm"
```

## Build Windows executable

```powershell
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
.\build.ps1
```

`build.ps1` automatically bundles `assets/` into the app output when the folder exists.

Output executable:

- `dist\BlastProgram\BlastProgram.exe`
