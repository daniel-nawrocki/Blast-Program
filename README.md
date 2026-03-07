# Blast-Program

Section 1 scaffold for a downloadable desktop app using Python + PySide6.

## Included in this section

- Start Screen with navigation
- Vibration Tool menu
- Placeholders for:
  - Vibration Estimate
  - Timing Solver
  - Empirical Formula Calculation
  - Diagram Maker
  - Small Quick Cheat Sheets

## Local run

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
$env:PYTHONPATH="src"
python main.py
```

## Build Windows executable

```powershell
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
.\build.ps1
```

Output executable:

- `dist\BlastProgram\BlastProgram.exe`
