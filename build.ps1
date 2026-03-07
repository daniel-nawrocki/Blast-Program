param(
    [string]$AppName = "BlastProgram"
)

python -m PyInstaller `
    --noconfirm `
    --clean `
    --windowed `
    --name $AppName `
    --paths src `
    main.py
