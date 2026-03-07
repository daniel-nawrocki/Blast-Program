param(
    [string]$AppName = "BlastProgram"
)

$assetsPath = Join-Path $PSScriptRoot "assets"
$addDataArg = $null
if (Test-Path $assetsPath) {
    # Windows PyInstaller expects "source;dest" for --add-data.
    $addDataArg = "$assetsPath;assets"
}

$pyinstallerArgs = @(
    "--noconfirm",
    "--clean",
    "--windowed",
    "--name", $AppName,
    "--paths", "src"
)

if ($addDataArg) {
    $pyinstallerArgs += @("--add-data", $addDataArg)
}

$pyinstallerArgs += "main.py"

python -m PyInstaller @pyinstallerArgs
