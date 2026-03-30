$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

$python = Join-Path $root '.venv\Scripts\python.exe'
if (-not (Test-Path $python)) {
    throw 'Virtual environment Python was not found at .venv\Scripts\python.exe'
}

& $python -m pip install --upgrade pyinstaller
& $python -m PyInstaller --noconfirm --clean Reliability.spec

Write-Host ''
Write-Host 'Build complete.'
Write-Host 'Executable:' (Join-Path $root 'dist\Reliability.exe')
