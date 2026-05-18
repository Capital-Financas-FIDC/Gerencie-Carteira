# Build script do core Python (PyInstaller) para empacotar dentro do Electron.
# Uso (PowerShell):
#   cd backend
#   ./build_core.ps1

$ErrorActionPreference = "Stop"

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$srcEntry = Join-Path $here "src/gerencie_carteira.py"
$buildRoot = Join-Path $here "../build/dist-python"
$workPath = Join-Path $here "../build/pyinstaller-work"
$specPath = Join-Path $here "../build/pyinstaller-spec"
$appResources = Join-Path $here "../app/resources"

Write-Host "Compilando gerencie_carteira_core.exe com PyInstaller..." -ForegroundColor Cyan

pyinstaller `
    --onefile `
    --console `
    --name "gerencie_carteira_core" `
    --distpath $buildRoot `
    --workpath $workPath `
    --specpath $specPath `
    --paths "$here/src" `
    --hidden-import "win32com.client" `
    --hidden-import "win32com.gen_py" `
    $srcEntry

if (-not $?) {
    Write-Host "PyInstaller falhou." -ForegroundColor Red
    exit 1
}

$exeSrc = Join-Path $buildRoot "gerencie_carteira_core.exe"
$exeDst = Join-Path $appResources "gerencie_carteira_core.exe"

if (-not (Test-Path $appResources)) { New-Item -ItemType Directory -Path $appResources | Out-Null }

Copy-Item -Force $exeSrc $exeDst
Write-Host "OK: $exeDst" -ForegroundColor Green

Write-Host ""
Write-Host "Para validar o binario isoladamente:" -ForegroundColor Yellow
Write-Host "  & '$exeDst'"
Write-Host ""
Write-Host "Para gerar o installer Electron:" -ForegroundColor Yellow
Write-Host "  cd $here/../app; npm run dist"
