# ==============================================================================
# build-app.ps1 — Compila o app completo e entrega em .\Aplicativo\
#
# Etapas: core Python (PyInstaller) -> build TS/Vite -> electron-builder (dir)
#         -> relocacao para .\Aplicativo (fora de build/, facil de achar).
#
# O electron-builder falha ao extrair symlinks macOS do cache winCodeSign em
# maquinas sem privilegio de symlink. Isso so afeta o instalador NSIS — a
# pasta win-unpacked e gerada normalmente. Este script trata esse erro como
# nao-fatal e prossegue se o build foi produzido.
#
# Uso (PowerShell, na raiz do repo):
#   ./build-app.ps1
# ==============================================================================

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host "[1/4] Core Python (PyInstaller)..." -ForegroundColor Cyan
& "$root\backend\build_core.ps1"
if (-not $?) { throw "build_core.ps1 falhou" }

Write-Host "[2/4] Build TypeScript + Vite..." -ForegroundColor Cyan
Push-Location "$root\app"
try {
    if (-not (Test-Path "node_modules")) { npm install }
    npm run build
    if (-not $?) { throw "npm run build falhou" }

    Write-Host "[3/4] electron-builder (pasta win-unpacked)..." -ForegroundColor Cyan
    # Erro de symlink winCodeSign e esperado e nao-fatal; ignorado aqui.
    try { npx electron-builder --win --x64 --dir } catch { }
}
finally { Pop-Location }

$unpacked = "$root\build\dist-electron\win-unpacked"
if (-not (Test-Path "$unpacked\Gerencie Carteira.exe")) {
    throw "win-unpacked nao foi gerado - verifique os logs acima."
}

Write-Host "[4/4] Relocando para .\Aplicativo ..." -ForegroundColor Cyan
$dst = "$root\Aplicativo"
# /MIR espelha (remove arquivos obsoletos no destino). Mantem a fonte em
# build/ (gitignored, sobrescrita no proximo build) para evitar lock no delete.
robocopy $unpacked $dst /MIR /R:2 /W:2 /NFL /NDL /NJH /NP | Out-Null
if ($LASTEXITCODE -ge 8) { throw "robocopy falhou (codigo $LASTEXITCODE)" }

Write-Host ""
Write-Host "OK -> $dst\Gerencie Carteira.exe" -ForegroundColor Green
Write-Host "Atalho: clique duplo em '.\Gerencie Carteira.cmd' na raiz." -ForegroundColor Yellow
