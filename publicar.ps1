# ==============================================================================
# publicar.ps1 — Builda o app e publica na pasta de rede.
#
# Fluxo: build local (build-app.ps1 -> .\Aplicativo) -> espelha p/
#        A:\...\Software\Aplicativo -> garante o atalho .cmd na pasta-pai.
#
# - O /MIR do robocopy remove o build anterior automaticamente.
# - O .exe principal NAO e renomeado: Electron 33 quebra ASAR integrity
#   (crash 0x80000003 no boot) quando o exe principal carrega outro nome.
#   A versao do app continua visivel via `app:version` (IPC) na propria UI.
# - NAO usa NSIS: dispensa privilegio de admin / criacao de symlink.
#
# Uso (PowerShell, na raiz do repo):
#   ./publicar.ps1
# ==============================================================================

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

# Destino na rede. O "U" acentuado de PUBLICA e montado via code point para o
# script nao depender da codificacao do arquivo .ps1.
$pastaRede = "A:\PUBLICA\GERENCIE CARTEIRA P$([char]0xDA)BLICA"
$destAplicativo = Join-Path $pastaRede "Software\Aplicativo"

# [1/5] Versao — fonte unica: app/package.json
$pkg = Get-Content (Join-Path $root "app\package.json") -Raw | ConvertFrom-Json
$versao = $pkg.version
if (-not $versao) { throw "Nao foi possivel ler a versao de app/package.json" }
Write-Host "[1/5] Publicando versao $versao" -ForegroundColor Cyan

# [2/5] Build local (reusa build-app.ps1 -> .\Aplicativo)
Write-Host "[2/5] Build local (build-app.ps1)..." -ForegroundColor Cyan
& "$root\build-app.ps1"
$origem = Join-Path $root "Aplicativo"
if (-not (Test-Path -LiteralPath (Join-Path $origem "Gerencie Carteira.exe"))) {
    throw "Build local nao encontrado em $origem"
}

# Marcador de versao: o launcher .cmd compara este arquivo (share) com a copia
# local p/ decidir quando recopiar. Gravado em $origem ANTES do /MIR p/ a rede,
# entao viaja junto no espelhamento (e e recopiado p/ o local pelo launcher).
Set-Content -LiteralPath (Join-Path $origem "versao.txt") -Value $versao -Encoding ascii -NoNewline

# [3/5] Verifica a rede
if (-not (Test-Path -LiteralPath $pastaRede)) {
    throw "Pasta de rede inacessivel: $pastaRede"
}
New-Item -ItemType Directory -Force -Path $destAplicativo | Out-Null

# [4/5] Espelha p/ a rede. /MIR remove o build antigo (inclusive o .exe da
#       versao anterior, que tem nome diferente) — nunca ficam dois builds.
Write-Host "[4/5] Espelhando para $destAplicativo ..." -ForegroundColor Cyan
robocopy $origem $destAplicativo /MIR /R:3 /W:3 /NFL /NDL /NJH /NP | Out-Null
if ($LASTEXITCODE -ge 8) { throw "robocopy falhou (codigo $LASTEXITCODE)" }

# Verifica que o .exe principal chegou ao destino. NAO renomear:
# Electron 33 quebra ASAR integrity (crash 0x80000003) quando o exe
# principal recebe outro nome.
$exePrincipal = Join-Path $destAplicativo "Gerencie Carteira.exe"
if (-not (Test-Path -LiteralPath $exePrincipal)) {
    throw "Esperado '$exePrincipal' apos o espelhamento, mas nao foi encontrado."
}

# [5/5] Launcher .cmd na pasta-pai. Sobrescrito a cada publish (assim melhorias
#       no launcher se propagam). Em vez de rodar o exe direto do share (~356 MB
#       transmitidos + varredura de antivirus a CADA abertura), o launcher copia
#       o app p/ %LOCALAPPDATA% UMA vez por versao e roda do disco local — as
#       aberturas seguintes ficam instantaneas. O exe mantem o nome neutro
#       (renomear quebra ASAR integrity no Electron 33).
$cmdPath = Join-Path $pastaRede "Gerencie Carteira.cmd"
$cmdConteudo = @'
@echo off
setlocal enableextensions
REM ============================================================
REM Launcher Gerencie Carteira — roda do disco local p/ velocidade.
REM Copia o app do share p/ %LOCALAPPDATA% apenas quando a versao muda;
REM nas aberturas seguintes inicia direto do local (sem stream de rede).
REM ============================================================
set "SHARE=%~dp0Software\Aplicativo"
set "LOCAL=%LOCALAPPDATA%\Gerencie Carteira\Aplicativo"
set "EXE=Gerencie Carteira.exe"

set "SVER="
set "LVER="
if exist "%SHARE%\versao.txt" set /p SVER=<"%SHARE%\versao.txt"
if exist "%LOCAL%\versao.txt" set /p LVER=<"%LOCAL%\versao.txt"

if not exist "%LOCAL%\%EXE%" goto copia
if not "%SVER%"=="%LVER%" goto copia
goto inicia

:copia
echo Preparando aplicativo local (versao %SVER%)... aguarde.
robocopy "%SHARE%" "%LOCAL%" /MIR /R:3 /W:3 /NFL /NDL /NJH /NJS /NP >nul
if errorlevel 8 (
  echo Falha ao copiar do share; abrindo direto da rede...
  start "" "%SHARE%\%EXE%"
  goto fim
)

:inicia
if not exist "%LOCAL%\%EXE%" (
  start "" "%SHARE%\%EXE%"
  goto fim
)
start "" "%LOCAL%\%EXE%"

:fim
endlocal
'@
Set-Content -LiteralPath $cmdPath -Value $cmdConteudo -Encoding ascii
Write-Host "[5/5] Launcher (copia-para-local) gravado: $cmdPath" -ForegroundColor Green

Write-Host ""
Write-Host "OK -> publicado v$versao em: $exePrincipal" -ForegroundColor Green
