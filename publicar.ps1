# ==============================================================================
# publicar.ps1 — Builda o app e publica na pasta de rede.
#
# Fluxo: build local (build-app.ps1 -> .\Aplicativo) -> espelha p/
#        A:\...\Software\Aplicativo -> renomeia o .exe com a versao ->
#        garante o atalho .cmd na pasta-pai.
#
# - O /MIR do robocopy remove o build anterior automaticamente: nunca ficam dois.
# - A versao no nome do .exe e DERIVADA de app/package.json (fonte unica) —
#   nao e hardcoded em lugar nenhum.
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

# Versao no nome do .exe distribuido (derivada da fonte unica).
$exeNeutro = Join-Path $destAplicativo "Gerencie Carteira.exe"
$exeVersao = "Gerencie Carteira $versao.exe"
if (-not (Test-Path -LiteralPath $exeNeutro)) {
    throw "Esperado '$exeNeutro' apos o espelhamento, mas nao foi encontrado."
}
try {
    Rename-Item -LiteralPath $exeNeutro -NewName $exeVersao -Force
} catch {
    throw "Nao foi possivel renomear o .exe na rede para '$exeVersao'. " +
          "Alguem pode estar com o app aberto. Detalhe: $_"
}

# [5/5] Atalho .cmd na pasta-pai. Escrito uma unica vez: usa curinga para achar
#       o .exe atual, entao nao precisa ser regerado a cada versao.
$cmdPath = Join-Path $pastaRede "Gerencie Carteira.cmd"
if (-not (Test-Path -LiteralPath $cmdPath)) {
    $cmdConteudo = @'
@echo off
REM Atalho de um clique. Acha o .exe atual em Software\Aplicativo\
REM (o nome carrega a versao; o curinga dispensa atualizar este arquivo).
for %%f in ("%~dp0Software\Aplicativo\Gerencie Carteira*.exe") do (
  start "" "%%~ff"
  goto :eof
)
echo Nenhum executavel encontrado em Software\Aplicativo\
pause
'@
    Set-Content -LiteralPath $cmdPath -Value $cmdConteudo -Encoding ascii
    Write-Host "[5/5] Atalho criado: $cmdPath" -ForegroundColor Green
} else {
    Write-Host "[5/5] Atalho ja existe (mantido): $cmdPath" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "OK -> publicado: $destAplicativo\$exeVersao" -ForegroundColor Green
