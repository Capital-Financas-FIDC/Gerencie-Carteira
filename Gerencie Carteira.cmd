@echo off
REM Atalho de um clique para o app empacotado.
REM O executavel real fica em .\Aplicativo\ (regenerado por build-app.ps1).
start "" "%~dp0Aplicativo\Gerencie Carteira.exe"
