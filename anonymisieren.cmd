@echo off
rem ---------------------------------------------------------------------------
rem  Excel-Anonymisierer - Start ohne Installation.
rem
rem  Die Node.js-Laufzeit liegt entpackt im Projektordner - es ist nichts
rem  einzurichten. Gesucht wird in dieser Reihenfolge:
rem    1. die mitgelieferte Laufzeit in node-v24.19.0-win-x64\
rem    2. ein installiertes Node.js aus dem Suchpfad
rem
rem  Aufruf:  anonymisieren.cmd "C:\Daten\kunden.xlsx" --list
rem ---------------------------------------------------------------------------
setlocal

set "ROOT=%~dp0"
set "NODE_EXE=%ROOT%node-v24.19.0-win-x64\node.exe"
set "BUNDLE=%ROOT%dist\anonymize-xlsx.cjs"

if not exist "%BUNDLE%" (
    echo Fehler: "%BUNDLE%" wurde nicht gefunden.
    echo Bitte das Repository vollstaendig herunterladen und entpacken.
    exit /b 1
)

if exist "%NODE_EXE%" goto :run

where node >nul 2>nul
if errorlevel 1 (
    echo Fehler: Keine Node.js-Laufzeit gefunden.
    echo Erwartet wurde "%NODE_EXE%".
    echo Bitte das Repository vollstaendig herunterladen und entpacken,
    echo oder Node.js ab Version 20 installieren.
    exit /b 1
)

node "%BUNDLE%" %*
exit /b %errorlevel%

:run
"%NODE_EXE%" "%BUNDLE%" %*
exit /b %errorlevel%
