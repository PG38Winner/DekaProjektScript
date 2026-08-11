@echo off
rem ---------------------------------------------------------------------------
rem  Excel-Anonymisierer - Start ohne Installation.
rem
rem  Sucht eine Node.js-Laufzeit in dieser Reihenfolge:
rem    1. bereits entpackte Laufzeit im Projektordner
rem    2. entpackt die mitgelieferte ZIP-Datei, falls noetig
rem    3. ein installiertes Node.js aus dem Suchpfad
rem
rem  Aufruf:  anonymisieren.cmd "C:\Daten\kunden.xlsx" --list
rem ---------------------------------------------------------------------------
setlocal

set "ROOT=%~dp0"
set "NODE_DIR=%ROOT%node-v24.19.0-win-x64"
set "NODE_EXE=%NODE_DIR%\node.exe"
set "NODE_ZIP=%ROOT%node-v24.19.0-win-x64.zip"
set "BUNDLE=%ROOT%dist\anonymize-xlsx.cjs"

if not exist "%BUNDLE%" (
    echo Fehler: "%BUNDLE%" wurde nicht gefunden.
    echo Bitte das Repository vollstaendig auschecken.
    exit /b 1
)

if exist "%NODE_EXE%" goto :run

if exist "%NODE_ZIP%" (
    echo Entpacke die mitgelieferte Node.js-Laufzeit ...
    rem -Command wertet keine Skriptdatei aus, deshalb ist keine gelockerte
    rem Ausfuehrungsrichtlinie noetig.
    powershell -NoProfile -Command "Expand-Archive -LiteralPath '%NODE_ZIP%' -DestinationPath '%ROOT%.' -Force"
    if exist "%NODE_EXE%" goto :run
)

where node >nul 2>nul
if errorlevel 1 (
    echo Fehler: Keine Node.js-Laufzeit gefunden.
    echo Entweder "%NODE_ZIP%" bereitstellen oder Node.js ab Version 20 installieren.
    exit /b 1
)

node "%BUNDLE%" %*
exit /b %errorlevel%

:run
"%NODE_EXE%" "%BUNDLE%" %*
exit /b %errorlevel%
