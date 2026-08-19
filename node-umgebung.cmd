@echo off
rem ---------------------------------------------------------------------------
rem  Oeffnet eine Eingabeaufforderung, in der die mitgelieferte Node.js-Laufzeit
rem  im Suchpfad liegt. Danach laesst sich "node" direkt eintippen, ohne
rem  Node.js zu installieren.
rem
rem  Beenden mit "exit".
rem ---------------------------------------------------------------------------
setlocal

set "ROOT=%~dp0"
set "NODE_DIR=%ROOT%node-v24.19.0-win-x64"

if not exist "%NODE_DIR%\node.exe" (
    echo Fehler: "%NODE_DIR%\node.exe" wurde nicht gefunden.
    echo Bitte das Repository vollstaendig herunterladen und entpacken.
    pause
    exit /b 1
)

set "PATH=%NODE_DIR%;%PATH%"
cd /d "%ROOT%"

echo.
echo  Node.js-Umgebung bereit:
for /f "delims=" %%v in ('node -v') do echo    node %%v
echo.
echo  Beispiele:
echo    node dist\maskierer.cjs "C:\Daten\kunden.xlsx" --list
echo    node dist\maskierer.cjs "C:\Daten\kunden.xlsx" --keep "KundenID"
echo.
echo  Hinweis: npm ist nicht enthalten - fuer die Weiterentwicklung ein
echo  vollstaendiges Node.js installieren.
echo.

cmd /k
