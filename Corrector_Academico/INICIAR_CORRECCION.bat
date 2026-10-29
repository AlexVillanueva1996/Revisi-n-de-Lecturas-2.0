@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"
echo ==========================================
echo  SISTEMA DE CORRECCION ACADEMICA
echo ==========================================
if not exist "instalacion.json" goto instalar
if not exist ".venv\Scripts\python.exe" goto instalar
goto ejecutar
:instalar
echo Preparando todas las dependencias. Necesita Internet durante la instalacion.
powershell.exe -NoLogo -NoProfile -ExecutionPolicy RemoteSigned -File "%~dp0instalador\Instalar.ps1" -CarpetaProyecto "%~dp0."
if errorlevel 1 goto error
:ejecutar
".venv\Scripts\python.exe" -m src.arranque
set "RESULTADO=%ERRORLEVEL%"
if "%RESULTADO%"=="1" goto error
if "%RESULTADO%"=="2" echo Hay documentos pendientes de revision. Consulte los informes.
echo Resultados en "%~dp0SALIDAS"
choice /C SN /N /M "Abrir la carpeta de resultados? [S/N]: "
if errorlevel 2 goto fin
start "" "%~dp0SALIDAS"
:fin
pause
exit /b %RESULTADO%
:error
echo ERROR: Revise el mensaje anterior y SALIDAS\INFORMES. No se modificaron los originales.
pause
exit /b 1
