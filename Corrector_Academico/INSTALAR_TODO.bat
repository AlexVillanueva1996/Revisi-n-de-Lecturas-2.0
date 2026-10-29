@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"
echo INSTALACION COMPLETA DEL CORRECTOR ACADEMICO
echo Necesita Internet. Windows puede pedir permisos para las dependencias.
powershell.exe -NoLogo -NoProfile -ExecutionPolicy RemoteSigned -File "%~dp0instalador\Instalar.ps1" -CarpetaProyecto "%~dp0."
set "RESULTADO=%ERRORLEVEL%"
if not "%RESULTADO%"=="0" echo La instalacion fallo. Consulte instalacion.log; no se declaro lista.
pause
exit /b %RESULTADO%
