Unicode true
!include "MUI2.nsh"
!include "x64.nsh"
!ifndef PAYLOAD
  !error "Falta PAYLOAD (carpeta preparada por construir.py)"
!endif
!ifndef OUTPUT
  !error "Falta OUTPUT (ruta al .exe)"
!endif
Name "Corrector Académico"
OutFile "${OUTPUT}"
InstallDir "$LOCALAPPDATA\CorrectorAcademico"
RequestExecutionLevel user
SetCompressor /SOLID lzma
ShowInstDetails show
VIProductVersion "2.1.0.0"
VIAddVersionKey /LANG=1034 "ProductName" "Corrector Académico"
VIAddVersionKey /LANG=1034 "FileDescription" "Instalador del Corrector Académico y sus dependencias"
VIAddVersionKey /LANG=1034 "FileVersion" "2.1.0"
!define MUI_WELCOMEPAGE_TITLE "Instalación completa del Corrector Académico"
!define MUI_WELCOMEPAGE_TEXT "Este asistente instala el programa y prepara Python, LibreOffice, Graphviz, Java y LanguageTool local.$\r$\n$\r$\nNecesita Internet y puede descargar varios cientos de MB. Windows puede solicitar permisos para instalar dependencias. Al continuar autoriza aceptar sus acuerdos de instalación.$\r$\n$\r$\nLos documentos se procesan localmente. Las entradas existentes y config.json se conservan."
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!define MUI_FINISHPAGE_RUN "$INSTDIR\INICIAR_CORRECCION.bat"
!define MUI_FINISHPAGE_RUN_TEXT "Abrir el Corrector Académico"
!insertmacro MUI_PAGE_FINISH
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_LANGUAGE "Spanish"

Function .onInit
  ${IfNot} ${IsNativeAMD64}
    MessageBox MB_ICONSTOP "Este instalador requiere Windows x64 (procesador Intel/AMD de 64 bits)."
    Abort
  ${EndIf}
FunctionEnd

Section "Instalar y verificar" SEC01
  SetOutPath "$INSTDIR"
  File /r "${PAYLOAD}/*"
  CreateDirectory "$INSTDIR\ENTRADAS\INDICE"
  CreateDirectory "$INSTDIR\ENTRADAS\LECTURAS"
  CreateDirectory "$INSTDIR\SALIDAS\WORD"
  CreateDirectory "$INSTDIR\SALIDAS\INFORMES"
  CreateDirectory "$INSTDIR\SALIDAS\VISTAS_PREVIAS"
  CreateDirectory "$INSTDIR\SALIDAS\COMPRIMIDO"
  CreateDirectory "$INSTDIR\TEMPORALES"
  IfFileExists "$INSTDIR\config.json" conservar_config
  CopyFiles /SILENT "$INSTDIR\config.predeterminada.json" "$INSTDIR\config.json"
  conservar_config:
  DetailPrint "Preparando dependencias y ejecutando pruebas locales..."
  nsExec::ExecToLog '"$SYSDIR\WindowsPowerShell\v1.0\powershell.exe" -NoLogo -NoProfile -ExecutionPolicy RemoteSigned -File "$INSTDIR\instalador\Instalar.ps1" -CarpetaProyecto "$INSTDIR"'
  Pop $0
  StrCmp $0 "0" correcto
    SetErrorLevel 1
    MessageBox MB_ICONSTOP "No se completó la instalación. Consulte $INSTDIR\instalacion.log. No se ha marcado como lista. Puede ejecutar INSTALAR_TODO.bat para reintentar."
    Abort
  correcto:
  CreateDirectory "$SMPROGRAMS\Corrector Académico"
  CreateShortCut "$SMPROGRAMS\Corrector Académico\Iniciar.lnk" "$INSTDIR\INICIAR_CORRECCION.bat" "" "$SYSDIR\shell32.dll" 1
  CreateShortCut "$SMPROGRAMS\Corrector Académico\Carpeta de Lecturas.lnk" "$INSTDIR\ENTRADAS\LECTURAS"
  CreateShortCut "$DESKTOP\Corrector Académico.lnk" "$INSTDIR\INICIAR_CORRECCION.bat"
  WriteUninstaller "$INSTDIR\Desinstalar.exe"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\CorrectorAcademico" "DisplayName" "Corrector Académico"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\CorrectorAcademico" "DisplayVersion" "2.1.0"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\CorrectorAcademico" "InstallLocation" "$INSTDIR"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\CorrectorAcademico" "UninstallString" '$\"$INSTDIR\Desinstalar.exe$\"'
SectionEnd

Section "Uninstall"
  # Conservar ENTRADAS, SALIDAS, TEMPORALES y configuración: contienen trabajo del usuario.
  RMDir /r "$INSTDIR\src"
  RMDir /r "$INSTDIR\tests"
  RMDir /r "$INSTDIR\scripts"
  RMDir /r "$INSTDIR\instalador"
  RMDir /r "$INSTDIR\.venv"
  RMDir /r "$INSTDIR\dependencias"
  Delete "$INSTDIR\INICIAR_CORRECCION.bat"
  Delete "$INSTDIR\INSTALAR_TODO.bat"
  Delete "$INSTDIR\instalacion.json"
  Delete "$INSTDIR\instalacion.pendiente.json"
  Delete "$INSTDIR\instalacion.json.anterior"
  Delete "$INSTDIR\requirements.txt"
  Delete "$INSTDIR\requirements-opcionales.txt"
  Delete "$INSTDIR\config.predeterminada.json"
  Delete "$INSTDIR\Desinstalar.exe"
  Delete "$DESKTOP\Corrector Académico.lnk"
  RMDir /r "$SMPROGRAMS\Corrector Académico"
  DeleteRegKey HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\CorrectorAcademico"
  DetailPrint "Sus documentos y configuración se han conservado en $INSTDIR."
SectionEnd
