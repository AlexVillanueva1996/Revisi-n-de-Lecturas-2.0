param([string]$Raiz = (Split-Path -Parent $PSScriptRoot))
$ErrorActionPreference = 'Stop'
$script = Join-Path $Raiz 'instalador/Instalar.ps1'
$tokens=$null; $errors=$null
[System.Management.Automation.Language.Parser]::ParseFile($script,[ref]$tokens,[ref]$errors) | Out-Null
if ($errors.Count) { throw ($errors | Out-String) }
. $script -SoloCargarFunciones
$manifest = Get-Content (Join-Path $Raiz 'instalador/dependencias.json') -Raw | ConvertFrom-Json
if ($manifest.languagetool.sha256 -notmatch '^[a-f0-9]{64}$') { throw 'Hash inválido.' }
if ($manifest.winget.python -ne 'Python.Python.3.12') { throw 'Python no está fijado a la versión elegida.' }
$temp = Join-Path ([IO.Path]::GetTempPath()) ('test_instalador_'+[Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory $temp | Out-Null
try {
    $file = Join-Path $temp 'archivo verificado.zip'
    [IO.File]::WriteAllText($file,'prueba de checksum')
    $hash = (Get-FileHash $file -Algorithm SHA256).Hash
    Descargar-Verificado 'https://example.org/no-se-consulta' $file $hash
    $failed=$false
    try { Descargar-Verificado 'https://example.org/no-se-consulta' $file ('0'*64) } catch { $failed=$true }
    if (-not $failed) { throw 'No se rechazó un checksum incorrecto.' }
    $failed=$false
    try { Descargar-Verificado 'http://example.org' $file $hash } catch { $failed=$true }
    if (-not $failed) { throw 'No se rechazó una URL sin TLS.' }
    $executable = (Get-Process -Id $PID).Path
    Ejecutar-Comando -Programa $executable -Argumentos @('-NoProfile','-Command','exit 0')
    $failed=$false
    try { Ejecutar-Comando -Programa $executable -Argumentos @('-NoProfile','-Command','exit 7') } catch { $failed=$true }
    if (-not $failed) { throw 'Se ocultó un fallo de proceso.' }
    Write-Host 'OK: parser PowerShell, manifiesto, caché verificada, rechazo de hash/TLS y propagación de errores.'
} finally { Remove-Item -LiteralPath $temp -Recurse -Force }
