#requires -Version 5.1
[CmdletBinding()]
param([string]$CarpetaProyecto = (Split-Path -Parent $PSScriptRoot), [switch]$SoloCargarFunciones)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Actualizar-Ruta {
    $machine = [Environment]::GetEnvironmentVariable('Path', 'Machine')
    $user = [Environment]::GetEnvironmentVariable('Path', 'User')
    $env:Path = "$machine;$user;$env:Path"
}

function Ejecutar-Comando {
    param([string]$Programa, [string[]]$Argumentos)
    $previous = $ErrorActionPreference
    try {
        # Java escribe su versión en stderr incluso cuando el resultado es correcto.
        $ErrorActionPreference = 'Continue'
        & $Programa @Argumentos | Out-Host
        $code = $LASTEXITCODE
    } finally { $ErrorActionPreference = $previous }
    if ($code -ne 0) { throw "Falló $Programa (código $code). Consulte instalacion.log." }
}

function Obtener-WinGet {
    Actualizar-Ruta
    $command = Get-Command winget.exe -ErrorAction SilentlyContinue
    if ($command) { return $command.Source }
    Write-Host 'Preparando el administrador oficial de paquetes de Microsoft...'
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force | Out-Null
    Install-Module -Name Microsoft.WinGet.Client -Repository PSGallery -Scope CurrentUser -Force -AllowClobber
    Import-Module Microsoft.WinGet.Client
    Repair-WinGetPackageManager -Force -Latest | Out-Host
    Actualizar-Ruta
    $command = Get-Command winget.exe -ErrorAction SilentlyContinue
    if (-not $command) { throw 'WinGet no está disponible después de repararlo. Consulte instalacion.log.' }
    return $command.Source
}

function Instalar-Paquete {
    param([string]$Id, [string]$WinGet)
    Write-Host "Instalando $Id desde WinGet..."
    Ejecutar-Comando -Programa $WinGet -Argumentos @('install', '--id', $Id, '--exact', '--source', 'winget',
        '--silent', '--accept-package-agreements', '--accept-source-agreements', '--disable-interactivity')
    # WinGet mantiene la verificación de firmas/checksums y gestiona dependencias.
    Actualizar-Ruta
}

function Buscar-Python {
    $paths = @("$env:LOCALAPPDATA\Programs\Python\Python312\python.exe", "$env:ProgramFiles\Python312\python.exe")
    foreach ($key in @('HKCU:\Software\Python\PythonCore\3.12\InstallPath', 'HKLM:\Software\Python\PythonCore\3.12\InstallPath')) {
        if (Test-Path $key) { $paths += Join-Path (Get-Item $key).GetValue('') 'python.exe' }
    }
    foreach ($path in $paths) {
        if (Test-Path -LiteralPath $path) {
            & $path -c 'import sys; sys.exit(0 if sys.version_info[:2] == (3,12) and sys.maxsize > 2**32 else 1)' 2>$null
            if ($LASTEXITCODE -eq 0) { return $path }
        }
    }
    return $null
}

function Buscar-Herramienta {
    param([ValidateSet('dot','soffice','java')][string]$Nombre)
    $command = Get-Command "$Nombre.exe" -ErrorAction SilentlyContinue
    $candidates = @()
    if ($command) { $candidates += $command.Source }
    $folders = switch ($Nombre) {
        'soffice' { @("$env:ProgramFiles\LibreOffice", "${env:ProgramFiles(x86)}\LibreOffice") }
        'dot' { @(Get-ChildItem -Path "$env:ProgramFiles\Graphviz*" -Directory -ErrorAction SilentlyContinue | ForEach-Object FullName) }
        'java' { @("$env:ProgramFiles\Eclipse Adoptium", "$env:ProgramFiles\Java") }
    }
    foreach ($folder in $folders) {
        if ($folder -and (Test-Path -LiteralPath $folder)) {
            $candidates += @(Get-ChildItem -LiteralPath $folder -Recurse -Filter "$Nombre.exe" -File -ErrorAction SilentlyContinue | ForEach-Object FullName)
        }
    }
    foreach ($candidate in ($candidates | Select-Object -Unique)) {
        if ($Nombre -eq 'java') {
            $previous = $ErrorActionPreference
            try {
                $ErrorActionPreference = 'Continue'
                $version = (& $candidate -version 2>&1 | Out-String)
                $code = $LASTEXITCODE
            } finally { $ErrorActionPreference = $previous }
            if ($code -ne 0 -or $version -notmatch 'version\s+"(\d+)' -or [int]$Matches[1] -lt 17) { continue }
        }
        return $candidate
    }
    return $null
}

function Descargar-Verificado {
    param([string]$Url, [string]$Archivo, [string]$Sha256)
    if (-not $Url.StartsWith('https://') -or $Sha256 -notmatch '^[a-fA-F0-9]{64}$') { throw 'La descarga necesita HTTPS y un checksum SHA256 válido.' }
    if (Test-Path -LiteralPath $Archivo) {
        if ((Get-FileHash -LiteralPath $Archivo -Algorithm SHA256).Hash -eq $Sha256) { return }
        throw "El archivo en caché no supera el checksum: $Archivo. No se usará."
    }
    $partial = "$Archivo.partial"
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $Url -OutFile $partial -UseBasicParsing
        if ((Get-FileHash -LiteralPath $partial -Algorithm SHA256).Hash -ne $Sha256) { throw 'El checksum de LanguageTool no coincide. Se detiene la instalación.' }
        Move-Item -LiteralPath $partial -Destination $Archivo
    } finally {
        if (Test-Path -LiteralPath $partial) { Remove-Item -LiteralPath $partial -Force }
    }
}

function Instalar-Corrector {
    param([string]$Raiz)
    if ([Environment]::OSVersion.Platform -ne 'Win32NT' -or -not [Environment]::Is64BitOperatingSystem) { throw 'Este instalador requiere Windows 10/11 de 64 bits.' }
    if ([Environment]::OSVersion.Version.Build -lt 17763) { throw 'Se necesita Windows 10 1809 o posterior.' }
    $Raiz = (Resolve-Path -LiteralPath $Raiz).Path
    Set-Location -LiteralPath $Raiz
    $manifest = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'dependencias.json') -Raw | ConvertFrom-Json
    $deps = Join-Path $Raiz 'dependencias'
    New-Item -ItemType Directory -Path $deps -Force | Out-Null
    $marker = Join-Path $Raiz 'instalacion.json'
    if (Test-Path $marker) { Move-Item $marker "$marker.anterior" -Force }
    $winget = Obtener-WinGet
    $python = Buscar-Python
    if (-not $python) { Instalar-Paquete $manifest.winget.python $winget; $python = Buscar-Python }
    if (-not $python) { throw 'No se encontró Python 3.12 x64 después de instalarlo.' }
    $tools = @{}
    foreach ($pair in @(@('soffice','libreoffice'), @('dot','graphviz'), @('java','java'))) {
        $path = Buscar-Herramienta $pair[0]
        if (-not $path) { Instalar-Paquete $manifest.winget.($pair[1]) $winget; $path = Buscar-Herramienta $pair[0] }
        if (-not $path) { throw "No se encontró $($pair[0]) después de instalarlo." }
        $tools[$pair[0]] = $path
    }
    Ejecutar-Comando -Programa $tools.java -Argumentos @('-version')
    Ejecutar-Comando -Programa $tools.dot -Argumentos @('-V')
    Ejecutar-Comando -Programa $tools.soffice -Argumentos @('--version')
    Write-Host 'Preparando el entorno virtual y las bibliotecas...'
    Ejecutar-Comando -Programa $python -Argumentos @('-m', 'venv', '.venv')
    $venv = Join-Path $Raiz '.venv\Scripts\python.exe'
    Ejecutar-Comando -Programa $venv -Argumentos @('-m', 'pip', 'install', '--disable-pip-version-check', '-r', 'requirements.txt', '-r', 'requirements-opcionales.txt')
    Ejecutar-Comando -Programa $venv -Argumentos @('-m', 'pip', 'check')
    $archive = Join-Path $deps "LanguageTool-$($manifest.languagetool.version).zip"
    Write-Host 'Descargando y verificando el motor lingüístico local...'
    Descargar-Verificado $manifest.languagetool.url $archive $manifest.languagetool.sha256
    $lt = Join-Path $deps $manifest.languagetool.carpeta
    $tempExtract = Join-Path $deps ('extraer_' + [Guid]::NewGuid().ToString('N'))
    try {
        Expand-Archive -LiteralPath $archive -DestinationPath $tempExtract
        $extracted = Join-Path $tempExtract $manifest.languagetool.carpeta
        if (-not (Test-Path (Join-Path $extracted 'languagetool-server.jar'))) { throw 'Distribución LanguageTool incompleta.' }
        if (Test-Path $lt) { Remove-Item -LiteralPath $lt -Recurse -Force }
        Move-Item -LiteralPath $extracted -Destination $lt
    } finally { if (Test-Path $tempExtract) { Remove-Item -LiteralPath $tempExtract -Recurse -Force } }
    $installation = @{ version = 2; python = $venv; java = $tools.java; dot = $tools.dot;
        soffice = $tools.soffice; languagetool = $lt; languagetool_sha256 = $manifest.languagetool.sha256 }
    $pending = Join-Path $Raiz 'instalacion.pendiente.json'
    $installation | ConvertTo-Json | Set-Content -LiteralPath $pending -Encoding UTF8
    Write-Host 'Probando Word, PDF, diagramas y LanguageTool español/inglés...'
    Ejecutar-Comando -Programa $venv -Argumentos @('-m', 'src.arranque', '--verificar-instalacion', '--instalacion', $pending)
    Move-Item -LiteralPath $pending -Destination $marker -Force
    Write-Host 'Instalación completada y comprobada. Abra INICIAR_CORRECCION.bat.'
}

if (-not $SoloCargarFunciones) {
    $log = Join-Path $CarpetaProyecto 'instalacion.log'; $started = $false
    try {
        Start-Transcript -LiteralPath $log -Append | Out-Null; $started = $true
        Instalar-Corrector $CarpetaProyecto
    } catch {
        Write-Host "ERROR DE INSTALACIÓN: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Registro: $log"
        exit 1
    } finally { if ($started) { Stop-Transcript | Out-Null } }
}
