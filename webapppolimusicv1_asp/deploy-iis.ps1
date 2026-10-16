<#
================================================================================
 PoliMusic (ASP clasico) - Script de despliegue en IIS
================================================================================
 Copia la aplicacion a la carpeta de IIS, crea el grupo de aplicaciones y el
 sitio, habilita Classic ASP y ajusta los limites de subida de archivos.

 IMPORTANTE: ejecutar en PowerShell ABIERTO COMO ADMINISTRADOR.

 Uso tipico:
   .\deploy-iis.ps1
   .\deploy-iis.ps1 -Target "C:\inetpub\wwwroot" -Port 80
   .\deploy-iis.ps1 -SiteName "PoliMusic" -Port 8080

 Requisitos previos (una sola vez, en "Activar o desactivar caracteristicas"):
   - Internet Information Services
   - World Wide Web Services
       - Application Development Features -> ASP
       - Common HTTP Features -> Default Document
       - Common HTTP Features -> Static Content
   - IIS Management Console

 Nota: Classic ASP NO necesita .NET ni el Hosting Bundle.
================================================================================
#>

[CmdletBinding()]
param(
    [string] $Source   = $PSScriptRoot,
    [string] $Target   = "C:\inetpub\wwwroot\PoliMusic",
    [string] $SiteName = "PoliMusic",
    [string] $AppPool  = "PoliMusicAppPool",
    [int]    $Port     = 8080,
    [int]    $MaxUploadMb = 100,
    [switch] $SkipSite
)

$ErrorActionPreference = "Stop"

function Write-Step($t) { Write-Host "`n==> $t" -ForegroundColor Cyan }
function Write-Ok($t)   { Write-Host "    [OK] $t" -ForegroundColor Green }
function Write-Warn2($t){ Write-Host "    [!]  $t" -ForegroundColor Yellow }
function Write-Err($t)  { Write-Host "    [X]  $t" -ForegroundColor Red }

Write-Host ""
Write-Host "==================================================" -ForegroundColor White
Write-Host " PoliMusic (ASP clasico) - Despliegue en IIS" -ForegroundColor White
Write-Host "==================================================" -ForegroundColor White

# ---------------------------------------------------------------- comprobaciones
$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Err "Este script requiere privilegios de administrador."
    Write-Host "    Abre PowerShell con 'Ejecutar como administrador' y vuelve a intentarlo." -ForegroundColor Yellow
    exit 1
}
Write-Ok "Ejecutando con privilegios de administrador."

if (-not (Test-Path (Join-Path $Source "default.asp"))) {
    Write-Err "No se encontro 'default.asp' en: $Source"
    Write-Host "    Ejecuta el script desde la carpeta del proyecto o pasa -Source." -ForegroundColor Yellow
    exit 1
}
Write-Ok "Origen verificado: $Source"

Import-Module WebAdministration -ErrorAction Stop
Write-Ok "Modulo WebAdministration cargado."

# ------------------------------------------------------------- Classic ASP
Write-Step "Comprobando el soporte de Classic ASP"

$aspDll = "$env:SystemRoot\System32\inetsrv\asp.dll"
if (-not (Test-Path $aspDll)) {
    Write-Err "Classic ASP no esta instalado (falta asp.dll)."
    Write-Host "    Activalo en: Panel de control -> Programas -> Activar o desactivar" -ForegroundColor Yellow
    Write-Host "    caracteristicas de Windows -> Internet Information Services ->" -ForegroundColor Yellow
    Write-Host "    World Wide Web Services -> Application Development Features -> ASP" -ForegroundColor Yellow
    exit 1
}
Write-Ok "asp.dll encontrado."

# ---------------------------------------------------------------- copia
Write-Step "Copiando la aplicacion a $Target"

if (Test-Path $Target) {
    $esDesplieguePrevio = Test-Path (Join-Path $Target "default.asp")
    if ($esDesplieguePrevio) {
        $backup = "$Target._backup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        Write-Warn2 "Ya hay un despliegue anterior en el destino. Se respaldara en:"
        Write-Host "       $backup" -ForegroundColor Yellow
        Move-Item -Path $Target -Destination $backup
        Write-Ok "Respaldo creado."
    } else {
        Write-Warn2 "La carpeta destino ya existe (por ejemplo wwwroot). Se copiara encima sin borrarla."
    }
}

New-Item -ItemType Directory -Path $Target -Force | Out-Null

# Copia selectiva: solo lo que forma parte de la aplicacion
$includeDirs = @("admin", "assets", "includes", "SongFilesAndImages_resources")
$excludeNames = @("docs", ".workbuddy-ai", "deploy-iis.ps1", "README.md", ".git", ".gitignore")

foreach ($dir in $includeDirs) {
    $src = Join-Path $Source $dir
    if (Test-Path $src) {
        Copy-Item -Path $src -Destination (Join-Path $Target $dir) -Recurse -Force
    }
}

Get-ChildItem -Path $Source -Filter "*.asp" -File | ForEach-Object {
    Copy-Item -Path $_.FullName -Destination $Target -Force
}
Get-ChildItem -Path $Source -Filter "*.sql" -File | ForEach-Object {
    Copy-Item -Path $_.FullName -Destination $Target -Force
}

Write-Ok "Archivos copiados: $((Get-ChildItem $Target -Recurse -File).Count) archivos."

# Carpetas que la app usa para subidas
foreach ($sub in @("SongFiles","UserImages","PlaylistImages","AlbumImages","ArtistImages","uploads")) {
    $p = Join-Path $Target "SongFilesAndImages_resources\$sub"
    if (-not (Test-Path $p)) { New-Item -ItemType Directory -Path $p -Force | Out-Null }
}
Write-Ok "Carpetas de recursos verificadas."

# --------------------------------------------------------- grupo de aplicaciones
Write-Step "Configurando el grupo de aplicaciones '$AppPool'"

if (-not (Test-Path "IIS:\AppPools\$AppPool")) {
    New-WebAppPool -Name $AppPool | Out-Null
    Write-Ok "Grupo de aplicaciones creado."
} else {
    Write-Ok "El grupo de aplicaciones ya existia."
}

# Classic ASP no usa codigo administrado
Set-ItemProperty "IIS:\AppPools\$AppPool" -Name managedRuntimeVersion -Value ""
Set-ItemProperty "IIS:\AppPools\$AppPool" -Name managedPipelineMode  -Value "Integrated"
Set-ItemProperty "IIS:\AppPools\$AppPool" -Name startMode           -Value "AlwaysRunning"
Write-Ok "Configurado como 'Sin codigo administrado' (No Managed Code)."

# -------------------------------------------------------------------- permisos
Write-Step "Ajustando permisos NTFS"

foreach ($id in @("IIS_IUSRS", "IIS AppPool\$AppPool")) {
    try {
        $acl = Get-Acl $Target
        $rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            $id, "ReadAndExecute,Write", "ContainerInherit,ObjectInherit", "None", "Allow")
        $acl.SetAccessRule($rule)
        Set-Acl -Path $Target -AclObject $acl
        Write-Ok "Permisos concedidos a: $id"
    } catch {
        Write-Warn2 "No se pudieron aplicar permisos a '$id': $($_.Exception.Message)"
    }
}

# ------------------------------------------------------- configuracion de ASP
Write-Step "Configurando Classic ASP (rutas padre y limites de subida)"

$maxBytes = $MaxUploadMb * 1MB

try {
    # Permite usar "../" en las directivas #include (lo usan las paginas de /admin)
    Set-WebConfigurationProperty -PSPath 'MACHINE/WEBROOT/APPHOST' `
        -Filter "system.webServer/asp" -Name "enableParentPaths" -Value $true
    Write-Ok "Rutas padre habilitadas (enableParentPaths = true)."

    # Limites de ASP: tamano de peticion y buffer
    Set-WebConfigurationProperty -PSPath 'MACHINE/WEBROOT/APPHOST' `
        -Filter "system.webServer/asp/limits" -Name "maxRequestEntityAllowed" -Value $maxBytes
    Set-WebConfigurationProperty -PSPath 'MACHINE/WEBROOT/APPHOST' `
        -Filter "system.webServer/asp/limits" -Name "bufferingLimit" -Value $maxBytes
    Write-Ok "Limite de subida de ASP: $MaxUploadMb MB."

    # Limite del filtrado de peticiones de IIS
    Set-WebConfigurationProperty -PSPath 'MACHINE/WEBROOT/APPHOST' `
        -Filter "system.webServer/security/requestFiltering/requestLimits" `
        -Name "maxAllowedContentLength" -Value $maxBytes
    Write-Ok "Limite de contenido de IIS: $MaxUploadMb MB."

    # Errores detallados de ASP al navegador (util en desarrollo)
    Set-WebConfigurationProperty -PSPath 'MACHINE/WEBROOT/APPHOST' `
        -Filter "system.webServer/asp" -Name "scriptErrorSentToBrowser" -Value $true
    Write-Ok "Errores de ASP visibles en el navegador."

} catch {
    Write-Warn2 "No se pudo aplicar toda la configuracion de ASP: $($_.Exception.Message)"
    Write-Host "    Puedes hacerlo a mano con:" -ForegroundColor Yellow
    Write-Host "      %windir%\system32\inetsrv\appcmd set config /section:asp /enableParentPaths:true" -ForegroundColor Yellow
}

# ------------------------------------------------------------------------ sitio
if (-not $SkipSite) {
    Write-Step "Creando/actualizando el sitio IIS '$SiteName' en el puerto $Port"

    if (Get-Website -Name $SiteName -ErrorAction SilentlyContinue) {
        Write-Warn2 "El sitio '$SiteName' ya existe; se actualizara su configuracion."
        Set-ItemProperty "IIS:\Sites\$SiteName" -Name physicalPath -Value $Target
        Set-ItemProperty "IIS:\Sites\$SiteName" -Name applicationPool -Value $AppPool
    } else {
        New-Website -Name $SiteName -PhysicalPath $Target -ApplicationPool $AppPool -Port $Port -Force | Out-Null
        Write-Ok "Sitio creado."
    }

    Start-WebAppPool -Name $AppPool -ErrorAction SilentlyContinue
    if ((Get-Website -Name $SiteName).State -ne "Started") { Start-Website -Name $SiteName }
    Write-Ok "Sitio iniciado."
}

# --------------------------------------------------------------- documento inicial
Write-Step "Verificando el documento predeterminado"

$defaults = (Get-WebConfigurationProperty -PSPath "IIS:\Sites\$SiteName" -Filter "system.webServer/defaultDocument/files" -Name "." -ErrorAction SilentlyContinue) |
            ForEach-Object { $_.value }
if ($defaults -notcontains "default.asp") {
    Add-WebConfigurationProperty -PSPath "IIS:\Sites\$SiteName" `
        -Filter "system.webServer/defaultDocument/files" -Name "." -Value @{value="default.asp"} -ErrorAction SilentlyContinue
    Write-Ok "Se agrego 'default.asp' a los documentos predeterminados."
} else {
    Write-Ok "'default.asp' ya esta en los documentos predeterminados."
}

# ------------------------------------------------------------------- resumen
Write-Host ""
Write-Host "==================================================" -ForegroundColor Green
Write-Host " Despliegue completado" -ForegroundColor Green
Write-Host "==================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Carpeta    : $Target"
Write-Host "  App Pool   : $AppPool  (No Managed Code)"
Write-Host "  Sitio      : $SiteName"
if (-not $SkipSite) { Write-Host "  URL        : http://localhost:$Port/" }
Write-Host ""
Write-Host "  Credenciales de prueba:" -ForegroundColor White
Write-Host "    Administrador : userAdmin  / 12345678"
Write-Host "    Usuario       : userNormal / 12345678"
Write-Host ""
Write-Host "  Base de datos: BDD_PoliMusic (localhost)" -ForegroundColor White
Write-Host "    Si la base no existe, ejecuta BDD_PoliMusic.sql en SQL Server."
Write-Host "    Al abrir la app se agregan solos los campos de genero que faltan."
Write-Host ""
Write-Host "  Si aparece un error 500:" -ForegroundColor White
Write-Host "    - Revisa que Classic ASP este activado (asp.dll)"
Write-Host "    - Revisa que enableParentPaths este en true"
Write-Host "    - Visor de eventos -> Registros de Windows -> Aplicacion"
Write-Host ""
