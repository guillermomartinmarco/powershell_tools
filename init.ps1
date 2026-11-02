# Obtiene la ruta de la carpeta donde está este script
$DirectorioActual = $PSScriptRoot
$ComandoInyeccion = ". `"$DirectorioActual\init.ps1`""

# 1. Crear el perfil si no existe
if (!(Test-Path -Path $PROFILE)) {
    Write-Host "Creando archivo de perfil en $PROFILE..." -ForegroundColor Cyan
    New-Item -ItemType File -Path $PROFILE -Force | Out-Null
}

# 2. Inyectarse en el perfil si no está presente
$ContenidoPerfil = Get-Content $PROFILE -Raw -ErrorAction SilentlyContinue
if ($ContenidoPerfil -notmatch [regex]::Escape($ComandoInyeccion)) {
    Write-Host "Inyectando init.ps1 en tu `$PROFILE..." -ForegroundColor Yellow
    Add-Content -Path $PROFILE -Value "`n# Cargar entorno personalizado de PowerShell"
    Add-Content -Path $PROFILE -Value $ComandoInyeccion
}

# 3. Cargar las funciones
$RutaLoad = Join-Path $DirectorioActual "load.ps1"
if (Test-Path $RutaLoad) {
    . $RutaLoad
} else {
    Write-Warning "No se encontró load.ps1 en $DirectorioActual"
}
