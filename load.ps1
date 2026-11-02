$DirectorioLib = Join-Path $PSScriptRoot "lib"

# ---------------------------------------------------------
# DICCIONARIO DE ALIAS (Tus atajos de teclado)
# ---------------------------------------------------------
$DiccionarioAlias = @{
    'ma'    = 'Mostrar-Arbol'
    'arbol' = 'Mostrar-Arbol'
    'doc2docx' = 'Convert-DocToDocx'
    # 'lt'  = 'Limpiar-Temporales' (Ejemplo futuro)
}

# Inicializar variables globales de rastreo
if ($null -eq $global:MisFuncionesCargadas) { $global:MisFuncionesCargadas = @() }
if ($null -eq $global:MisAliasCargados)     { $global:MisAliasCargados = @() }

# --- CARGAR FUNCIONES ---
if (Test-Path $DirectorioLib) {
    $Archivos = Get-ChildItem -Path $DirectorioLib -Filter "*.ps1"
    $FuncionesDetectadas = @()

    foreach ($Archivo in $Archivos) {
        $Lineas = Get-Content $Archivo.FullName
        foreach ($Linea in $Lineas) {
            if ($Linea -match '(?i)^\s*function\s+([a-zA-Z0-9_-]+)') {
                $FuncionesDetectadas += $matches[1]
            }
        }
        . $Archivo.FullName
    }

    $global:MisFuncionesCargadas += $FuncionesDetectadas
    $global:MisFuncionesCargadas = $global:MisFuncionesCargadas | Select-Object -Unique
}

# --- CARGAR ALIAS ---
$AliasRegistrados = 0
foreach ($Atajo in $DiccionarioAlias.Keys) {
    $NombreFuncion = $DiccionarioAlias[$Atajo]
    
    # Solo crea el alias si la función existe en la memoria
    if (Get-Command $NombreFuncion -ErrorAction SilentlyContinue) {
        Set-Alias -Name $Atajo -Value $NombreFuncion -Scope Global -Force
        $global:MisAliasCargados += $Atajo
        $AliasRegistrados++
    }
}
$global:MisAliasCargados = $global:MisAliasCargados | Select-Object -Unique

Write-Host "✅ Se cargaron $($FuncionesDetectadas.Count) funciones y $AliasRegistrados alias." -ForegroundColor Green
