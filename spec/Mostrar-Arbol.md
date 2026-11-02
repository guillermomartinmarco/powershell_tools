# Función Powershell Mostrar-Arbol

Muestra el arbol de directorios a fondo. Hacer esta función en Powershell
Puede tener filtro por tipo de archivos

```powershell
function Mostrar-Arbol {
    param (
        [Parameter(Position=0)]
        [string]$Ruta = ".",
        
        [Parameter(Position=1)]
        [string]$Filtro = "*",
        
        [string]$Prefijo = ""
    )

    # ---------------------------------------------------------
    # CATÁLOGO DE EMOJIS (Fácil de editar y expandir)
    # ---------------------------------------------------------
    $Iconos = @{
        '.zip'  = '📦 '
        '.pdf'  = '📑 '
        '.html' = '🌐 '
        '.js'   = '🌐 '
        '.css'  = '🌐 '
        '.md'   = '📝 '
        '.png'  = '🖼️ '
        '.jpg'  = '🖼️ '
        '.jpeg' = '🖼️ '
        '.bmp'  = '🖼️ '
        '.gif'  = '🖼️ '
        '.tiff' = '🖼️ '
        '.dwg'  = '📐 '
        '.dxf'  = '📐 '
        '.stp'  = '🧊 '
    }

    # Truco: Si pasas un comodín como primer parámetro (ej: Mostrar-Arbol *.zip)
    if ($Ruta -match '[\*\?]') {
        $Filtro = $Ruta
        $Ruta = "."
    }

    # Obtener los elementos de la ruta actual
    $TodosLosElementos = Get-ChildItem -LiteralPath $Ruta -ErrorAction SilentlyContinue | Sort-Object Name
    
    # Filtrar los elementos
    $Elementos = @()
    foreach ($e in $TodosLosElementos) {
        if (-not $e.PSIsContainer) {
            if ($e.Name -like $Filtro) { $Elementos += $e }
        } else {
            if ($Filtro -eq "*") {
                $Elementos += $e
            } else {
                $TieneArchivos = Get-ChildItem -Path $e.FullName -Filter $Filtro -Recurse -File -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($TieneArchivos) { $Elementos += $e }
            }
        }
    }

    $Total = $Elementos.Count
    for ($i = 0; $i -lt $Total; $i++) {
        $Elemento = $Elementos[$i]
        $EsElUltimo = ($i -eq $Total - 1)
        
        # Elegir los caracteres de la rama
        if ($EsElUltimo) { $SimboloRama = "└── " } else { $SimboloRama = "├── " }
        
        # ---------------------------------------------------------
        # ASIGNAR EL EMOJI ADECUADO
        # ---------------------------------------------------------
        if ($Elemento.PSIsContainer) {
            $Emoji = "📁 "
        } else {
            # Convertimos la extensión a minúsculas por si hay archivos .ZIP o .Zip
            $Ext = $Elemento.Extension.ToLower() 
            
            if ($Iconos.ContainsKey($Ext)) {
                $Emoji = $Iconos[$Ext]
            } else {
                $Emoji = "📄 " # Predeterminado si no está en el catálogo
            }
        }
        
        # Mostrar el elemento actual con su jerarquía exacta y el emoji
        Write-Output "${Prefijo}${SimboloRama}${Emoji}$($Elemento.Name)"
        
        # Si es una carpeta, entrar de forma recursiva
        if ($Elemento.PSIsContainer) {
            if ($EsElUltimo) { $Espacio = "    " } else { $Espacio = "│   " }
            $SiguientePrefijo = $Prefijo + $Espacio
            
            # Llamada recursiva manteniendo el filtro
            Mostrar-Arbol -Ruta $Elemento.FullName -Filtro $Filtro -Prefijo $SiguientePrefijo
        }
    }
}
```

Mostrar-Arbol
Mostrar-Arbol *3D*
