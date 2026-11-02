function Mostrar-Arbol {
    [CmdletBinding()]
    param (
        [Parameter(Position=0)]
        [string]$Ruta = ".",
        
        [Parameter(Position=1)]
        [string]$Filtro = "*",
        
        [Parameter(ValueFromPipeline)]
        [Alias('e')]
        [string[]]$Excluir = @(),
        
        [Alias('g')]
        [switch]$Gitignore,

        [Alias('h')]
        [switch]$Help,

        [string]$Prefijo = ""
    )

    # ---------------------------------------------------------
    # MENÚ DE AYUDA (--help o -h)
    # ---------------------------------------------------------
    if ($Help) {
        Write-Host "📦 COMANDO: Mostrar-Arbol" -ForegroundColor Cyan
        Write-Host "Muestra la estructura de directorios en forma de árbol con iconos y emojis.`n"
        Write-Host "SINTAXIS:" -ForegroundColor Yellow
        Write-Host "  Mostrar-Arbol [[-Ruta] <String>] [[-Filtro] <String>] [-Excluir <String[]>] [-Gitignore] [-Help]`n"
        Write-Host "PARÁMETROS:" -ForegroundColor Yellow
        Write-Host "  -Ruta        Carpeta base para iniciar el árbol. Por defecto es '.' (directorio actual)."
        Write-Host "  -Filtro      Muestra solo archivos que coincidan (ej: '*.py')."
        Write-Host "  -Excluir, -e Lista de carpetas o archivos exactos a omitir (ej: -e '.venv','dist')."
        Write-Host "  -Gitignore, -g Busca un archivo .gitignore en la ruta y procesa sus reglas de exclusión."
        Write-Host "  -Help, -h    Muestra esta pantalla de ayuda.`n"
        Write-Host "EJEMPLOS:" -ForegroundColor Yellow
        Write-Host "  Mostrar-Arbol -e .venv, .vscode" -ForegroundColor Gray
        Write-Host "  Mostrar-Arbol -g" -ForegroundColor Gray
        return
    }

    # ---------------------------------------------------------
    # PROCESAR REGLAS DE .GITIGNORE Y EXCLUSIONES
    # ---------------------------------------------------------
    $ReglasExcluir = [System.Collections.Generic.List[string]]::new()
    if ($Excluir) { $ReglasExcluir.AddRange($Excluir) }

    if ($Gitignore) {
        $RutaGitignore = Join-Path (Resolve-Path $Ruta).Path ".gitignore"
        if (Test-Path $RutaGitignore) {
            $LineasGit = Get-Content $RutaGitignore | ForEach-Object { $_.Trim() } | Where-Object { $_ -and -not $_.StartsWith('#') }
            foreach ($linea in $LineasGit) {
                $Limpia = $linea -replace '/$', ''      
                $Limpia = $Limpia -replace '^\+/', ''   
                if ($Limpia) { $ReglasExcluir.Add($Limpia) }
            }
        }
    }
    $ListaFinalExcluir = $ReglasExcluir | Select-Object -Unique

    # ---------------------------------------------------------
    # CATÁLOGO DE EMOJIS
    # ---------------------------------------------------------
    $Iconos = @{
        '.7z'   = '📦 '; '.ai'   = '🎨 '; '.avi'  = '🎬 '; '.bat'  = '📜 '
        '.bmp'  = '🖼️ '; '.c'    = '📜 '; '.cfg'  = '⚙️ '; '.cpp'  = '📜 '
        '.csv'  = '📊 '; '.css'  = '🌐 '; '.db'   = '🗄️ '; '.dll'  = '⚙️ '
        '.doc'  = '📘 '; '.docx' = '📘 '; '.dwg'  = '📐 '; '.dxf'  = '📐 '
        '.exe'  = '⚙️ '; '.flac' = '🎵 '; '.gif'  = '🖼️ '; '.gz'   = '📦 '
        '.html' = '🌐 '; '.ico'  = '🖼️ '; '.ini'  = '⚙️ '; '.java' = '📜 '
        '.jpeg' = '🖼️ '; '.jpg'  = '🖼️ '; '.js'   = '🌐 '; '.json' = '⚙️ '
        '.jt'   = '🧊 '; '.key'  = '🔑 '; '.m'    = '📜 '; '.md'   = '📝 '
        '.mkv'  = '🎬 '; '.mov'  = '🎬 '; '.mp3'  = '🎵 '; '.mp4'  = '🎬 '
        '.obj'  = '🧊 '; '.pdf'  = '📖 '; '.php'  = '🌐 '; '.png'  = '🖼️ '
        '.ppt'  = '📽️ '; '.pptx' = '📽️ '; '.ps1'  = '📜 '; '.psd'  = '🎨 '
        '.py'   = '📜 '; '.rar'  = '📦 '; '.sh'   = '📜 '; '.sql'  = '🗄️ '
        '.step' = '🧊 '; '.stl'  = '🧊 '; '.stp'  = '🧊 '; '.svg'  = '🎨 '
        '.tar'  = '📦 '; '.tiff' = '🖼️ '; '.txt'  = '📄 '; '.wav'  = '🎵 '
        '.xls'  = '📊 '; '.xlsm' = '📊 '; '.xlsx' = '📊 '; '.xml'  = '⚙️ '
        '.yaml' = '⚙️ '; '.yml'  = '⚙️ '; '.zip'  = '📦 '
    }

    if ($Ruta -match '[\*\?]') {
        $Filtro = $Ruta
        $Ruta = "."
    }

    # ---------------------------------------------------------
    # LECTURA Y FILTRADO MANUAL SEGURO (Sin -Exclude nativo)
    # ---------------------------------------------------------
    # Traemos todo de la carpeta actual (incluyendo ocultos con -Force)
    $TodosLosElementos = Get-ChildItem -LiteralPath $Ruta -Force -ErrorAction SilentlyContinue | Sort-Object Name
    
    # Aplicamos el filtro de exclusión real antes de procesar nada
    if ($ListaFinalExcluir.Count -gt 0) {
        $TodosLosElementos = $TodosLosElementos | Where-Object { $_.Name -notIn $ListaFinalExcluir }
    }

    $Elementos = @()
    foreach ($e in $TodosLosElementos) {
        if (-not $e.PSIsContainer) {
            if ($e.Name -like $Filtro) { $Elementos += $e }
        } else {
            if ($Filtro -eq "*") {
                $Elementos += $e
            } else {
                # Para la verificación interna del filtro también limpiamos las carpetas excluidas
                $SubElementos = Get-ChildItem -Path $e.FullName -Recurse -File -Force -ErrorAction SilentlyContinue
                if ($ListaFinalExcluir.Count -gt 0) {
                    $SubElementos = $SubElementos | Where-Object { $_.Name -notIn $ListaFinalExcluir }
                }
                
                $TieneArchivos = $SubElementos | Where-Object { $_.Name -like $Filtro } | Select-Object -First 1
                if ($TieneArchivos) { $Elementos += $e }
            }
        }
    }

    # ---------------------------------------------------------
    # RENDERIZADO DEL ÁRBOL
    # ---------------------------------------------------------
    $Total = $Elementos.Count
    for ($i = 0; $i -lt $Total; $i++) {
        $Elemento = $Elementos[$i]
        $EsElUltimo = ($i -eq $Total - 1)
        
        $SimboloRama = if ($EsElUltimo) { "└── " } else { "├── " }
        
        if ($Elemento.PSIsContainer) {
            $Emoji = "📁 "
        } else {
            $Ext = $Elemento.Extension.ToLower() 
            $Emoji = if ($Iconos.ContainsKey($Ext)) { $Iconos[$Ext] } else { "📄 " }
        }
        
        Write-Output "${Prefijo}${SimboloRama}${Emoji}$($Elemento.Name)"
        
        if ($Elemento.PSIsContainer) {
            $Espacio = if ($EsElUltimo) { "    " } else { "│   " }
            $SiguientePrefijo = $Prefijo + $Espacio
            
            # Pasamos recursivamente la lista de exclusiones procesadas
            Mostrar-Arbol -Ruta $Elemento.FullName -Filtro $Filtro -Excluir $ListaFinalExcluir -Prefijo $SiguientePrefijo
        }
    }
}
