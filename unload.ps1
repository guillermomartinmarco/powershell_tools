# --- DESCARGAR ALIAS ---
if ($null -ne $global:MisAliasCargados -and $global:MisAliasCargados.Count -gt 0) {
    $ContadorAlias = 0
    foreach ($Atajo in $global:MisAliasCargados) {
        $RutaAlias = "Alias:\$Atajo"
        if (Test-Path $RutaAlias) {
            Remove-Item $RutaAlias -Force
            $ContadorAlias++
        }
    }
    $global:MisAliasCargados = @()
    Write-Host "🧹 Se eliminaron $ContadorAlias alias." -ForegroundColor Cyan
}

# --- DESCARGAR FUNCIONES ---
if ($null -ne $global:MisFuncionesCargadas -and $global:MisFuncionesCargadas.Count -gt 0) {
    $ContadorFunciones = 0
    foreach ($Funcion in $global:MisFuncionesCargadas) {
        $RutaFuncion = "Function:\$Funcion"
        if (Test-Path $RutaFuncion) {
            Remove-Item $RutaFuncion -Force
            $ContadorFunciones++
        }
    }
    $global:MisFuncionesCargadas = @()
    Write-Host "🧹 Se descargaron $ContadorFunciones funciones." -ForegroundColor Cyan
} else {
    Write-Host "No hay funciones pendientes de descargar." -ForegroundColor Gray
}
