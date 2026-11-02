function Convert-DocToDocx {
    <#
    .SYNOPSIS
        Script de alto rendimiento para la conversion masiva de archivos .doc a .docx.

    .DESCRIPTION
        Convierte los archivos .doc de la carpeta actual a .docx utilizando
        prioritariamente el motor nativo de conversion de Microsoft Office
        (wordconv.exe). Si no esta disponible, utiliza automatizacion COM con
        tipado estricto para evitar bloqueos.

    .PARAMETER DeleteSource
        Elimina los archivos .doc originales tras una conversion exitosa.

    .PARAMETER Overwrite
        Sobrescribe los archivos .docx si ya existen previamente.
    #>

    [CmdletBinding()]
    param (
        [switch]$DeleteSource,
        [switch]$Overwrite
    )

    $ErrorActionPreference = "Stop"
    $WorkingDir = (Get-Location).Path

    Clear-Host
    Write-Host "==================================================================" -ForegroundColor Cyan
    Write-Host "        CONVERTIDOR MASIVO DE DOCUMENTOS (.DOC -> .DOCX)         " -ForegroundColor Yellow
    Write-Host "==================================================================" -ForegroundColor Cyan
    Write-Host " Carpeta de trabajo: $WorkingDir" -ForegroundColor DarkGray
    Write-Host ""

    # Desbloquear preventivamente archivos de la zona de seguridad de Windows
    Get-ChildItem -LiteralPath $WorkingDir -Filter "*.doc" -File | Unblock-File -ErrorAction SilentlyContinue

    # Buscar la herramienta nativa wordconv.exe de Office
    $wordconvCandidates = @(
        "${env:ProgramFiles}\Microsoft Office\root\Office16\wordconv.exe",
        "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\wordconv.exe",
        "${env:ProgramFiles}\Microsoft Office\Office16\wordconv.exe",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office16\wordconv.exe",
        "${env:ProgramFiles}\Microsoft Office\Office15\wordconv.exe",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office15\wordconv.exe",
        "${env:ProgramFiles}\Microsoft Office\Office14\wordconv.exe",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office14\wordconv.exe"
    )

    $wordconvPath = $null
    foreach ($path in $wordconvCandidates) {
        if (Test-Path -LiteralPath $path) {
            $wordconvPath = $path
            break
        }
    }

    if ($wordconvPath) {
        Write-Host "[i] Motor seleccionado: Microsoft WordConv Nativo (Alta Velocidad)" -ForegroundColor Green
    } else {
        Write-Host "[i] Motor seleccionado: Microsoft Word COM Automation" -ForegroundColor Yellow
    }

    # Obtener archivos .doc
    $docFiles = @(Get-ChildItem -LiteralPath $WorkingDir -Filter "*.doc" -File | Where-Object { $_.Name -notlike "~$*" })
    $totalFiles = $docFiles.Count

    if ($totalFiles -eq 0) {
        Write-Host "[!] No se encontraron archivos .doc para convertir en esta carpeta." -ForegroundColor Yellow
        return
    }

    Write-Host (" Se encontraron {0} archivo(s) para procesar." -f $totalFiles) -ForegroundColor Green
    Write-Host "------------------------------------------------------------------" -ForegroundColor DarkGray

    $startTime = Get-Date
    $successCount = 0
    $skippedCount = 0
    $failCount = 0
    $currentIndex = 0

    if ($wordconvPath) {
        # MODO 1: MOTOR NATIVO WORDCONV (RAPIDO Y SIN BLOQUEOS)
        try {
            foreach ($file in $docFiles) {
                $currentIndex++
                $percent = [math]::Round(($currentIndex / $totalFiles) * 100)
                $docxPath = [string](Join-Path -Path $file.DirectoryName -ChildPath ($file.BaseName + ".docx"))

                $statusMsg = "Procesando ({0} de {1}): {2}" -f $currentIndex, $totalFiles, $file.Name
                Write-Progress -Activity "Convirtiendo .doc a .docx" -Status $statusMsg -PercentComplete $percent

                if ((Test-Path -LiteralPath $docxPath) -and (-not $Overwrite)) {
                    Write-Host ("  [OMITIDO] {0} (.docx ya existe)" -f $file.Name) -ForegroundColor Yellow
                    $skippedCount++
                    continue
                }

                # Argumentos: -oice (Open XML Conversion Engine) -nme (No More Extension)
                $processArgs = @("-oice", "-nme", "`"$($file.FullName)`"", "`"$docxPath`"")
                $proc = Start-Process -FilePath $wordconvPath -ArgumentList $processArgs -NoNewWindow -Wait -PassThru

                if ((Test-Path -LiteralPath $docxPath) -and (Get-Item -LiteralPath $docxPath).Length -gt 0) {
                    $successCount++
                    Write-Host ("  [OK] ({0}/{1}) {2}" -f $currentIndex, $totalFiles, $file.Name) -ForegroundColor Green

                    if ($DeleteSource) {
                        Remove-Item -LiteralPath $file.FullName -Force
                    }
                } else {
                    $failCount++
                    Write-Host ("  [ERROR] No se pudo convertir '{0}' (Codigo de salida: {1})" -f $file.Name, $proc.ExitCode) -ForegroundColor Red
                }
            }
        }
        finally {
            Write-Progress -Activity "Convirtiendo .doc a .docx" -Completed
        }
    }
    else {
        # MODO 2: MOTOR COM (CON FIX DE SERIALIZACION DE CADENAS)
        $runningWords = Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue
        if ($null -ne $runningWords) {
            $runningWords | Stop-Process -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 1
        }

        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0

        try {
            foreach ($file in $docFiles) {
                $currentIndex++
                $percent = [math]::Round(($currentIndex / $totalFiles) * 100)
                [string]$docxPath = Join-Path -Path $file.DirectoryName -ChildPath ($file.BaseName + ".docx")
                [string]$sourcePath = $file.FullName

                $statusMsg = "Procesando ({0} de {1}): {2}" -f $currentIndex, $totalFiles, $file.Name
                Write-Progress -Activity "Convirtiendo .doc a .docx" -Status $statusMsg -PercentComplete $percent

                if ((Test-Path -LiteralPath $docxPath) -and (-not $Overwrite)) {
                    Write-Host ("  [OMITIDO] {0} (.docx ya existe)" -f $file.Name) -ForegroundColor Yellow
                    $skippedCount++
                    continue
                }

                try {
                    $doc = $word.Documents.Open([string]$sourcePath, $false, $true)
                    # Formato wdFormatXMLDocument = 12, wdFormatDocumentDefault = 16
                    $doc.SaveAs2([string]$docxPath, [int]16)
                    $doc.Close(0)
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null

                    $successCount++
                    Write-Host ("  [OK] ({0}/{1}) {2}" -f $currentIndex, $totalFiles, $file.Name) -ForegroundColor Green

                    if ($DeleteSource) {
                        Remove-Item -LiteralPath $sourcePath -Force
                    }
                }
                catch {
                    $failCount++
                    Write-Host ("  [ERROR] Fallo en '{0}': {1}" -f $file.Name, $_.Exception.Message) -ForegroundColor Red
                }
            }
        }
        finally {
            Write-Progress -Activity "Convirtiendo .doc a .docx" -Completed
            if ($null -ne $word) {
                $word.Quit(0)
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
            }
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    }

    $elapsed = (Get-Date) - $startTime

    Write-Host ""
    Write-Host "==================================================================" -ForegroundColor Cyan
    Write-Host "                    RESUMEN DE LA OPERACION                       " -ForegroundColor Yellow
    Write-Host "==================================================================" -ForegroundColor Cyan
    Write-Host (" Total detectados:         {0}" -f $totalFiles) -ForegroundColor White
    Write-Host (" Convertidos con exito:    {0}" -f $successCount) -ForegroundColor Green
    Write-Host (" Omitidos (Ya existian):   {0}" -f $skippedCount) -ForegroundColor Yellow

    $errorColor = if ($failCount -gt 0) { "Red" } else { "DarkGray" }
    Write-Host (" Errores de conversion:    {0}" -f $failCount) -ForegroundColor $errorColor

    Write-Host (" Tiempo total:             {0}m {1}s" -f $elapsed.Minutes, $elapsed.Seconds) -ForegroundColor Cyan
    Write-Host "==================================================================" -ForegroundColor Cyan
    Write-Host ""
}
