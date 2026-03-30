<#
.SYNOPSIS
    Trouve les doublons dans la colonne "trigramme" sur plusieurs fichiers Excel

.PARAMETER FilePaths
    Chemins des fichiers Excel (accepte plusieurs fichiers)

.PARAMETER FolderPath
    Chemin d'un dossier contenant les fichiers Excel

.PARAMETER SheetName
    Nom de l'onglet (défaut: "Objets de gestion")

.PARAMETER ColumnName
    Nom de la colonne à analyser (défaut: "trigramme")

.PARAMETER OutputPath
    Fichier de sortie (optionnel)

.EXAMPLE
    .\Find-Duplicates-Multi-Trigramme.ps1 -FilePaths "fichier1.xlsx","fichier2.xlsx"
    .\Find-Duplicates-Multi-Trigramme.ps1 -FolderPath "C:\dossier"
#>

param(
    [Parameter(Mandatory=$false)]
    [string[]]$FilePaths,
    [string]$FolderPath = "",
    [string]$SheetName = "Objets de gestion",
    [string]$ColumnName = "trigramme",
    [string]$OutputPath = ""
)

# Collecter les fichiers
$allFiles = @()

if ($FolderPath -ne "" -and (Test-Path $FolderPath)) {
    $allFiles = Get-ChildItem -Path $FolderPath -Filter "*.xlsx" | Where-Object { $_.Name -notlike "~*" -and $_.Name -notlike "*_Doublons_*" -and $_.Name -notlike "Doublons_*" } | Select-Object -ExpandProperty FullName
} elseif ($FilePaths -and $FilePaths.Count -gt 0) {
    $allFiles = $FilePaths
} else {
    Write-Host "ERREUR: Spécifiez -FilePaths ou -FolderPath" -ForegroundColor Red
    exit 1
}

if ($allFiles.Count -eq 0) {
    Write-Host "ERREUR: Aucun fichier Excel trouvé" -ForegroundColor Red
    exit 1
}

# Vérifier les fichiers
foreach ($f in $allFiles) {
    if (-not (Test-Path $f)) {
        Write-Host "ERREUR: Fichier introuvable: $f" -ForegroundColor Red
        exit 1
    }
}

if ($OutputPath -eq "") {
    if ($FolderPath -ne "") {
        $OutputPath = Join-Path $FolderPath "Doublons_Multi_Trigramme.xlsx"
    } else {
        $fileDir = Split-Path -Parent $allFiles[0]
        $OutputPath = Join-Path $fileDir "Doublons_Multi_Trigramme.xlsx"
    }
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "RECHERCHE DES DOUBLONS MULTI-FICHIERS" -ForegroundColor Cyan
Write-Host "Colonne: TRIGRAMME" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Fichiers: $($allFiles.Count)"
Write-Host "Onglet: $SheetName"
Write-Host ""

$allDataRows = @()
$headers = $null

# Fonction pour lire un fichier Excel
function Read-ExcelFile {
    param($FilePath, $SheetName, $ColumnName)

    $fileName = [System.IO.Path]::GetFileName($FilePath)
    $result = @{
        Success = $false
        Headers = $null
        Rows = @()
    }

    $excel = $null
    $workbook = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = $false

        $workbook = $excel.Workbooks.Open($FilePath, 0, $true)

        # Trouver l'onglet
        $sourceSheet = $null
        foreach ($ws in $workbook.Worksheets) {
            if ($ws.Name -eq $SheetName) {
                $sourceSheet = $ws
                break
            }
        }

        if ($null -eq $sourceSheet) {
            Write-Host "  ATTENTION: Onglet '$SheetName' introuvable" -ForegroundColor Yellow
            return $result
        }

        # Trouver la colonne cible
        $usedRange = $sourceSheet.UsedRange
        $lastRow = $usedRange.Rows.Count
        $lastCol = $usedRange.Columns.Count
        $targetColIndex = $null
        $headerRow = $null

        for ($row = 1; $row -le [Math]::Min(30, $lastRow); $row++) {
            for ($col = 1; $col -le $lastCol; $col++) {
                if ($sourceSheet.Cells.Item($row, $col).Text -eq $ColumnName) {
                    $targetColIndex = $col
                    $headerRow = $row
                    break
                }
            }
            if ($targetColIndex) { break }
        }

        if (-not $targetColIndex) {
            Write-Host "  ATTENTION: Colonne '$ColumnName' introuvable" -ForegroundColor Yellow
            return $result
        }

        # Lire les en-têtes
        $fileHeaders = @()
        for ($col = 1; $col -le $lastCol; $col++) {
            $fileHeaders += $sourceSheet.Cells.Item($headerRow, $col).Text
        }
        $result.Headers = $fileHeaders

        # Lire toutes les données
        $dataRange = $sourceSheet.Range($sourceSheet.Cells($headerRow + 1, 1), $sourceSheet.Cells($lastRow, $lastCol))
        $allData = $dataRange.Value2

        $numRows = $lastRow - $headerRow
        $dataRows = @()

        if ($numRows -eq 1) {
            $row = @{}
            $row["Fichier Source"] = $fileName
            for ($col = 0; $col -lt $lastCol; $col++) {
                $val = if ($allData -ne $null -and $col -eq 0) { $allData.ToString().Trim() } else { "" }
                $row[$fileHeaders[$col]] = $val
            }
            $dataRows += [PSCustomObject]$row
        } else {
            for ($i = 1; $i -le $numRows; $i++) {
                $row = @{}
                $row["Fichier Source"] = $fileName
                for ($col = 1; $col -le $lastCol; $col++) {
                    $val = if ($allData[$i, $col] -ne $null) { $allData[$i, $col].ToString().Trim() } else { "" }
                    $row[$fileHeaders[$col - 1]] = $val
                }
                $dataRows += [PSCustomObject]$row
            }
        }

        $result.Rows = $dataRows
        $result.Success = $true
        Write-Host "  $numRows lignes lues" -ForegroundColor Green

    } catch {
        Write-Host "  ERREUR: $_" -ForegroundColor Red
    } finally {
        if ($workbook) {
            try { $workbook.Close($false) } catch { }
            try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null } catch { }
        }
        if ($excel) {
            try { $excel.Quit() } catch { }
            try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null } catch { }
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }

    return $result
}

# Lire chaque fichier
foreach ($filePath in $allFiles) {
    $fileName = [System.IO.Path]::GetFileName($filePath)
    Write-Host "Lecture: $fileName" -ForegroundColor Gray

    $result = Read-ExcelFile -FilePath $filePath -SheetName $SheetName -ColumnName $ColumnName

    if ($result.Success) {
        if ($null -eq $headers) {
            $headers = $result.Headers
        }
        $allDataRows += $result.Rows
    }

    Start-Sleep -Milliseconds 500
}

if ($null -eq $headers -or $allDataRows.Count -eq 0) {
    Write-Host "ERREUR: Aucune donnée trouvée" -ForegroundColor Red
    exit 1
}

# Ajouter "Fichier Source" aux en-têtes
$outputHeaders = @("Fichier Source") + $headers

# Compter les occurrences
$valueCounts = @{}
foreach ($row in $allDataRows) {
    $val = $row.$ColumnName
    if ($val -and $val -ne '') {
        if ($valueCounts.ContainsKey($val)) {
            $valueCounts[$val]++
        } else {
            $valueCounts[$val] = 1
        }
    }
}

# Trouver les doublons
$duplicates = $valueCounts.GetEnumerator() | Where-Object { $_.Value -gt 1 } | Select-Object -ExpandProperty Key

Write-Host ""
Write-Host "Lignes totales: $($allDataRows.Count) | Valeurs uniques: $($valueCounts.Count) | Doublons: $($duplicates.Count)" -ForegroundColor Yellow
Write-Host ""

if ($duplicates.Count -eq 0) {
    Write-Host "Aucun doublon trouvé" -ForegroundColor Green
    exit 0
}

$displayCount = [Math]::Min(50, $duplicates.Count)
foreach ($dup in ($duplicates | Sort-Object | Select-Object -First $displayCount)) {
    Write-Host "  $dup ($($valueCounts[$dup])x)"
}
if ($duplicates.Count -gt 50) {
    Write-Host "  ... et $(($duplicates.Count - 50)) autres" -ForegroundColor Gray
}
Write-Host ""

# Filtrer les lignes avec doublons
$dupRows = $allDataRows | Where-Object { $_.$ColumnName -in $duplicates } | Sort-Object -Property $ColumnName

Write-Host "Export de $($dupRows.Count) lignes..." -ForegroundColor Yellow

# Créer le fichier Excel de sortie
$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false

    $outWorkbook = $excel.Workbooks.Add()
    $outSheet = $outWorkbook.Worksheets.Item(1)
    $outSheet.Name = "Doublons trigramme"

    # Créer tableau pour export
    $outArray = New-Object 'object[,]' ($dupRows.Count + 1), $outputHeaders.Count

    for ($col = 0; $col -lt $outputHeaders.Count; $col++) {
        $outArray[0, $col] = $outputHeaders[$col]
    }

    $rowIdx = 1
    foreach ($dataRow in $dupRows) {
        for ($col = 0; $col -lt $outputHeaders.Count; $col++) {
            $outArray[$rowIdx, $col] = if ($dataRow.($outputHeaders[$col]) -ne $null) { $dataRow.($outputHeaders[$col]) } else { "" }
        }
        $rowIdx++
    }

    # Écrire dans Excel
    $range = $outSheet.Range($outSheet.Cells(1, 1), $outSheet.Cells($dupRows.Count + 1, $outputHeaders.Count))
    $range.Value2 = $outArray

    # Formater
    $outSheet.Rows.Item(1).Font.Bold = $true
    $outSheet.Rows.Item(1).Interior.Color = 15773696
    for ($col = 1; $col -le $outputHeaders.Count; $col++) {
        try { $outSheet.Columns.Item($col).AutoFit() | Out-Null } catch {}
    }
    $outSheet.Range($outSheet.Cells(1, 1), $outSheet.Cells(1, $outputHeaders.Count)).AutoFilter() | Out-Null

    # Sauvegarder
    if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
    $outWorkbook.SaveAs($OutputPath)
    $outWorkbook.Close($false)

    Write-Host ""
    Write-Host "TERMINÉ" -ForegroundColor Green
    Write-Host "Fichier: $OutputPath" -ForegroundColor Cyan
    Write-Host ""

} catch {
    Write-Host "ERREUR: $_" -ForegroundColor Red
    exit 1
} finally {
    if ($excel) {
        try { $excel.Workbooks.Close() } catch {}
        try { $excel.Quit() } catch {}
        try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
