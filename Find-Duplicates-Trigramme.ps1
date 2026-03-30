<#
.SYNOPSIS
    Trouve les doublons dans la colonne "trigramme" et exporte toutes les lignes correspondantes

.PARAMETER FilePath
    Chemin complet du fichier Excel

.PARAMETER SheetName
    Nom de l'onglet (défaut: "Objets de gestion")

.PARAMETER ColumnName
    Nom de la colonne à analyser (défaut: "trigramme")

.PARAMETER OutputPath
    Fichier de sortie (optionnel)

.EXAMPLE
    .\Find-Duplicates-Trigramme.ps1 -FilePath "fichier.xlsx"
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$FilePath,
    [string]$SheetName = "Objets de gestion",
    [string]$ColumnName = "trigramme",
    [string]$OutputPath = ""
)

if (-not (Test-Path $FilePath)) {
    Write-Host "ERREUR: Fichier introuvable: $FilePath" -ForegroundColor Red
    exit 1
}

if ($OutputPath -eq "") {
    $fileDir = Split-Path -Parent $FilePath
    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    $OutputPath = Join-Path $fileDir "$fileName`_Doublons_Trigramme.xlsx"
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "RECHERCHE DES DOUBLONS - TRIGRAMME" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Fichier: $FilePath"
Write-Host "Onglet: $SheetName"
Write-Host ""

$excel = $null
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
        Write-Host "ERREUR: Onglet '$SheetName' introuvable" -ForegroundColor Red
        $workbook.Close($false)
        exit 1
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
        Write-Host "ERREUR: Colonne '$ColumnName' introuvable" -ForegroundColor Red
        $workbook.Close($false)
        exit 1
    }

    # Lire toutes les données
    $dataRange = $sourceSheet.Range($sourceSheet.Cells($headerRow + 1, 1), $sourceSheet.Cells($lastRow, $lastCol))
    $allData = $dataRange.Value2

    # Lire les en-têtes
    $headers = @()
    for ($col = 1; $col -le $lastCol; $col++) {
        $headers += $sourceSheet.Cells.Item($headerRow, $col).Text
    }

    $workbook.Close($false)

    # Traiter les données
    $dataRows = @()
    $numRows = $lastRow - $headerRow

    if ($numRows -eq 1) {
        $row = @{}
        for ($col = 0; $col -lt $lastCol; $col++) {
            $val = if ($allData -ne $null -and $col -eq 0) { $allData.ToString().Trim() } else { "" }
            $row[$headers[$col]] = $val
        }
        $dataRows += [PSCustomObject]$row
    } else {
        for ($i = 1; $i -le $numRows; $i++) {
            $row = @{}
            for ($col = 1; $col -le $lastCol; $col++) {
                $val = if ($allData[$i, $col] -ne $null) { $allData[$i, $col].ToString().Trim() } else { "" }
                $row[$headers[$col - 1]] = $val
            }
            $dataRows += [PSCustomObject]$row
        }
    }

    # Compter les occurrences
    $valueCounts = @{}
    foreach ($row in $dataRows) {
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

    Write-Host "Lignes: $numRows | Valeurs uniques: $($valueCounts.Count) | Doublons: $($duplicates.Count)" -ForegroundColor Yellow
    Write-Host ""

    if ($duplicates.Count -eq 0) {
        Write-Host "Aucun doublon trouvé" -ForegroundColor Green
        exit 0
    }

    # Afficher les 50 premiers doublons
    $displayCount = [Math]::Min(50, $duplicates.Count)
    foreach ($dup in ($duplicates | Sort-Object | Select-Object -First $displayCount)) {
        Write-Host "  $dup ($($valueCounts[$dup])x)"
    }
    if ($duplicates.Count -gt 50) {
        Write-Host "  ... et $(($duplicates.Count - 50)) autres" -ForegroundColor Gray
    }
    Write-Host ""

    # Filtrer les lignes avec doublons
    $dupRows = $dataRows | Where-Object { $_.$ColumnName -in $duplicates } | Sort-Object -Property $ColumnName

    Write-Host "Export de $($dupRows.Count) lignes..." -ForegroundColor Yellow

    # Créer le fichier Excel
    $outWorkbook = $excel.Workbooks.Add()
    $outSheet = $outWorkbook.Worksheets.Item(1)
    $outSheet.Name = "Doublons trigramme"

    # Créer tableau pour export
    $outArray = New-Object 'object[,]' ($dupRows.Count + 1), $headers.Count

    for ($col = 0; $col -lt $headers.Count; $col++) {
        $outArray[0, $col] = $headers[$col]
    }

    $rowIdx = 1
    foreach ($dataRow in $dupRows) {
        for ($col = 0; $col -lt $headers.Count; $col++) {
            $outArray[$rowIdx, $col] = if ($dataRow.($headers[$col]) -ne $null) { $dataRow.($headers[$col]) } else { "" }
        }
        $rowIdx++
    }

    # Écrire dans Excel
    $range = $outSheet.Range($outSheet.Cells(1, 1), $outSheet.Cells($dupRows.Count + 1, $headers.Count))
    $range.Value2 = $outArray

    # Formater
    $outSheet.Rows.Item(1).Font.Bold = $true
    $outSheet.Rows.Item(1).Interior.Color = 15773696
    for ($col = 1; $col -le $headers.Count; $col++) {
        try { $outSheet.Columns.Item($col).AutoFit() | Out-Null } catch {}
    }
    $outSheet.Range($outSheet.Cells(1, 1), $outSheet.Cells(1, $headers.Count)).AutoFilter() | Out-Null

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
