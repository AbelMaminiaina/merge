# ==============================================================================
# Merge-Simple.ps1
# ==============================================================================
# Script simplifié pour fusionner les onglets "Objets de gestion"
# Gère les doublons sur "Nom FR"
# ==============================================================================

param(
    [string]$Folder = "C:\Users\amami\github\merge\test",
    [string]$Output = "C:\Users\amami\github\merge\test\Fusion.xlsx"
)

Write-Host "=== FUSION SIMPLE ===" -ForegroundColor Cyan

# Fermer Excel
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

# Récupérer les fichiers
$Files = Get-ChildItem -Path $Folder -Filter "*.xlsx" | Where-Object { $_.FullName -ne $Output }

if ($Files.Count -eq 0) {
    Write-Host "Aucun fichier Excel" -ForegroundColor Red
    exit
}

Write-Host "Fichiers: $($Files.Count)`n"

# Stockage des données
$AllData = @{}
$Headers = @()
$doublons = 0

# Lire chaque fichier
foreach ($file in $Files) {
    Write-Host "Lecture: $($file.Name)..." -ForegroundColor Yellow -NoNewline

    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false

    try {
        $WB = $Excel.Workbooks.Open($file.FullName, 0, $true)
        $Sheet = $WB.Worksheets | Where-Object { $_.Name -eq "Objets de gestion" }

        if ($Sheet) {
            $Data = $Sheet.UsedRange.Value2
            $Rows = $Sheet.UsedRange.Rows.Count
            $Cols = $Sheet.UsedRange.Columns.Count

            # En-têtes (première lecture uniquement)
            if ($Headers.Count -eq 0) {
                for ($c = 1; $c -le $Cols; $c++) {
                    $Headers += $Data[1, $c]
                }
            }

            # Trouver index "Nom FR"
            $NomFRIdx = 0
            for ($c = 0; $c -lt $Headers.Count; $c++) {
                if ($Headers[$c] -eq "Nom FR") { $NomFRIdx = $c + 1; break }
            }

            # Lire les données
            $added = 0
            for ($r = 2; $r -le $Rows; $r++) {
                $key = $Data[$r, $NomFRIdx]
                if ([string]::IsNullOrEmpty($key)) { continue }

                if (-not $AllData.ContainsKey($key)) {
                    $row = @()
                    for ($c = 1; $c -le $Cols; $c++) {
                        $row += $Data[$r, $c]
                    }
                    $AllData[$key] = $row
                    $added++
                } else {
                    $doublons++
                }
            }
            Write-Host " $added lignes" -ForegroundColor Green
        } else {
            Write-Host " Onglet non trouve" -ForegroundColor DarkYellow
        }

        $WB.Close($false)
    }
    finally {
        $Excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
        [System.GC]::Collect()
    }
}

if ($AllData.Count -eq 0) {
    Write-Host "Aucune donnee!" -ForegroundColor Red
    exit
}

# Créer le fichier de sortie
Write-Host "`nCreation fichier..." -ForegroundColor Yellow

$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$Excel.DisplayAlerts = $false

try {
    $WB = $Excel.Workbooks.Add()
    $Sheet = $WB.Worksheets.Item(1)
    $Sheet.Name = "Objets de gestion"

    # Préparer données en bloc (plus rapide)
    $rowCount = $AllData.Count + 1++
    $colCount = $Headers.Count
    $OutputData = New-Object 'object[,]' $rowCount, $colCount

    # En-têtes
    for ($c = 0; $c -lt $colCount; $c++) {
        $OutputData[0, $c] = $Headers[$c]
    }

    # Données triées par Nom FR
    $row = 1
    foreach ($key in ($AllData.Keys | Sort-Object)) {
        $data = $AllData[$key]
        for ($c = 0; $c -lt $data.Count; $c++) {
            $OutputData[$row, $c] = $data[$c]
        }
        $row++
    }

    # Écriture en bloc
    $lastCol = [char]([int][char]'A' + [Math]::Min($colCount - 1, 25))
    $Range = $Sheet.Range("A1:$($lastCol)$rowCount")
    $Range.Value2 = $OutputData

    # Formatage simple
    $lastCol = [char]([int][char]'A' + $Headers.Count - 1)
    $Sheet.Range("A1:$($lastCol)1").Font.Bold = $true
    $Sheet.Range("A1:$($lastCol)1").Interior.ColorIndex = 15
    $Sheet.Range("A1:$($lastCol)1").AutoFilter() | Out-Null
    $Sheet.UsedRange.Columns.AutoFit() | Out-Null

    # Première colonne (date) : format français + aligné à droite
    $dateRange = $Sheet.Range("A2:A$rowCount")
    $dateRange.NumberFormat = "[$-40C]dd/mm/yyyy"  # 40C = code locale français
    $dateRange.HorizontalAlignment = -4152  # xlRight

    # Sauvegarder
    if (Test-Path $Output) { Remove-Item $Output -Force }
    $WB.SaveAs($Output, 51)
    $WB.Close($false)

    Write-Host "Fichier cree!" -ForegroundColor Green
}
finally {
    $Excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
    [System.GC]::Collect()
}

# Résumé
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "FUSION TERMINEE" -ForegroundColor Cyan
Write-Host "Uniques: $($AllData.Count) | Doublons: $doublons" -ForegroundColor White
Write-Host "========================================" -ForegroundColor Cyan

# Ouvrir
Start-Process $Output
