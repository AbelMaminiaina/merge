<#
.SYNOPSIS
    Analyse les fichiers Excel - Unicite Nom FR et Trigramme avec descriptions et predictions

.PARAMETER FilePath
    Chemin du fichier Excel

.PARAMETER SheetName
    Nom de l'onglet (defaut: "Objets de gestion")

.EXAMPLE
    .\Analyze-ExcelData.ps1 -FilePath "fichier.xlsx"
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$FilePath,
    [string]$SheetName = "Objets de gestion"
)

$NomFRColumn = "Nom FR"
$TrigrammeColumn = "trigramme"

if (-not (Test-Path $FilePath)) {
    Write-Host "ERREUR: Fichier introuvable: $FilePath" -ForegroundColor Red
    exit 1
}

Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host "              ANALYSE DES DONNEES EXCEL - UNICITE ET PREDICTIONS               " -ForegroundColor Cyan
Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host "Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Host "Fichier: $([System.IO.Path]::GetFileName($FilePath))"
Write-Host "Onglet: $SheetName"
Write-Host ""

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false

    Write-Host "Ouverture du fichier..." -ForegroundColor Gray
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

    $usedRange = $sourceSheet.UsedRange
    $lastRow = $usedRange.Rows.Count
    $lastCol = $usedRange.Columns.Count

    Write-Host "Lecture des donnees (bulk)..." -ForegroundColor Gray

    # Lire TOUTES les donnees en une seule operation (rapide!)
    $allData = $usedRange.Value2

    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    $excel = $null
    [System.GC]::Collect()

    Write-Host "Analyse en cours..." -ForegroundColor Gray

    # Trouver les colonnes dans les en-tetes
    $nomFRCol = 0
    $trigrammeCol = 0
    $headerRow = 0

    for ($row = 1; $row -le [Math]::Min(30, $lastRow); $row++) {
        for ($col = 1; $col -le $lastCol; $col++) {
            $cellValue = ""
            if ($lastRow -eq 1 -and $lastCol -eq 1) {
                $cellValue = if ($allData) { $allData.ToString() } else { "" }
            } elseif ($lastRow -eq 1) {
                $cellValue = if ($allData[1, $col]) { $allData[1, $col].ToString() } else { "" }
            } elseif ($lastCol -eq 1) {
                $cellValue = if ($allData[$row, 1]) { $allData[$row, 1].ToString() } else { "" }
            } else {
                $cellValue = if ($allData[$row, $col]) { $allData[$row, $col].ToString() } else { "" }
            }

            if ($cellValue -eq $NomFRColumn -and $nomFRCol -eq 0) {
                $nomFRCol = $col
                $headerRow = $row
            }
            if ($cellValue -eq $TrigrammeColumn -and $trigrammeCol -eq 0) {
                $trigrammeCol = $col
                if ($headerRow -eq 0) { $headerRow = $row }
            }
        }
        if ($nomFRCol -gt 0 -and $trigrammeCol -gt 0) { break }
    }

    if ($headerRow -eq 0) {
        Write-Host "ERREUR: Colonnes '$NomFRColumn' et '$TrigrammeColumn' introuvables" -ForegroundColor Red
        exit 1
    }

    # Extraire les valeurs
    $numRows = $lastRow - $headerRow
    $nomFRValues = @()
    $trigrammeValues = @()

    for ($row = $headerRow + 1; $row -le $lastRow; $row++) {
        if ($nomFRCol -gt 0) {
            $val = ""
            if ($lastCol -eq 1) {
                $val = if ($allData[$row, 1]) { $allData[$row, 1].ToString().Trim() } else { "" }
            } else {
                $val = if ($allData[$row, $nomFRCol]) { $allData[$row, $nomFRCol].ToString().Trim() } else { "" }
            }
            $nomFRValues += $val
        }
        if ($trigrammeCol -gt 0) {
            $val = ""
            if ($lastCol -eq 1) {
                $val = if ($allData[$row, 1]) { $allData[$row, 1].ToString().Trim() } else { "" }
            } else {
                $val = if ($allData[$row, $trigrammeCol]) { $allData[$row, $trigrammeCol].ToString().Trim() } else { "" }
            }
            $trigrammeValues += $val
        }
    }

    # Affichage des resultats
    Write-Host ""
    Write-Host "--------------------------------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host "  STATISTIQUES GENERALES" -ForegroundColor White
    Write-Host "--------------------------------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host "  Total lignes de donnees: $numRows"
    Write-Host ""

    # === NOM FR ===
    Write-Host "  NOM FR" -ForegroundColor Cyan
    Write-Host "  -------"

    $nomFRDupCount = 0
    $nomFRUniqueCount = 0
    $nomFREmptyCount = 0
    $nomFRIsUnique = $true
    $nomFRDuplicates = @()

    if ($nomFRCol -gt 0) {
        $nomFRCounts = @{}
        foreach ($val in $nomFRValues) {
            if ([string]::IsNullOrWhiteSpace($val)) {
                $nomFREmptyCount++
            } else {
                if ($nomFRCounts.ContainsKey($val)) {
                    $nomFRCounts[$val]++
                } else {
                    $nomFRCounts[$val] = 1
                }
            }
        }

        $nomFRDuplicates = @($nomFRCounts.GetEnumerator() | Where-Object { $_.Value -gt 1 })
        $nomFRUniqueCount = $nomFRCounts.Count
        $nomFRDupCount = $nomFRDuplicates.Count
        $nomFRIsUnique = ($nomFRDupCount -eq 0)

        Write-Host "    Valeurs uniques:    $nomFRUniqueCount"
        Write-Host "    Valeurs en doublon: $nomFRDupCount"
        Write-Host "    Valeurs vides:      $nomFREmptyCount"
        Write-Host -NoNewline "    EST UNIQUE:         "
        if ($nomFRIsUnique) {
            Write-Host "OUI" -ForegroundColor Green
        } else {
            Write-Host "NON" -ForegroundColor Red
        }

        if ($nomFRDuplicates.Count -gt 0) {
            Write-Host ""
            Write-Host "    Doublons (top 10):" -ForegroundColor Yellow
            $top10 = $nomFRDuplicates | Sort-Object { $_.Value } -Descending | Select-Object -First 10
            foreach ($dup in $top10) {
                Write-Host "      - '$($dup.Key)' ($($dup.Value)x)"
            }
        }

        $nonEmpty = @($nomFRValues | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        if ($nonEmpty.Count -gt 0) {
            $avgLen = ($nonEmpty | ForEach-Object { $_.Length } | Measure-Object -Average).Average
            Write-Host ""
            Write-Host "    Patterns:" -ForegroundColor Magenta
            Write-Host "      - Longueur moyenne: $([Math]::Round($avgLen, 1)) caracteres"
        }
    } else {
        Write-Host "    Colonne non trouvee" -ForegroundColor Red
    }

    Write-Host ""

    # === TRIGRAMME ===
    Write-Host "  TRIGRAMME" -ForegroundColor Cyan
    Write-Host "  ----------"

    $triDupCount = 0
    $triUniqueCount = 0
    $triEmptyCount = 0
    $triIsUnique = $true
    $triDuplicates = @()

    if ($trigrammeCol -gt 0) {
        $triCounts = @{}
        foreach ($val in $trigrammeValues) {
            if ([string]::IsNullOrWhiteSpace($val)) {
                $triEmptyCount++
            } else {
                if ($triCounts.ContainsKey($val)) {
                    $triCounts[$val]++
                } else {
                    $triCounts[$val] = 1
                }
            }
        }

        $triDuplicates = @($triCounts.GetEnumerator() | Where-Object { $_.Value -gt 1 })
        $triUniqueCount = $triCounts.Count
        $triDupCount = $triDuplicates.Count
        $triIsUnique = ($triDupCount -eq 0)

        Write-Host "    Valeurs uniques:    $triUniqueCount"
        Write-Host "    Valeurs en doublon: $triDupCount"
        Write-Host "    Valeurs vides:      $triEmptyCount"
        Write-Host -NoNewline "    EST UNIQUE:         "
        if ($triIsUnique) {
            Write-Host "OUI" -ForegroundColor Green
        } else {
            Write-Host "NON" -ForegroundColor Red
        }

        if ($triDuplicates.Count -gt 0) {
            Write-Host ""
            Write-Host "    Doublons (top 10):" -ForegroundColor Yellow
            $top10 = $triDuplicates | Sort-Object { $_.Value } -Descending | Select-Object -First 10
            foreach ($dup in $top10) {
                Write-Host "      - '$($dup.Key)' ($($dup.Value)x)"
            }
        }

        $nonEmpty = @($trigrammeValues | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        $valid3 = @($nonEmpty | Where-Object { $_.Length -eq 3 })
        Write-Host ""
        Write-Host "    Patterns:" -ForegroundColor Magenta
        Write-Host "      - Trigrammes valides (3 car.): $($valid3.Count)/$($nonEmpty.Count)"
    } else {
        Write-Host "    Colonne non trouvee" -ForegroundColor Red
    }

    # Score de qualite
    $score = 100
    if ($nomFRDupCount -gt 0 -and $nomFRUniqueCount -gt 0) {
        $score -= [Math]::Min(25, ($nomFRDupCount / $nomFRUniqueCount) * 100)
    }
    if ($triDupCount -gt 0 -and $triUniqueCount -gt 0) {
        $score -= [Math]::Min(25, ($triDupCount / $triUniqueCount) * 100)
    }
    if ($numRows -gt 0) {
        $score -= [Math]::Min(20, (($nomFREmptyCount + $triEmptyCount) / ($numRows * 2)) * 100)
    }
    $score = [Math]::Max(0, [Math]::Round($score, 1))

    Write-Host ""
    Write-Host "--------------------------------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host -NoNewline "  SCORE DE QUALITE: "
    if ($score -ge 90) { Write-Host "$score/100" -ForegroundColor Green }
    elseif ($score -ge 70) { Write-Host "$score/100" -ForegroundColor Yellow }
    elseif ($score -ge 50) { Write-Host "$score/100" -ForegroundColor DarkYellow }
    else { Write-Host "$score/100" -ForegroundColor Red }

    Write-Host ""
    Write-Host "  PREDICTIONS ET RECOMMANDATIONS" -ForegroundColor Magenta
    Write-Host "  -------------------------------"

    if ($nomFRDupCount -gt 0) {
        Write-Host "    ATTENTION: $nomFRDupCount valeurs Nom FR en doublon peuvent causer des conflits" -ForegroundColor Red
    }
    if ($triDupCount -gt 0) {
        Write-Host "    ATTENTION: $triDupCount trigrammes en doublon - risque d'erreurs d'identification" -ForegroundColor Red
    }
    if ($nomFREmptyCount -gt 0) {
        Write-Host "    INFO: $nomFREmptyCount lignes sans Nom FR - donnees potentiellement incompletes" -ForegroundColor Yellow
    }
    if ($triEmptyCount -gt 0) {
        Write-Host "    INFO: $triEmptyCount lignes sans trigramme - donnees potentiellement incompletes" -ForegroundColor Yellow
    }

    if ($score -ge 90) {
        Write-Host "    QUALITE: Excellente - donnees coherentes et uniques" -ForegroundColor Cyan
    } elseif ($score -ge 70) {
        Write-Host "    QUALITE: Bonne - quelques corrections recommandees" -ForegroundColor Cyan
    } elseif ($score -ge 50) {
        Write-Host "    QUALITE: Moyenne - nettoyage necessaire avant utilisation" -ForegroundColor Cyan
    } else {
        Write-Host "    QUALITE: Faible - revision majeure des donnees recommandee" -ForegroundColor Cyan
    }

    Write-Host ""
    Write-Host "================================================================================" -ForegroundColor Cyan
    Write-Host "                              ANALYSE TERMINEE                                  " -ForegroundColor Cyan
    Write-Host "================================================================================" -ForegroundColor Cyan

} catch {
    Write-Host "ERREUR: $_" -ForegroundColor Red
    exit 1
} finally {
    if ($excel) {
        try { $excel.Quit() } catch {}
        try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
