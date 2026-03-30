# ==============================================================================
# Merge-ExcelObjets.ps1
# ==============================================================================
# Description : Fusionne les onglets "Objets de gestion" de tous les fichiers
#               Excel d'un dossier en un seul fichier, sans doublons.
# Clé unique  : Colonne "Nom FR" (évite les doublons)
# Formatage   : Copie le formatage exact du premier fichier source
# ==============================================================================

# ------------------------------------------------------------------------------
# PARAMETRES D'ENTREE
# ------------------------------------------------------------------------------
# $Folder : Dossier contenant les fichiers Excel sources
# $Output : Chemin du fichier Excel de sortie (fusion)
# ------------------------------------------------------------------------------
param(
    [string]$Folder = "C:\Users\amami\GitHub\merge\test",
    [string]$Output = "C:\Users\amami\GitHub\merge\test\merge\Maquette.xlsx"
)

Write-Host "=== FUSION DES ONGLETS OBJETS DE GESTION ===" -ForegroundColor Cyan

# ------------------------------------------------------------------------------
# FERMETURE DES INSTANCES EXCEL
# ------------------------------------------------------------------------------
# Ferme toutes les instances Excel en cours pour éviter les conflits COM
# Attente de 5 secondes pour laisser le temps aux processus de se terminer
# ------------------------------------------------------------------------------
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 5

# ------------------------------------------------------------------------------
# RECUPERATION DES FICHIERS SOURCES
# ------------------------------------------------------------------------------
# Liste tous les fichiers .xlsx du dossier, en excluant le fichier de sortie
# ------------------------------------------------------------------------------
$Sources = Get-ChildItem -Path $Folder -Filter "*.xlsx" |
           Where-Object { $_.FullName -ne $Output } |
           Select-Object -ExpandProperty FullName

if ($Sources.Count -eq 0) {
    Write-Host "Aucun fichier Excel trouvé" -ForegroundColor Red
    exit
}

Write-Host "Fichiers: $($Sources.Count) | Sortie: $Output`n"

# ------------------------------------------------------------------------------
# VARIABLES DE STOCKAGE
# ------------------------------------------------------------------------------
# $AllData         : Hashtable pour stocker toutes les lignes (clé = Nom FR)
# $AllHeaders      : Liste de tous les en-têtes uniques trouvés
# $totalLignes     : Compteur total de lignes lues
# $doublons        : Compteur de doublons ignorés
# ------------------------------------------------------------------------------
$AllData = @{}
$AllHeaders = [System.Collections.ArrayList]@()
$totalLignes = 0
$doublons = 0

# ==============================================================================
# FONCTION : Read-ExcelFile
# ==============================================================================
# Description : Lit un fichier Excel et extrait les données de l'onglet
#               "Objets de gestion" avec système de retry automatique
# Paramètres  :
#   - $FilePath : Chemin du fichier Excel à lire
#   - $MaxRetries : Nombre maximum de tentatives (défaut: 3)
# Retourne    : Hashtable avec Data, Rows, Cols
# ==============================================================================
function Read-ExcelFile {
    param(
        $FilePath,
        [int]$MaxRetries = 3
    )

    $result = $null
    $attempt = 0

    while ($attempt -lt $MaxRetries -and $null -eq $result) {
        $attempt++

        if ($attempt -gt 1) {
            Write-Host "  Tentative $attempt/$MaxRetries..." -ForegroundColor Yellow -NoNewline
            Start-Sleep -Seconds 3
        } else {
            Write-Host "Lecture de $FilePath..." -ForegroundColor DarkYellow
        }

        $Excel = $null
        $Workbook = $null
        $Sheet = $null
        $UsedRange = $null
        $Worksheets = $null

        try {
            # ------------------------------------------------------------------
            # INITIALISATION DE L'APPLICATION EXCEL
            # ------------------------------------------------------------------
            $Excel = New-Object -ComObject Excel.Application
            $Excel.Visible = $false
            $Excel.DisplayAlerts = $false
            $Excel.ScreenUpdating = $false

            # Ouvrir le fichier en lecture seule
            $Workbook = $Excel.Workbooks.Open($FilePath, $false, $true)
            $Worksheets = $Workbook.Worksheets

            # ------------------------------------------------------------------
            # RECHERCHE DE L'ONGLET "Objets de gestion"
            # ------------------------------------------------------------------
            foreach ($ws in $Worksheets) {
                if ($ws.Name -eq "Objets de gestion") {
                    $Sheet = $ws
                    break
                }
            }

            if ($Sheet) {
                # --------------------------------------------------------------
                # EXTRACTION DES DONNEES
                # --------------------------------------------------------------
                $UsedRange = $Sheet.UsedRange
                $ColCount = $UsedRange.Columns.Count

                $result = @{
                    Data = $UsedRange.Value2
                    Rows = $UsedRange.Rows.Count
                    Cols = $ColCount
                }
            }

            $Workbook.Close($false)
        }
        catch {
            if ($attempt -eq $MaxRetries) {
                Write-Host "Erreur finale: $_" -ForegroundColor Red
            } else {
                Write-Host "Erreur (retry...): $_" -ForegroundColor DarkYellow -NoNewline
            }
        }
        finally {
            # ------------------------------------------------------------------
            # LIBERATION DE TOUS LES OBJETS COM (IMPORTANT!)
            # ------------------------------------------------------------------
            if ($UsedRange) {
                try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($UsedRange) | Out-Null } catch { }
                $UsedRange = $null
            }
            if ($Sheet) {
                try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Sheet) | Out-Null } catch { }
                $Sheet = $null
            }
            if ($Worksheets) {
                try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Worksheets) | Out-Null } catch { }
                $Worksheets = $null
            }
            if ($Workbook) {
                try { $Workbook.Close($false) } catch { }
                try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook) | Out-Null } catch { }
                $Workbook = $null
            }
            if ($Excel) {
                try { $Excel.Quit() } catch { }
                try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null } catch { }
                $Excel = $null
            }
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            [System.GC]::Collect()
            Start-Sleep -Seconds 2
        }
    }

    return $result
}

# ==============================================================================
# BOUCLE DE LECTURE DES FICHIERS SOURCES
# ==============================================================================
# Parcourt chaque fichier Excel et extrait les données de l'onglet
# ==============================================================================
foreach ($sourceFile in $Sources) {
    $fileName = Split-Path $sourceFile -Leaf
    Write-Host "Lecture de $fileName..." -ForegroundColor Yellow -NoNewline

    $result = Read-ExcelFile -FilePath $sourceFile

    if ($null -eq $result) {
        Write-Host " Onglet non trouvé" -ForegroundColor DarkYellow
        continue
    }
    $Data = $result.Data
    $RowCount = $result.Rows
    $ColCount = $result.Cols

    # --------------------------------------------------------------------------
    # EXTRACTION DES EN-TETES (ligne 1)
    # --------------------------------------------------------------------------
    $Headers = @()
    for ($c = 1; $c -le $ColCount; $c++) {
        $Headers += $Data[1, $c]
    }

    # --------------------------------------------------------------------------
    # RECHERCHE DE L'INDEX DE LA COLONNE "Nom FR"
    # --------------------------------------------------------------------------
    # Cette colonne sert de clé unique pour éviter les doublons
    # --------------------------------------------------------------------------
    $NomFRIndex = -1
    for ($c = 0; $c -lt $Headers.Count; $c++) {
        if ($Headers[$c] -eq "Nom FR") {
            $NomFRIndex = $c + 1
            break
        }
    }

    if ($NomFRIndex -eq -1) {
        Write-Host " Colonne 'Nom FR' non trouvée" -ForegroundColor DarkYellow
        continue
    }

    # --------------------------------------------------------------------------
    # COLLECTE DES EN-TETES UNIQUES
    # --------------------------------------------------------------------------
    # Ajoute les nouveaux en-têtes à la liste globale
    # --------------------------------------------------------------------------
    foreach ($h in $Headers) {
        if ($h -and $AllHeaders -notcontains $h) { [void]$AllHeaders.Add($h) }
    }

    # --------------------------------------------------------------------------
    # EXTRACTION DES DONNEES (lignes 2 à N)
    # --------------------------------------------------------------------------
    # Stocke chaque ligne dans $AllData avec "Nom FR" comme clé
    # Les doublons (même "Nom FR") sont ignorés
    # --------------------------------------------------------------------------
    $added = 0
    for ($row = 2; $row -le $RowCount; $row++) {
        $key = $Data[$row, $NomFRIndex]
        if ([string]::IsNullOrEmpty($key)) { continue }

        if (-not $AllData.ContainsKey($key)) {
            # Créer un dictionnaire ordonné pour cette ligne
            $rowData = [ordered]@{}
            for ($c = 1; $c -le $ColCount; $c++) {
                if ($Headers[$c-1]) {
                    $rowData[$Headers[$c-1]] = $Data[$row, $c]
                }
            }
            $AllData[$key] = $rowData
            $added++
        } else {
            $doublons++
        }
        $totalLignes++
    }

    Write-Host " $($RowCount - 1) lignes, $added ajoutées" -ForegroundColor Green
}

# Vérification qu'il y a des données à écrire
if ($AllData.Count -eq 0) {
    Write-Host "`nAucune donnée!" -ForegroundColor Red
    exit
}

# ==============================================================================
# ECRITURE DANS LE FICHIER EXISTANT
# ==============================================================================

# ------------------------------------------------------------------------------
# INITIALISATION D'EXCEL POUR L'ECRITURE
# ------------------------------------------------------------------------------
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$Excel.DisplayAlerts = $false
$Excel.ScreenUpdating = $false

try {
    # --------------------------------------------------------------------------
    # OUVERTURE DU FICHIER EXISTANT
    # --------------------------------------------------------------------------
    $WB = $Excel.Workbooks.Open($Output)

    # Supprimer les feuilles supplémentaires (garder une seule)
    while ($WB.Worksheets.Count -gt 1) {
        $WB.Worksheets.Item($WB.Worksheets.Count).Delete()
    }

    $Sheet = $WB.Worksheets.Item(1)
    $Sheet.Name = "Objets de gestion"

    # --------------------------------------------------------------------------
    # PREPARATION DES DONNEES EN BLOC
    # --------------------------------------------------------------------------
    # Création d'un tableau 2D pour écriture en une seule opération
    # (beaucoup plus rapide que cellule par cellule)
    # --------------------------------------------------------------------------
    $headerCount = $AllHeaders.Count
    $dataCount = $AllData.Count
    $OutputData = New-Object 'object[,]' ($dataCount + 1), $headerCount

    # Écrire les en-têtes (ligne 0 du tableau)
    for ($c = 0; $c -lt $headerCount; $c++) {
        $OutputData[0, $c] = $AllHeaders[$c]
    }

    # --------------------------------------------------------------------------
    # ECRITURE DES DONNEES TRIEES PAR "Nom FR"
    # --------------------------------------------------------------------------
    $SortedKeys = $AllData.Keys | Sort-Object
    $row = 1
    foreach ($key in $SortedKeys) {
        $obj = $AllData[$key]
        for ($c = 0; $c -lt $headerCount; $c++) {
            $h = $AllHeaders[$c]
            if ($obj.Contains($h)) {
                $OutputData[$row, $c] = $obj[$h]
            }
        }
        $row++
    }

    # --------------------------------------------------------------------------
    # ECRITURE EN BLOC DANS EXCEL
    # --------------------------------------------------------------------------
    $lastCol = [char]([int][char]'A' + [Math]::Min($headerCount - 1, 25))
    $Range = $Sheet.Range("A1:$($lastCol)$($dataCount + 1)")
    $Range.Value2 = $OutputData

    # ==========================================================================
    # APPLICATION DU FORMATAGE
    # ==========================================================================
    Write-Host "Application du formatage..." -ForegroundColor Yellow

    # Formatage de l'en-tête
    $Sheet.Range("A1:$($lastCol)1").Font.Bold = $true
    $Sheet.Range("A1:$($lastCol)1").Interior.ColorIndex = 15

    # Format dd/mm/yyyy pour la première colonne (dates)
    $Sheet.Range("A2:A$($dataCount + 1)").NumberFormat = "dd/mm/yyyy"

    # AutoFit des colonnes
    $Sheet.UsedRange.Columns.AutoFit() | Out-Null

    # ==========================================================================
    # CREATION DE L'ONGLET METADATA
    # ==========================================================================
    # Onglet contenant les statistiques de la fusion
    # ==========================================================================
    $SheetM = $WB.Worksheets.Add([System.Reflection.Missing]::Value, $Sheet)
    $SheetM.Name = "metadata"
    $SheetM.Cells.Item(1,1) = "Info"; $SheetM.Cells.Item(1,2) = "Details"
    $SheetM.Cells.Item(2,1) = "Fichiers"; $SheetM.Cells.Item(2,2) = $Sources.Count
    $SheetM.Cells.Item(3,1) = "Lignes lues"; $SheetM.Cells.Item(3,2) = $totalLignes
    $SheetM.Cells.Item(4,1) = "Doublons"; $SheetM.Cells.Item(4,2) = $doublons
    $SheetM.Cells.Item(5,1) = "Uniques"; $SheetM.Cells.Item(5,2) = $AllData.Count
    $SheetM.Cells.Item(6,1) = "Date"; $SheetM.Cells.Item(6,2) = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $SheetM.Range("A1:B1").Font.Bold = $true
    $SheetM.UsedRange.Columns.AutoFit() | Out-Null

    # --------------------------------------------------------------------------
    # SAUVEGARDE DU FICHIER
    # --------------------------------------------------------------------------
    # 51 = xlOpenXMLWorkbook (.xlsx)
    # --------------------------------------------------------------------------
    $WB.SaveAs($Output, 51)
    $WB.Close($false)

    Write-Host "Fichier mis à jour!" -ForegroundColor Green
}
finally {
    # --------------------------------------------------------------------------
    # LIBERATION DES RESSOURCES COM
    # --------------------------------------------------------------------------
    try { $Excel.Quit() } catch { }
    try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null } catch { }
    [System.GC]::Collect()
}

# ==============================================================================
# RESUME FINAL
# ==============================================================================
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "FUSION TERMINEE" -ForegroundColor Cyan
Write-Host "Fichiers: $($Sources.Count) | Lignes: $totalLignes | Doublons: $doublons | Uniques: $($AllData.Count)" -ForegroundColor White
Write-Host "========================================" -ForegroundColor Cyan

# ==============================================================================
# OUVERTURE AUTOMATIQUE DU FICHIER
# ==============================================================================
Start-Process $Output
