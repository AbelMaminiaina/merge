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
    [string]$Folder = "\\uf11d001\EIDDIJON\App\IBIQ\03-Maintenances\020-Evolutions\RU555496-POD_Workflow-Nomenclature\09-Livrable\Fichiers outils",
    [string]$Output = "\\uf11d001\EIDDIJON\App\IBIQ\03-Maintenances\020-Evolutions\RU555496-POD_Workflow-Nomenclature\09-Livrable\Fichiers outils\Merge-IBIX_XXXX_Outil_Onglet_Objets de gestion.xlsx"
)

Write-Host "=== FUSION DES ONGLETS OBJETS DE GESTION ===" -ForegroundColor Cyan

# ------------------------------------------------------------------------------
# FERMETURE DES INSTANCES EXCEL
# ------------------------------------------------------------------------------
# Ferme toutes les instances Excel en cours pour éviter les conflits COM
# Attente de 3 secondes pour laisser le temps aux processus de se terminer
# ------------------------------------------------------------------------------
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 3

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
# $SourceFormatting: Stocke le formatage du premier fichier pour le reproduire
# ------------------------------------------------------------------------------
$AllData = @{}
$AllHeaders = [System.Collections.ArrayList]@()
$totalLignes = 0
$doublons = 0
$SourceFormatting = $null

# ==============================================================================
# FONCTION : Read-ExcelFile
# ==============================================================================
# Description : Lit un fichier Excel et extrait les données de l'onglet
#               "Objets de gestion"
# Paramètres  :
#   - $FilePath         : Chemin du fichier Excel à lire
#   - $CaptureFormatting: Si activé, capture aussi le formatage des cellules
# Retourne    : Hashtable avec Data, Rows, Cols et optionnellement Formatting
# ==============================================================================
function Read-ExcelFile {
    param($FilePath, [switch]$CaptureFormatting)

    $Excel = $null
    $result = $null
    Write-Host "Lecture de $FilePath..." -ForegroundColor DarkYellow
    try {
        # ------------------------------------------------------------------
        # INITIALISATION DE L'APPLICATION EXCEL
        # ------------------------------------------------------------------
        # Création d'une instance Excel invisible pour la lecture
        # ------------------------------------------------------------------
        $Excel = New-Object -ComObject Excel.Application
        $Excel.Visible = $false
        $Excel.DisplayAlerts = $false
        $Excel.ScreenUpdating = $false

        # Ouvrir le fichier en lecture seule (paramètre $true)
        $Workbook = $Excel.Workbooks.Open($FilePath, $false, $true)

        # ------------------------------------------------------------------
        # RECHERCHE DE L'ONGLET "Objets de gestion"
        # ------------------------------------------------------------------
        $Sheet = $null
        foreach ($ws in $Workbook.Worksheets) {
            if ($ws.Name -eq "Objets de gestion") {
                $Sheet = $ws
                break
            }
        }

        if ($Sheet) {
            # --------------------------------------------------------------
            # EXTRACTION DES DONNEES
            # --------------------------------------------------------------
            # Value2 retourne les valeurs brutes (dates = nombres sériels)
            # --------------------------------------------------------------
            $UsedRange = $Sheet.UsedRange
            $ColCount = $UsedRange.Columns.Count

            $result = @{
                Data = $UsedRange.Value2
                Rows = $UsedRange.Rows.Count
                Cols = $ColCount
            }

            if ($CaptureFormatting) {
                $result.Formatting = @{
                    HeaderRowHeight = $Sheet.Rows.Item(1).RowHeight
                    HasAutoFilter = $Sheet.AutoFilterMode
                }
            }
        }

        $Workbook.Close($false)
    }
    catch {
        Write-Host "Erreur: $_" -ForegroundColor Red
    }
    finally {
        # ------------------------------------------------------------------
        # LIBERATION DES RESSOURCES COM
        # ------------------------------------------------------------------
        # Important pour éviter les fuites mémoire et processus fantômes
        # ------------------------------------------------------------------
        if ($Excel) {
            try { $Excel.Quit() } catch { }
            try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null } catch { }
        }
        [System.GC]::Collect()
        Start-Sleep -Milliseconds 500
    }

    return $result
}

# ==============================================================================
# BOUCLE DE LECTURE DES FICHIERS SOURCES
# ==============================================================================
# Parcourt chaque fichier Excel et extrait les données de l'onglet
# Le premier fichier sert de référence pour le formatage
# ==============================================================================
$isFirst = $true
foreach ($sourceFile in $Sources) {
    $fileName = Split-Path $sourceFile -Leaf
    Write-Host "Lecture de $fileName..." -ForegroundColor Yellow -NoNewline

    # --------------------------------------------------------------------------
    # LECTURE AVEC OU SANS CAPTURE DU FORMATAGE
    # --------------------------------------------------------------------------
    # Premier fichier : capture le formatage pour le reproduire dans la fusion
    # Fichiers suivants : lecture des données uniquement
    # --------------------------------------------------------------------------
    if ($isFirst) {
        $result = Read-ExcelFile -FilePath $sourceFile -CaptureFormatting
        if ($result -and $result.Formatting) {
            $SourceFormatting = $result.Formatting
            Write-Host " (formatage capturé)" -ForegroundColor Magenta -NoNewline
        }
        $isFirst = $false
    } else {
        $result = Read-ExcelFile -FilePath $sourceFile
    }

    if ($null -eq $result) {
        Write-Host " Onglet non trouvé" -ForegroundColor DarkYellow
        continue
    }
    Write-Host "Lecture de 2..." -ForegroundColor Yellow -NoNewline
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
    # APPLICATION DU FORMATAGE SOURCE
    # ==========================================================================
    # Reproduit exactement le formatage du premier fichier lu
    # ==========================================================================
    if ($SourceFormatting) {
        Write-Host "Application du formatage source..." -ForegroundColor Yellow
        Write-Host "Formats capturés:" -ForegroundColor Cyan

        # ----------------------------------------------------------------------
        # HAUTEUR DE LIGNE DES EN-TETES
        # ----------------------------------------------------------------------
        if ($SourceFormatting.HeaderRowHeight) {
            $Sheet.Rows.Item(1).RowHeight = $SourceFormatting.HeaderRowHeight
        }
        # ----------------------------------------------------------------------
        # ACTIVATION DE L'AUTOFILTER SI PRESENT DANS LA SOURCE
        # ----------------------------------------------------------------------
        if ($SourceFormatting.HasAutoFilter) {
            $Sheet.Range("A1:$($lastCol)1").AutoFilter() | Out-Null
        }
    } else {
        # ----------------------------------------------------------------------
        # FORMATAGE PAR DEFAUT (si pas de source)
        # ----------------------------------------------------------------------
        $Sheet.Range("A1:$($lastCol)1").Font.Bold = $true
        $Sheet.Range("A1:$($lastCol)1").Interior.ColorIndex = 15
        $Sheet.UsedRange.Columns.AutoFit() | Out-Null
    }

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
