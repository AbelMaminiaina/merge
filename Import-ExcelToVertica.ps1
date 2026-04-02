<#
.SYNOPSIS
    Exporte les donnees de l'onglet "Objets de gestion" de plusieurs fichiers Excel
    vers un script SQL pour Vertica, puis l'execute via DbVisualizer.

.PARAMETER FolderPath
    Chemin du dossier contenant les fichiers Excel.

.PARAMETER OutputPath
    Chemin du dossier de sortie pour le fichier SQL.

.EXAMPLE
    .\Import-ExcelToVertica.ps1 -FolderPath "C:\Users\amami\GitHub\merge\test"
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$FolderPath,

    [Parameter(Mandatory=$false)]
    [string]$OutputPath = "C:\Users\amami\GitHub\merge\output"
)

# Configuration
$TableName       = "ref_obj_gestion"
$SheetName       = "Objets de gestion"
$SqlFileName     = "import_vertica.sql"
$DbVisCmdPath    = "C:\Program Files\DbVisualizer\dbviscmd.bat"
$DbVisConnection = "vertica-NI"

# Creer le dossier de sortie si necessaire
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

function Escape-SqlString {
    param([string]$Value)
    if ([string]::IsNullOrEmpty($Value)) { return "NULL" }
    $escaped = $Value -replace "'", "''"
    return "'$escaped'"
}

function Format-VerticaDate {
    param($DateValue)
    if ($null -eq $DateValue) { return "NULL" }
    return "'" + $DateValue.ToString("yyyy-MM-dd HH:mm:ss") + "'"
}

function Read-ExcelData {
    param(
        [string]$FilePath,
        [string]$TargetSheetName
    )

    $data     = [System.Collections.ArrayList]::new()
    $excel    = $null
    $workbook = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible        = $false
        $excel.DisplayAlerts  = $false
        $excel.Interactive    = $false
        $excel.ScreenUpdating = $false

        $workbook = $excel.Workbooks.Open($FilePath, $false, $true)
        $fileName = [System.IO.Path]::GetFileName($FilePath)

        $targetSheet = $null
        foreach ($sheet in $workbook.Worksheets) {
            if ($sheet.Name -eq $TargetSheetName -or $sheet.Name -match 'objets.*gestion') {
                $targetSheet = $sheet
                break
            }
        }

        if ($null -eq $targetSheet) {
            Write-Host "  Onglet '$TargetSheetName' non trouve dans $fileName" -ForegroundColor Yellow
            return $data
        }

        $usedRange = $targetSheet.UsedRange
        $lastRow   = $usedRange.Rows.Count
        $lastCol   = [Math]::Min($usedRange.Columns.Count, 6)

        Write-Host "  Lecture de $($lastRow - 1) lignes depuis '$($targetSheet.Name)'..." -ForegroundColor Gray

        $range     = $targetSheet.Range($targetSheet.Cells(2, 1), $targetSheet.Cells($lastRow, $lastCol))
        $allValues = $range.Value2

        for ($row = 1; $row -le ($lastRow - 1); $row++) {
            $dateValue       = $allValues[$row, 1]
            $usedValue       = $allValues[$row, 2]
            $nomFrValue      = $allValues[$row, 3]
            $definitionValue = $allValues[$row, 4]
            $nomEnValue      = $allValues[$row, 5]
            $trigrammeValue  = $allValues[$row, 6]

            $nomFrStr      = if ($null -ne $nomFrValue)      { $nomFrValue.ToString() }      else { "" }
            $trigrammeStr  = if ($null -ne $trigrammeValue)  { $trigrammeValue.ToString() }  else { "" }
            $definitionStr = if ($null -ne $definitionValue) { $definitionValue.ToString() } else { "" }

            if ([string]::IsNullOrWhiteSpace($nomFrStr) -and
                [string]::IsNullOrWhiteSpace($trigrammeStr) -and
                [string]::IsNullOrWhiteSpace($definitionStr)) {
                continue
            }

            $parsedDate = $null
            if ($null -ne $dateValue) {
                if ($dateValue -is [double]) {
                    try { $parsedDate = [DateTime]::FromOADate($dateValue) } catch { $parsedDate = $null }
                } elseif ($dateValue -is [string] -and -not [string]::IsNullOrWhiteSpace($dateValue)) {
                    try {
                        $parsedDate = [DateTime]::ParseExact($dateValue, "dd/MM/yyyy", [System.Globalization.CultureInfo]::InvariantCulture)
                    } catch {
                        try { $parsedDate = [DateTime]::Parse($dateValue) } catch { $parsedDate = $null }
                    }
                }
            }

            [void]$data.Add([PSCustomObject]@{
                Date            = $parsedDate
                Used            = if ($null -ne $usedValue) { $usedValue.ToString() } else { "" }
                NomFr           = $nomFrStr
                Definition      = $definitionStr
                NomEn           = if ($null -ne $nomEnValue) { $nomEnValue.ToString() } else { "" }
                Trigramme       = $trigrammeStr
                NomFichierExcel = $fileName
            })
        }

        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($range) | Out-Null

    } catch {
        Write-Host "  Erreur lors de la lecture de $FilePath : $_" -ForegroundColor Red
    } finally {
        try {
            if ($null -ne $workbook) {
                $workbook.Close($false)
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
            }
        } catch { }
        try {
            if ($null -ne $excel) {
                $excel.Quit()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            }
        } catch { }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }

    return $data
}

function Export-ToSql {
    param(
        [array]$Data,
        [string]$FilePath
    )

    $sqlContent = [System.Text.StringBuilder]::new()

    [void]$sqlContent.AppendLine("-- Script d'import pour Vertica")
    [void]$sqlContent.AppendLine("-- Genere le $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')")
    [void]$sqlContent.AppendLine("-- Total: $($Data.Count) enregistrements")
    [void]$sqlContent.AppendLine("")
    [void]$sqlContent.AppendLine("-- Creation de la table (si elle n'existe pas)")
    [void]$sqlContent.AppendLine("CREATE TABLE IF NOT EXISTS $TableName (")
    [void]$sqlContent.AppendLine("    Id IDENTITY(1,1),")
    [void]$sqlContent.AppendLine("    Date TIMESTAMP NULL,")
    [void]$sqlContent.AppendLine("    Used VARCHAR(255),")
    [void]$sqlContent.AppendLine("    NomFr VARCHAR(500),")
    [void]$sqlContent.AppendLine("    Definition VARCHAR(65000),")
    [void]$sqlContent.AppendLine("    NomEn VARCHAR(500),")
    [void]$sqlContent.AppendLine("    Trigramme VARCHAR(100),")
    [void]$sqlContent.AppendLine("    NomFichierExcel VARCHAR(500),")
    [void]$sqlContent.AppendLine("    DateImport TIMESTAMP DEFAULT CURRENT_TIMESTAMP")
    [void]$sqlContent.AppendLine(");")
    [void]$sqlContent.AppendLine("")
    [void]$sqlContent.AppendLine("-- Vider la table avant insertion")
    [void]$sqlContent.AppendLine("TRUNCATE TABLE $TableName;")
    [void]$sqlContent.AppendLine("")
    [void]$sqlContent.AppendLine("-- Insertion des donnees")

    $batchSize = 100
    $count     = 0

    foreach ($item in $Data) {
        $dateVal  = Format-VerticaDate -DateValue $item.Date
        $usedVal  = Escape-SqlString   -Value $item.Used
        $nomFrVal = Escape-SqlString   -Value $item.NomFr
        $defVal   = Escape-SqlString   -Value $item.Definition
        $nomEnVal = Escape-SqlString   -Value $item.NomEn
        $triVal   = Escape-SqlString   -Value $item.Trigramme
        $fileVal  = Escape-SqlString   -Value $item.NomFichierExcel

        [void]$sqlContent.AppendLine("INSERT INTO $TableName (Date, Used, NomFr, Definition, NomEn, Trigramme, NomFichierExcel)")
        [void]$sqlContent.AppendLine("VALUES ($dateVal, $usedVal, $nomFrVal, $defVal, $nomEnVal, $triVal, $fileVal);")

        $count++
        if ($count % $batchSize -eq 0) {
            [void]$sqlContent.AppendLine("")
            [void]$sqlContent.AppendLine("-- $count enregistrements inseres...")
            [void]$sqlContent.AppendLine("COMMIT;")
            [void]$sqlContent.AppendLine("")
        }
    }

    [void]$sqlContent.AppendLine("")
    [void]$sqlContent.AppendLine("COMMIT;")
    [void]$sqlContent.AppendLine("")
    [void]$sqlContent.AppendLine("-- Verification")
    [void]$sqlContent.AppendLine("SELECT COUNT(*) AS total_records FROM $TableName;")

    [System.IO.File]::WriteAllText($FilePath, $sqlContent.ToString(), [System.Text.Encoding]::UTF8)
}

function Invoke-DbVisCmd {
    param([string]$SqlFilePath)

    Write-Host ""
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host "  Execution via DbVisualizer (dbviscmd)" -ForegroundColor Cyan
    Write-Host "============================================" -ForegroundColor Cyan

    if (-not (Test-Path $DbVisCmdPath)) {
        Write-Host "dbviscmd.bat non trouve: $DbVisCmdPath" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Pour importer manuellement dans Vertica:" -ForegroundColor Yellow
        Write-Host "  1. Ouvrez DbVisualizer" -ForegroundColor White
        Write-Host "  2. Connectez-vous a $DbVisConnection" -ForegroundColor White
        Write-Host "  3. Ouvrez le fichier: $SqlFilePath" -ForegroundColor White
        Write-Host "  4. Executez le script (F5 ou Ctrl+Enter)" -ForegroundColor White
        return
    }

    Write-Host "Connexion  : $DbVisConnection" -ForegroundColor White
    Write-Host "Script SQL : $SqlFilePath" -ForegroundColor White
    Write-Host ""
    Write-Host "Execution en cours..." -ForegroundColor Yellow

    & $DbVisCmdPath -connection $DbVisConnection -script $SqlFilePath

    if ($LASTEXITCODE -eq 0) {
        Write-Host ""
        Write-Host "OK - Script SQL execute avec succes." -ForegroundColor Green
    } else {
        Write-Host ""
        Write-Host "ERREUR - dbviscmd a retourne le code : $LASTEXITCODE" -ForegroundColor Red
        Write-Host "Verifiez la connexion '$DbVisConnection' dans DbVisualizer." -ForegroundColor Yellow
    }
}

# ============================================
# SCRIPT PRINCIPAL
# ============================================

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  Export Excel vers Vertica (SQL)" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

if (-not (Test-Path $FolderPath)) {
    Write-Host "Erreur: Le dossier '$FolderPath' n'existe pas." -ForegroundColor Red
    exit 1
}

$excelFiles = Get-ChildItem -Path $FolderPath -Filter "*.xlsx" -File
if ($excelFiles.Count -eq 0) {
    Write-Host "Aucun fichier Excel trouve dans '$FolderPath'." -ForegroundColor Yellow
    exit 0
}

Write-Host "Fichiers Excel trouves : $($excelFiles.Count)" -ForegroundColor White
Write-Host "Dossier de sortie      : $OutputPath" -ForegroundColor White
Write-Host ""

$allData        = [System.Collections.ArrayList]::new()
$processedFiles = 0

foreach ($file in $excelFiles) {
    Write-Host ""
    Write-Host "Traitement: $($file.Name)" -ForegroundColor White

    $fileData = Read-ExcelData -FilePath $file.FullName -TargetSheetName $SheetName

    if ($fileData.Count -gt 0) {
        [void]$allData.AddRange($fileData)
        $processedFiles++
        Write-Host "  -> $($fileData.Count) enregistrements extraits" -ForegroundColor Green
    }
}

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "RESUME" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "Fichiers traites     : $processedFiles / $($excelFiles.Count)" -ForegroundColor White
Write-Host "Total enregistrements: $($allData.Count)" -ForegroundColor White
Write-Host ""

if ($allData.Count -gt 0) {
    $sqlPath = Join-Path $OutputPath $SqlFileName
    Write-Host "Generation du script SQL..." -ForegroundColor Cyan
    Export-ToSql -Data $allData -FilePath $sqlPath
    Write-Host "  -> $sqlPath" -ForegroundColor Green

    Invoke-DbVisCmd -SqlFilePath $sqlPath
}

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  Termine!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""