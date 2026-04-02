<#
.SYNOPSIS
    Exporte les donnees de l'onglet "Objets de gestion" de plusieurs fichiers Excel
    vers un script SQL, puis l'execute sur SQL Server Express.

.PARAMETER FolderPath
    Chemin du dossier contenant les fichiers Excel.

.PARAMETER OutputPath
    Chemin du dossier de sortie pour le fichier SQL.

.PARAMETER ServerInstance
    Instance SQL Server (ex: localhost\SQLEXPRESS, .\SQLEXPRESS02)

.PARAMETER DatabaseName
    Nom de la base de donnees.

.EXAMPLE
    .\Import-ExcelToSqlServer.ps1 -FolderPath "C:\Users\amami\GitHub\merge\test"
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$FolderPath,

    [Parameter(Mandatory=$false)]
    [string]$OutputPath = "C:\Users\amami\GitHub\merge\output",

    [Parameter(Mandatory=$false)]
    [string]$ServerInstance = ".\SQLEXPRESS02",

    [Parameter(Mandatory=$false)]
    [string]$DatabaseName = "MergeDB"
)

# Configuration
$TableName   = "ref_obj_gestion"
$SheetName   = "Objets de gestion"
$SqlFileName = "import_sqlserver.sql"

# Creer le dossier de sortie si necessaire
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

function Escape-SqlString {
    param([string]$Value)
    if ([string]::IsNullOrEmpty($Value)) { return "NULL" }
    $escaped = $Value -replace "'", "''"
    return "N'$escaped'"
}

function Format-SqlServerDate {
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
        [string]$FilePath,
        [string]$Database
    )

    $sqlContent = [System.Text.StringBuilder]::new()

    [void]$sqlContent.AppendLine("-- Script d'import pour SQL Server")
    [void]$sqlContent.AppendLine("-- Genere le $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')")
    [void]$sqlContent.AppendLine("-- Total: $($Data.Count) enregistrements")
    [void]$sqlContent.AppendLine("")
    [void]$sqlContent.AppendLine("USE [$Database];")
    [void]$sqlContent.AppendLine("GO")
    [void]$sqlContent.AppendLine("")
    [void]$sqlContent.AppendLine("-- Creation de la table (si elle n'existe pas)")
    [void]$sqlContent.AppendLine("IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = '$TableName')")
    [void]$sqlContent.AppendLine("BEGIN")
    [void]$sqlContent.AppendLine("    CREATE TABLE [$TableName] (")
    [void]$sqlContent.AppendLine("        [Id] INT IDENTITY(1,1) PRIMARY KEY,")
    [void]$sqlContent.AppendLine("        [Date] DATETIME NULL,")
    [void]$sqlContent.AppendLine("        [Used] NVARCHAR(255) NULL,")
    [void]$sqlContent.AppendLine("        [NomFr] NVARCHAR(500) NULL,")
    [void]$sqlContent.AppendLine("        [Definition] NVARCHAR(MAX) NULL,")
    [void]$sqlContent.AppendLine("        [NomEn] NVARCHAR(500) NULL,")
    [void]$sqlContent.AppendLine("        [Trigramme] NVARCHAR(100) NULL,")
    [void]$sqlContent.AppendLine("        [NomFichierExcel] NVARCHAR(500) NULL,")
    [void]$sqlContent.AppendLine("        [DateImport] DATETIME DEFAULT GETDATE()")
    [void]$sqlContent.AppendLine("    );")
    [void]$sqlContent.AppendLine("    CREATE INDEX IX_${TableName}_NomFr ON [$TableName] ([NomFr]);")
    [void]$sqlContent.AppendLine("    CREATE INDEX IX_${TableName}_Trigramme ON [$TableName] ([Trigramme]);")
    [void]$sqlContent.AppendLine("    CREATE INDEX IX_${TableName}_NomFichierExcel ON [$TableName] ([NomFichierExcel]);")
    [void]$sqlContent.AppendLine("END")
    [void]$sqlContent.AppendLine("GO")
    [void]$sqlContent.AppendLine("")
    [void]$sqlContent.AppendLine("-- Vider la table avant insertion")
    [void]$sqlContent.AppendLine("TRUNCATE TABLE [$TableName];")
    [void]$sqlContent.AppendLine("GO")
    [void]$sqlContent.AppendLine("")
    [void]$sqlContent.AppendLine("-- Insertion des donnees")

    $batchSize = 100
    $count     = 0

    foreach ($item in $Data) {
        $dateVal  = Format-SqlServerDate -DateValue $item.Date
        $usedVal  = Escape-SqlString -Value $item.Used
        $nomFrVal = Escape-SqlString -Value $item.NomFr
        $defVal   = Escape-SqlString -Value $item.Definition
        $nomEnVal = Escape-SqlString -Value $item.NomEn
        $triVal   = Escape-SqlString -Value $item.Trigramme
        $fileVal  = Escape-SqlString -Value $item.NomFichierExcel

        [void]$sqlContent.AppendLine("INSERT INTO [$TableName] ([Date], [Used], [NomFr], [Definition], [NomEn], [Trigramme], [NomFichierExcel])")
        [void]$sqlContent.AppendLine("VALUES ($dateVal, $usedVal, $nomFrVal, $defVal, $nomEnVal, $triVal, $fileVal);")

        $count++
        if ($count % $batchSize -eq 0) {
            [void]$sqlContent.AppendLine("GO")
            [void]$sqlContent.AppendLine("-- $count enregistrements inseres...")
            [void]$sqlContent.AppendLine("")
        }
    }

    [void]$sqlContent.AppendLine("GO")
    [void]$sqlContent.AppendLine("")
    [void]$sqlContent.AppendLine("-- Verification")
    [void]$sqlContent.AppendLine("SELECT COUNT(*) AS total_records FROM [$TableName];")
    [void]$sqlContent.AppendLine("GO")

    # UTF-8 sans BOM
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($FilePath, $sqlContent.ToString(), $utf8NoBom)
}

function Invoke-SqlScript {
    param(
        [string]$Server,
        [string]$Database,
        [string]$SqlFilePath
    )

    Write-Host ""
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host "  Execution sur SQL Server" -ForegroundColor Cyan
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host "Serveur  : $Server" -ForegroundColor White
    Write-Host "Database : $Database" -ForegroundColor White
    Write-Host "Script   : $SqlFilePath" -ForegroundColor White
    Write-Host ""

    # Verifier si sqlcmd est disponible
    $sqlcmd = Get-Command "sqlcmd" -ErrorAction SilentlyContinue

    if ($sqlcmd) {
        Write-Host "Execution via sqlcmd (UTF-8)..." -ForegroundColor Yellow
        & sqlcmd -S $Server -d $Database -E -i $SqlFilePath -f 65001 -b
        if ($LASTEXITCODE -eq 0) {
            Write-Host "OK - Script execute avec succes." -ForegroundColor Green
        } else {
            Write-Host "ERREUR - sqlcmd a retourne le code: $LASTEXITCODE" -ForegroundColor Red
        }
    } else {
        Write-Host "sqlcmd non trouve. Execution via PowerShell..." -ForegroundColor Yellow

        try {
            # Creer la base de donnees si elle n'existe pas
            $masterConn = New-Object System.Data.SqlClient.SqlConnection("Server=$Server;Database=master;Integrated Security=True;TrustServerCertificate=True;")
            $masterConn.Open()
            $createDbCmd = New-Object System.Data.SqlClient.SqlCommand("IF NOT EXISTS (SELECT name FROM sys.databases WHERE name = '$Database') CREATE DATABASE [$Database]", $masterConn)
            $createDbCmd.ExecuteNonQuery() | Out-Null
            $masterConn.Close()

            # Lire et executer le script SQL
            $sqlContent = [System.IO.File]::ReadAllText($SqlFilePath)

            # Separer par GO
            $batches = $sqlContent -split '\r?\nGO\r?\n'

            $conn = New-Object System.Data.SqlClient.SqlConnection("Server=$Server;Database=$Database;Integrated Security=True;TrustServerCertificate=True;")
            $conn.Open()

            $batchCount = 0
            foreach ($batch in $batches) {
                $batch = $batch.Trim()
                if (-not [string]::IsNullOrWhiteSpace($batch) -and $batch -ne "GO") {
                    # Ignorer USE statement car on est deja connecte
                    if ($batch -notmatch '^\s*USE\s+\[') {
                        try {
                            $cmd = New-Object System.Data.SqlClient.SqlCommand($batch, $conn)
                            $cmd.CommandTimeout = 300
                            $result = $cmd.ExecuteNonQuery()
                            $batchCount++
                        } catch {
                            Write-Host "Erreur batch $batchCount : $_" -ForegroundColor Red
                        }
                    }
                }
            }

            $conn.Close()
            Write-Host "OK - $batchCount batches executes avec succes." -ForegroundColor Green

        } catch {
            Write-Host "ERREUR: $_" -ForegroundColor Red
        }
    }
}

# ============================================
# SCRIPT PRINCIPAL
# ============================================

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  Export Excel vers SQL Server" -ForegroundColor Cyan
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
Write-Host "Serveur SQL Server     : $ServerInstance" -ForegroundColor White
Write-Host "Base de donnees        : $DatabaseName" -ForegroundColor White
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
    Export-ToSql -Data $allData -FilePath $sqlPath -Database $DatabaseName
    Write-Host "  -> $sqlPath" -ForegroundColor Green

    Invoke-SqlScript -Server $ServerInstance -Database $DatabaseName -SqlFilePath $sqlPath
}

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  Termine!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
