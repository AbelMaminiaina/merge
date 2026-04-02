<#
.SYNOPSIS
    Importe les donnees de l'onglet "Objets de gestion" de plusieurs fichiers Excel vers SQL Server.

.DESCRIPTION
    Ce script lit tous les fichiers Excel d'un dossier, extrait les donnees de l'onglet
    "Objets de gestion" et les insere en masse dans la table ref_obj_gestion de SQL Server.

.PARAMETER FolderPath
    Chemin du dossier contenant les fichiers Excel.

.PARAMETER ServerInstance
    Instance SQL Server (ex: localhost\SQLEXPRESS, .\SQLEXPRESS, localhost)

.PARAMETER DatabaseName
    Nom de la base de donnees.

.EXAMPLE
    .\Import-ExcelToSqlServer.ps1 -FolderPath "C:\Users\amami\GitHub\merge\test" -ServerInstance ".\SQLEXPRESS" -DatabaseName "MergeDB"
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$FolderPath,

    [Parameter(Mandatory=$false)]
    [string]$ServerInstance = ".\SQLEXPRESS02",

    [Parameter(Mandatory=$false)]
    [string]$DatabaseName = "MergeDB"
)

# Configuration
$TableName = "ref_obj_gestion"
$SheetName = "Objets de gestion"
$BatchSize = 1000

# Fonction pour creer la connexion SQL Server
function Get-SqlConnection {
    param([string]$Server, [string]$Database)

    $connectionString = "Server=$Server;Database=$Database;Integrated Security=True;TrustServerCertificate=True;"
    $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
    return $connection
}

# Fonction pour creer la base de donnees si elle n'existe pas
function Initialize-Database {
    param([string]$Server, [string]$Database)

    Write-Host "Verification de la base de donnees '$Database'..." -ForegroundColor Cyan

    # Connexion a master pour creer la base si necessaire
    $masterConnection = New-Object System.Data.SqlClient.SqlConnection("Server=$Server;Database=master;Integrated Security=True;TrustServerCertificate=True;")
    $masterConnection.Open()

    $checkDbQuery = "IF NOT EXISTS (SELECT name FROM sys.databases WHERE name = '$Database') CREATE DATABASE [$Database]"
    $cmd = New-Object System.Data.SqlClient.SqlCommand($checkDbQuery, $masterConnection)
    $cmd.ExecuteNonQuery() | Out-Null
    $masterConnection.Close()

    Write-Host "Base de donnees '$Database' prete." -ForegroundColor Green
}

# Fonction pour creer la table si elle n'existe pas
function Initialize-Table {
    param([System.Data.SqlClient.SqlConnection]$Connection)

    Write-Host "Verification de la table '$TableName'..." -ForegroundColor Cyan

    $createTableQuery = @"
    IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = '$TableName')
    BEGIN
        CREATE TABLE [$TableName] (
            [Id] INT IDENTITY(1,1) PRIMARY KEY,
            [Date] DATETIME NULL,
            [Used] NVARCHAR(255) NULL,
            [NomFr] NVARCHAR(500) NULL,
            [Definition] NVARCHAR(MAX) NULL,
            [NomEn] NVARCHAR(500) NULL,
            [Trigramme] NVARCHAR(100) NULL,
            [NomFichierExcel] NVARCHAR(500) NULL,
            [DateImport] DATETIME DEFAULT GETDATE()
        )

        CREATE INDEX IX_${TableName}_NomFr ON [$TableName] ([NomFr])
        CREATE INDEX IX_${TableName}_Trigramme ON [$TableName] ([Trigramme])
        CREATE INDEX IX_${TableName}_NomFichierExcel ON [$TableName] ([NomFichierExcel])
    END
"@

    $cmd = New-Object System.Data.SqlClient.SqlCommand($createTableQuery, $Connection)
    $cmd.ExecuteNonQuery() | Out-Null

    Write-Host "Table '$TableName' prete." -ForegroundColor Green
}

# Fonction pour vider la table avant insertion
function Clear-Table {
    param([System.Data.SqlClient.SqlConnection]$Connection)

    Write-Host "Vidage de la table '$TableName'..." -ForegroundColor Yellow

    $truncateQuery = "TRUNCATE TABLE [$TableName]"
    $cmd = New-Object System.Data.SqlClient.SqlCommand($truncateQuery, $Connection)
    $cmd.ExecuteNonQuery() | Out-Null

    Write-Host "Table '$TableName' videe." -ForegroundColor Green
}

# Fonction pour lire les donnees Excel (lecture en bloc pour performance)
function Read-ExcelData {
    param(
        [string]$FilePath,
        [string]$TargetSheetName
    )

    $data = [System.Collections.ArrayList]::new()
    $excel = $null
    $workbook = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.Interactive = $false
        $excel.ScreenUpdating = $false

        $workbook = $excel.Workbooks.Open($FilePath, $false, $true)  # ReadOnly = true
        $fileName = [System.IO.Path]::GetFileName($FilePath)

        # Chercher l'onglet "Objets de gestion"
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
        $lastRow = $usedRange.Rows.Count
        $lastCol = [Math]::Min($usedRange.Columns.Count, 6)

        Write-Host "  Lecture de $($lastRow - 1) lignes depuis '$($targetSheet.Name)'..." -ForegroundColor Gray

        # Lecture en bloc - BEAUCOUP plus rapide que cellule par cellule
        $range = $targetSheet.Range($targetSheet.Cells(2, 1), $targetSheet.Cells($lastRow, $lastCol))
        $allValues = $range.Value2

        # Traiter les donnees en memoire
        for ($row = 1; $row -le ($lastRow - 1); $row++) {
            $dateValue = $allValues[$row, 1]
            $usedValue = $allValues[$row, 2]
            $nomFrValue = $allValues[$row, 3]
            $definitionValue = $allValues[$row, 4]
            $nomEnValue = $allValues[$row, 5]
            $trigrammeValue = $allValues[$row, 6]

            # Convertir en string
            $nomFrStr = if ($null -ne $nomFrValue) { $nomFrValue.ToString() } else { "" }
            $trigrammeStr = if ($null -ne $trigrammeValue) { $trigrammeValue.ToString() } else { "" }
            $definitionStr = if ($null -ne $definitionValue) { $definitionValue.ToString() } else { "" }

            # Ignorer les lignes completement vides
            if ([string]::IsNullOrWhiteSpace($nomFrStr) -and
                [string]::IsNullOrWhiteSpace($trigrammeStr) -and
                [string]::IsNullOrWhiteSpace($definitionStr)) {
                continue
            }

            # Parser la date (Excel stocke les dates comme nombres)
            $parsedDate = $null
            if ($null -ne $dateValue) {
                if ($dateValue -is [double]) {
                    try {
                        $parsedDate = [DateTime]::FromOADate($dateValue)
                    } catch {
                        $parsedDate = $null
                    }
                } elseif ($dateValue -is [string] -and -not [string]::IsNullOrWhiteSpace($dateValue)) {
                    try {
                        $parsedDate = [DateTime]::ParseExact($dateValue, "dd/MM/yyyy", [System.Globalization.CultureInfo]::InvariantCulture)
                    } catch {
                        try {
                            $parsedDate = [DateTime]::Parse($dateValue)
                        } catch {
                            $parsedDate = $null
                        }
                    }
                }
            }

            [void]$data.Add([PSCustomObject]@{
                Date = $parsedDate
                Used = if ($null -ne $usedValue) { $usedValue.ToString() } else { "" }
                NomFr = $nomFrStr
                Definition = $definitionStr
                NomEn = if ($null -ne $nomEnValue) { $nomEnValue.ToString() } else { "" }
                Trigramme = $trigrammeStr
                NomFichierExcel = $fileName
            })
        }

        # Liberer la memoire du range
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

# Fonction pour inserer les donnees en masse avec SqlBulkCopy
function Import-DataToSqlServer {
    param(
        [System.Data.SqlClient.SqlConnection]$Connection,
        [array]$Data
    )

    if ($Data.Count -eq 0) {
        Write-Host "Aucune donnee a inserer." -ForegroundColor Yellow
        return 0
    }

    Write-Host "Insertion de $($Data.Count) enregistrements..." -ForegroundColor Cyan

    # Creer une DataTable
    $dataTable = New-Object System.Data.DataTable
    $dataTable.Columns.Add("Date", [DateTime]) | Out-Null
    $dataTable.Columns.Add("Used", [string]) | Out-Null
    $dataTable.Columns.Add("NomFr", [string]) | Out-Null
    $dataTable.Columns.Add("Definition", [string]) | Out-Null
    $dataTable.Columns.Add("NomEn", [string]) | Out-Null
    $dataTable.Columns.Add("Trigramme", [string]) | Out-Null
    $dataTable.Columns.Add("NomFichierExcel", [string]) | Out-Null

    # Permettre les valeurs NULL pour la colonne Date
    $dataTable.Columns["Date"].AllowDBNull = $true

    # Remplir la DataTable
    foreach ($item in $Data) {
        $row = $dataTable.NewRow()
        if ($null -ne $item.Date) {
            $row["Date"] = $item.Date
        } else {
            $row["Date"] = [DBNull]::Value
        }
        $row["Used"] = if ([string]::IsNullOrEmpty($item.Used)) { [DBNull]::Value } else { $item.Used }
        $row["NomFr"] = if ([string]::IsNullOrEmpty($item.NomFr)) { [DBNull]::Value } else { $item.NomFr }
        $row["Definition"] = if ([string]::IsNullOrEmpty($item.Definition)) { [DBNull]::Value } else { $item.Definition }
        $row["NomEn"] = if ([string]::IsNullOrEmpty($item.NomEn)) { [DBNull]::Value } else { $item.NomEn }
        $row["Trigramme"] = if ([string]::IsNullOrEmpty($item.Trigramme)) { [DBNull]::Value } else { $item.Trigramme }
        $row["NomFichierExcel"] = $item.NomFichierExcel
        $dataTable.Rows.Add($row)
    }

    # Utiliser SqlBulkCopy pour l'insertion en masse
    $bulkCopy = New-Object System.Data.SqlClient.SqlBulkCopy($Connection)
    $bulkCopy.DestinationTableName = $TableName
    $bulkCopy.BatchSize = $BatchSize
    $bulkCopy.BulkCopyTimeout = 600  # 10 minutes

    # Mapper les colonnes
    $bulkCopy.ColumnMappings.Add("Date", "Date") | Out-Null
    $bulkCopy.ColumnMappings.Add("Used", "Used") | Out-Null
    $bulkCopy.ColumnMappings.Add("NomFr", "NomFr") | Out-Null
    $bulkCopy.ColumnMappings.Add("Definition", "Definition") | Out-Null
    $bulkCopy.ColumnMappings.Add("NomEn", "NomEn") | Out-Null
    $bulkCopy.ColumnMappings.Add("Trigramme", "Trigramme") | Out-Null
    $bulkCopy.ColumnMappings.Add("NomFichierExcel", "NomFichierExcel") | Out-Null

    try {
        $bulkCopy.WriteToServer($dataTable)
        Write-Host "Insertion terminee avec succes!" -ForegroundColor Green
        return $Data.Count
    } catch {
        Write-Host "Erreur lors de l'insertion : $_" -ForegroundColor Red
        return 0
    } finally {
        $bulkCopy.Close()
        $dataTable.Dispose()
    }
}

# ============================================
# SCRIPT PRINCIPAL
# ============================================

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  Import Excel vers SQL Server" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

# Verifier que le dossier existe
if (-not (Test-Path $FolderPath)) {
    Write-Host "Erreur: Le dossier '$FolderPath' n'existe pas." -ForegroundColor Red
    exit 1
}

# Recuperer tous les fichiers Excel
$excelFiles = Get-ChildItem -Path $FolderPath -Filter "*.xlsx" -File
if ($excelFiles.Count -eq 0) {
    Write-Host "Aucun fichier Excel trouve dans '$FolderPath'." -ForegroundColor Yellow
    exit 0
}

Write-Host "Fichiers Excel trouves: $($excelFiles.Count)" -ForegroundColor White
Write-Host "Serveur: $ServerInstance" -ForegroundColor White
Write-Host "Base de donnees: $DatabaseName" -ForegroundColor White
Write-Host ""

# Initialiser la base de donnees et la table
try {
    Initialize-Database -Server $ServerInstance -Database $DatabaseName

    $connection = Get-SqlConnection -Server $ServerInstance -Database $DatabaseName
    $connection.Open()

    Initialize-Table -Connection $connection

    # Vider la table avant insertion
    Clear-Table -Connection $connection
} catch {
    Write-Host "Erreur de connexion a SQL Server: $_" -ForegroundColor Red
    Write-Host "Verifiez que SQL Server Express est installe et en cours d'execution." -ForegroundColor Yellow
    exit 1
}

# Collecter toutes les donnees
$allData = @()
$processedFiles = 0

foreach ($file in $excelFiles) {
    Write-Host ""
    Write-Host "Traitement: $($file.Name)" -ForegroundColor White

    $fileData = Read-ExcelData -FilePath $file.FullName -TargetSheetName $SheetName

    if ($fileData.Count -gt 0) {
        $allData += $fileData
        $processedFiles++
        Write-Host "  -> $($fileData.Count) enregistrements extraits" -ForegroundColor Green
    }
}

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "RESUME" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "Fichiers traites: $processedFiles / $($excelFiles.Count)" -ForegroundColor White
Write-Host "Total enregistrements: $($allData.Count)" -ForegroundColor White
Write-Host ""

# Inserer les donnees
if ($allData.Count -gt 0) {
    $insertedCount = Import-DataToSqlServer -Connection $connection -Data $allData
    Write-Host ""
    Write-Host "Enregistrements inseres: $insertedCount" -ForegroundColor Green
}

# Fermer la connexion
$connection.Close()

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  Import termine!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Cyan
