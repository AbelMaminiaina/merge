@echo off
REM Script pour analyser les doublons dans Nom FR et trigramme

setlocal

if "%~1"=="" (
    echo.
    echo ERREUR: Veuillez fournir le chemin du fichier Excel
    echo.
    echo Usage: find-duplicates-all.bat "C:\chemin\vers\fichier.xlsx"
    echo Ou glissez-deposez un fichier Excel sur ce script
    echo.
    pause
    exit /b 1
)

set "FILEPATH=%~1"

if not exist "%FILEPATH%" (
    echo.
    echo ERREUR: Le fichier n'existe pas: %FILEPATH%
    echo.
    pause
    exit /b 1
)

echo.
echo ========================================
echo ANALYSE DES DOUBLONS
echo ========================================
echo Fichier: %FILEPATH%
echo.

echo [1/2] Analyse "Nom FR"...
powershell.exe -ExecutionPolicy Bypass -File "%~dp0Find-Duplicates-NomFR.ps1" -FilePath "%FILEPATH%"

if %ERRORLEVEL% NEQ 0 (
    echo ERREUR lors de l'analyse "Nom FR"
    pause
    exit /b 1
)

echo [2/2] Analyse "trigramme"...
powershell.exe -ExecutionPolicy Bypass -File "%~dp0Find-Duplicates-Trigramme.ps1" -FilePath "%FILEPATH%"

if %ERRORLEVEL% NEQ 0 (
    echo ERREUR lors de l'analyse "trigramme"
    pause
    exit /b 1
)

echo.
echo ========================================
echo TERMINE
echo ========================================
echo Les fichiers ont ete crees dans le meme dossier.
echo.
pause
