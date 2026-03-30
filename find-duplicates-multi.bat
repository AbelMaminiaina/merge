@echo off
REM Script pour analyser les doublons sur plusieurs fichiers Excel
REM Accepte un dossier ou plusieurs fichiers Excel

setlocal enabledelayedexpansion

if "%~1"=="" (
    echo.
    echo ERREUR: Veuillez fournir le chemin d'un dossier ou des fichiers Excel
    echo.
    echo Usage: find-duplicates-multi.bat "C:\chemin\vers\dossier"
    echo    ou: find-duplicates-multi.bat "fichier1.xlsx" "fichier2.xlsx"
    echo    ou: glissez-deposez un dossier ou des fichiers Excel sur ce script
    echo.
    pause
    exit /b 1
)

set "FIRSTARG=%~1"

REM Verifier si c'est un dossier
if exist "%FIRSTARG%\*" (
    echo.
    echo ========================================
    echo ANALYSE DES DOUBLONS MULTI-FICHIERS
    echo ========================================
    echo Dossier: %FIRSTARG%
    echo.

    echo [1/2] Analyse "Nom FR"...
    powershell.exe -ExecutionPolicy Bypass -File "%~dp0Find-Duplicates-Multi-NomFR.ps1" -FolderPath "%FIRSTARG%"

    if !ERRORLEVEL! NEQ 0 (
        echo ERREUR lors de l'analyse "Nom FR"
        pause
        exit /b 1
    )

    echo [2/2] Analyse "trigramme"...
    powershell.exe -ExecutionPolicy Bypass -File "%~dp0Find-Duplicates-Multi-Trigramme.ps1" -FolderPath "%FIRSTARG%"

    if !ERRORLEVEL! NEQ 0 (
        echo ERREUR lors de l'analyse "trigramme"
        pause
        exit /b 1
    )
) else (
    REM C'est un ou plusieurs fichiers
    set "FILES="
    set "COUNT=0"

    :collectfiles
    if "%~1"=="" goto runanalysis
    if not exist "%~1" (
        echo ERREUR: Fichier introuvable: %~1
        pause
        exit /b 1
    )
    if "!FILES!"=="" (
        set "FILES=\"%~1\""
    ) else (
        set "FILES=!FILES!,\"%~1\""
    )
    set /a COUNT+=1
    shift
    goto collectfiles

    :runanalysis
    echo.
    echo ========================================
    echo ANALYSE DES DOUBLONS MULTI-FICHIERS
    echo ========================================
    echo Fichiers: !COUNT!
    echo.

    echo [1/2] Analyse "Nom FR"...
    powershell.exe -ExecutionPolicy Bypass -Command "& '%~dp0Find-Duplicates-Multi-NomFR.ps1' -FilePaths @(!FILES!)"

    if !ERRORLEVEL! NEQ 0 (
        echo ERREUR lors de l'analyse "Nom FR"
        pause
        exit /b 1
    )

    echo [2/2] Analyse "trigramme"...
    powershell.exe -ExecutionPolicy Bypass -Command "& '%~dp0Find-Duplicates-Multi-Trigramme.ps1' -FilePaths @(!FILES!)"

    if !ERRORLEVEL! NEQ 0 (
        echo ERREUR lors de l'analyse "trigramme"
        pause
        exit /b 1
    )
)

echo.
echo ========================================
echo TERMINE
echo ========================================
echo Fichiers generes:
echo   - Doublons_Multi_NomFR.xlsx
echo   - Doublons_Multi_Trigramme.xlsx
echo.
pause
