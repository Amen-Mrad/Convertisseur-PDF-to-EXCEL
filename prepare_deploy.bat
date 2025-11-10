@echo off
REM Script pour préparer le package de déploiement après build PyInstaller

echo ========================================
echo Preparation du package de deploiement
echo ========================================

REM Nettoyer le dossier deploy s'il existe
if exist "C:\Deploy\pdfTOexcel" (
    echo Suppression de l'ancien dossier deploy...
    rmdir /s /q "C:\Deploy\pdfTOexcel"
)

REM Créer le dossier deploy
mkdir "C:\Deploy\pdfTOexcel" 2>nul

REM Copier le contenu de dist\convertisseur_unifie
echo Copie de l'executable et des DLLs...
xcopy "dist\convertisseur_unifie\*" "C:\Deploy\pdfTOexcel\" /E /I /Y

REM Copier le dossier Converters
echo Copie du dossier Converters...
xcopy "Converters\*" "C:\Deploy\pdfTOexcel\Converters\" /E /I /Y

REM Copier le dossier logo
echo Copie du dossier logo...
xcopy "logo\*" "C:\Deploy\pdfTOexcel\logo\" /E /I /Y

REM Copier le script de lancement VBS
echo Copie du script de lancement...
copy "Configuration\Lancer_Convertisseur.vbs" "C:\Deploy\pdfTOexcel\" >nul

echo.
echo ========================================
echo Package pret dans C:\Deploy\pdfTOexcel
echo ========================================
echo.
echo Ce dossier contient tout ce qu'il faut pour deployer sur le serveur.
echo Copiez simplement tout le contenu de C:\Deploy\pdfTOexcel vers le serveur.
echo.
pause

