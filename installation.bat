@echo off
setlocal

echo ============================================
echo   Installation de Portes Logiques pour Word
echo ============================================
echo.

REM Verifier que le fichier .dotm est present
if not exist "%~dp0portes_logiques.dotm" (
    echo ERREUR : Le fichier portes_logiques.dotm est introuvable.
    echo Assurez-vous qu'il se trouve dans le meme dossier que ce script.
    echo.
    pause
    exit /b 1
)

REM Verifier si Word est ouvert
tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>nul | find /I "WINWORD.EXE" >nul
if not errorlevel 1 (
    echo ATTENTION : Word est actuellement ouvert.
    echo Veuillez fermer Word puis relancer ce script.
    echo.
    pause
    exit /b 1
)

REM Debloquer le fichier (retirer la Marque du Web si presente)
echo Deblocage du fichier...
powershell -NoProfile -Command "Unblock-File -Path '%~dp0portes_logiques.dotm'" 2>nul

REM Chemin de destination
set "DEST=%APPDATA%\Microsoft\Word\STARTUP"

REM Creer le dossier s'il n'existe pas
if not exist "%DEST%" (
    echo Creation du dossier STARTUP...
    mkdir "%DEST%"
    if errorlevel 1 (
        echo ERREUR : Impossible de creer le dossier.
        pause
        exit /b 1
    )
)

REM Copier le fichier
echo Copie de portes_logiques.dotm vers :
echo   %DEST%
echo.

copy /Y "%~dp0portes_logiques.dotm" "%DEST%\portes_logiques.dotm" >nul
if errorlevel 1 (
    echo ERREUR : La copie a echoue.
    echo Verifiez que le fichier n'est pas verrouille.
    pause
    exit /b 1
)

echo ============================================
echo   Installation terminee avec succes !
echo ============================================
echo.
pause
exit /b 0