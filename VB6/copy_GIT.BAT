@echo off

REM Définition des chemins sources et de destination
set "source_frx=C:\Anodisation\VBFeuilles"
set "source_bas=C:\Anodisation\VBModules"
set "source_bat=C:\Anodisation"
set "destination=Z:\ZDEV\TECAL\SOURCES\TecalAnodisation\VB6\"

REM Vérifie si le dossier de destination existe, sinon le crée
if not exist "%destination%" (
    echo Création du dossier de destination...
    mkdir "%destination%"
)

REM Copie des fichiers .frm
echo Copie des fichiers .frm depuis %source_frx% vers %destination%
xcopy "%source_frx%\*.frm" "%destination%" /Y

REM Copie des fichiers .bas
echo Copie des fichiers .bas depuis %source_bas% vers %destination%
xcopy "%source_bas%\*.bas" "%destination%" /Y

REM Copie des fichiers .bat
echo Copie des fichiers .bat depuis %source_bat% vers %destination%
xcopy "%source_bat%\*.bat" "%destination%" /Y

REM Copie du projet vbp
set "source_vbp=C:\Anodisation\Anodisation.vbp"
echo Copie des vbp depuis %source_vbp% vers %destination%
xcopy "%source_vbp%" "%destination%" /Y

REM Copie du projet vbp
set "source_ini=C:\Anodisation\config.ini"
echo Copie des vbp depuis %source_ini% vers %destination%
xcopy "%source_ini%" "%destination%" /Y


echo Copie terminée !
pause
