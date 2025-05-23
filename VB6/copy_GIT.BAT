@echo off

REM D�finition des chemins sources et de destination
set "source_frx=C:\Anodisation\VBFeuilles"
set "source_bas=C:\Anodisation\VBModules"
set "source_bat=C:\Anodisation"
set "source_class=C:\Anodisation\VBClasses"
set "destination=Z:\ZDEV\TECAL\SOURCES\TecalAnodisation\VB6\"

REM V�rifie si le dossier de destination existe, sinon le cr�e
if not exist "%destination%" (
    echo Cr�ation du dossier de destination...
    mkdir "%destination%"
)

REM Copie des fichiers .cls
echo Copie des fichiers .cls depuis %source_class% vers %destination%
xcopy "%source_class%\*.cls" "%destination%" /Y

REM Copie des fichiers .frm
echo Copie des fichiers .frm depuis %source_frx% vers %destination%
xcopy "%source_frx%\*.frm" "%destination%" /Y

REM Copie des fichiers .frx
echo Copie des fichiers .frx depuis %source_frx% vers %destination%
xcopy "%source_frx%\*.frx" "%destination%" /Y


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


echo Copie termin�e !
pause
