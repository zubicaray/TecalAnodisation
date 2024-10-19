@echo off
setlocal enabledelayedexpansion

rem Définit le répertoire de base
set "base_dir=C:\AnodisationTEST\VBFeuilles"

rem Définit les chaînes à rechercher et à remplacer
set "search_str=Provider=SQLNCLI11;Server=VB-LANLIGNE2-20\SQLEXPRESSANO;Database=ANODISATION;Uid=sa; Pwd=Jeff_nenette;"

set "replace_str=Provider=SQLOLEDB.1;Integrated Security=SSPI;Initial Catalog=ANODISATION;Uid=sa; Pwd=sa;Data Source=SRV2003\SQLEXPRESS;"
rem set "replace_str=Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=ANODISATION;Uid=sa; Pwd=sa;Data Source=SRV2003\SQLEXPRESS;Connect Timeout=3;"

rem Parcourt tous les fichiers .frm dans le répertoire et les sous-répertoires
for /r "%base_dir%" %%f in (*.frm) do (
    set "file=%%f"
    echo Traitement de !file!

    rem Vérifie si le fichier contient la chaîne à remplacer
    findstr /m /c:"%search_str%" "!file!" >nul
    if !errorlevel! == 0 (
        rem Remplace la chaîne dans le fichier
        echo Remplacement dans !file!
        powershell -command "(Get-Content '!file!') -replace [regex]::Escape('%search_str%'), '%replace_str%' | Set-Content '!file!'"
    )
)

echo Remplacement terminé.
pause
