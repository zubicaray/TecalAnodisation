@echo off
setlocal enabledelayedexpansion

rem D�finit le r�pertoire de base
set "base_dir=C:\AnodisationTEST\VBFeuilles"

rem D�finit les cha�nes � rechercher et � remplacer
set "search_str=Provider=SQLNCLI11;Server=VB-LANLIGNE2-20\SQLEXPRESSANO;Database=ANODISATION;Uid=sa; Pwd=Jeff_nenette;"

set "replace_str=Provider=SQLOLEDB.1;Integrated Security=SSPI;Initial Catalog=ANODISATION;Uid=sa; Pwd=sa;Data Source=SRV2003\SQLEXPRESS;"
rem set "replace_str=Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=ANODISATION;Uid=sa; Pwd=sa;Data Source=SRV2003\SQLEXPRESS;Connect Timeout=3;"

rem Parcourt tous les fichiers .frm dans le r�pertoire et les sous-r�pertoires
for /r "%base_dir%" %%f in (*.frm) do (
    set "file=%%f"
    echo Traitement de !file!

    rem V�rifie si le fichier contient la cha�ne � remplacer
    findstr /m /c:"%search_str%" "!file!" >nul
    if !errorlevel! == 0 (
        rem Remplace la cha�ne dans le fichier
        echo Remplacement dans !file!
        powershell -command "(Get-Content '!file!') -replace [regex]::Escape('%search_str%'), '%replace_str%' | Set-Content '!file!'"
    )
)

echo Remplacement termin�.
pause
