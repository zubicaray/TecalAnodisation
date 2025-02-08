@echo off
REM Définir les chemins vers VB6, le projet et le dossier de destination
@echo off
REM Définir les chemins vers VB6, le projet et le dossier de destination
SET VB6_PATH="C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"
SET PROJECT_PATH="C:\Anodisation\anodisation.vbp"
SET DEST_PATH="Z:\ANO"
SET BIN_PATH="C:\Anodisation\Anodisation_9.3.0.exe"

REM Compiler le projet
echo Compilation du projet anodisation...
%VB6_PATH% /make %PROJECT_PATH%

REM Vérifier si la compilation a réussi
IF %ERRORLEVEL% NEQ 0 (
    echo La compilation a échoué.
    exit /b %ERRORLEVEL%
) ELSE (
    echo La compilation a réussi !

    REM Génération des manifests avec mt.exe pour chaque OCX/DLL nécessaire
    echo Génération des manifests...
    
    mt.exe -nologo -manifest "COMDLG32.OCX.manifest" -outputresource:"%BIN_PATH%" || goto :error_mt
    mt.exe -nologo -manifest "mscomctl.OCX.manifest" -outputresource:"%BIN_PATH%" || goto :error_mt
    mt.exe -nologo -manifest "richtx32.OCX.manifest" -outputresource:"%BIN_PATH%" || goto :error_mt
    mt.exe -nologo -manifest "picclp32.OCX.manifest" -outputresource:"%BIN_PATH%" || goto :error_mt
    mt.exe -nologo -manifest "msdatgrd.OCX.manifest" -outputresource:"%BIN_PATH%" || goto :error_mt
    mt.exe -nologo -manifest "mscomct2.OCX.manifest" -outputresource:"%BIN_PATH%" || goto :error_mt
    mt.exe -nologo -manifest "msmask32.OCX.manifest" -outputresource:"%BIN_PATH%" || goto :error_mt
    mt.exe -nologo -manifest "tabctl32.OCX.manifest" -outputresource:"%BIN_PATH%" || goto :error_mt
    mt.exe -nologo -manifest "GRAPH32.OCX.manifest" -outputresource:"%BIN_PATH%" || goto :error_mt
    mt.exe -nologo -manifest "msdatlst.OCX.manifest" -outputresource:"%BIN_PATH%" || goto :error_mt
    mt.exe -nologo -manifest "AppOcxClient.OCX.manifest" -outputresource:"%BIN_PATH%" || goto :error_mt
    mt.exe -nologo -manifest "todg8.OCX.manifest" -outputresource:"%BIN_PATH%" || goto :error_mt
    mt.exe -nologo -manifest "c1sizer.OCX.manifest" -outputresource:"%BIN_PATH%" || goto :error_mt
    mt.exe -nologo -manifest "vsflex8l.OCX.manifest" -outputresource:"%BIN_PATH%" || goto :error_mt
    mt.exe -nologo -manifest "c1awk.OCX.manifest" -outputresource:"%BIN_PATH%" || goto :error_mt
    mt.exe -nologo -manifest "tizonex8.dll.manifest" -outputresource:"%BIN_PATH%" || goto :error_mt
    mt.exe -nologo -manifest "truedc8.OCX.manifest" -outputresource:"%BIN_PATH%" || goto :error_mt
    mt.exe -nologo -manifest "vsflex8d.OCX.manifest" -outputresource:"%BIN_PATH%" || goto :error_mt
    mt.exe -nologo -manifest "vsflex8n.OCX.manifest" -outputresource:"%BIN_PATH%" || goto :error_mt
    mt.exe -nologo -manifest "vsstr8.OCX.manifest" -outputresource:"%BIN_PATH%" || goto :error_mt
    mt.exe -nologo -manifest "vspdf8.OCX.manifest" -outputresource:"%BIN_PATH%" || goto :error_mt
    mt.exe -nologo -manifest "vsflex8.ocx.manifest" -outputresource:"%BIN_PATH%" || goto :error_mt
    mt.exe -nologo -manifest "MSBIND.DLL.manifest" -outputresource:"%BIN_PATH%" || goto :error_mt
    mt.exe -nologo -manifest "MSDBRPTR.DLL.manifest" -outputresource:"%BIN_PATH%" || goto :error_mt
    mt.exe -nologo -manifest "msstdfmt.dll.manifest" -outputresource:"%BIN_PATH%" || goto :error_mt

    REM Copier le binaire dans le dossier Z:\ANO
    echo Copie du fichier binaire dans %DEST_PATH%...
    COPY "%BIN_PATH%" "%DEST_PATH%" || goto :error_copy

    echo Le fichier a été copié avec succès dans %DEST_PATH% !
)

goto :end

:error_mt
echo Erreur lors de la génération des manifests avec mt.exe.
pause
exit /b 1

:error_copy
echo Échec de la copie du fichier.
pause
exit /b 1

:end
pause
