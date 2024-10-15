@echo off
FOR /f "tokens=*" %%s IN (FGammesAnodisation.frm) DO (
  SET Texts=%%s
)
set Texts=%Texts:SQLNCLI11=MSOLEDBSQL18%

FOR /F "tokens=* delims=" %%x IN (FGammesAnodisation.frm) DO SET text=%%x
ECHO %Texts% > "BATCH\FGammesAnodisation.frm" :: the path location of the txt file