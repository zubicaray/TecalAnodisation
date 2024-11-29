@echo off
SETLOCAL ENABLEDELAYEDEXPANSION


SET serverName=SRV-APP-ANOD\SQLEXPRESS
SET databaseName=ANODISATION
:: Remplacer par le nom de la base de données
SET username=sa
:: Remplacer par le nom d'utilisateur SQL Server
SET password=Jeff_nenette
SET tempOutputFile=exportDetailsProd.sql

SET tempOutputFile2=exportGammes.sql

SET firstNumfiche=00086500

IF EXIST "%tempOutputFile%" DEL "%tempOutputFile%"
IF EXIST "%tempOutputFile2%" DEL "%tempOutputFile2%"


sqlcmd -S %serverName% -d %databaseName% -U %username% -P %password% -Q "EXEC dbo.GetInsertsProd '%firstNumfiche%'" -o "%tempOutputFile%" -h -1 -r1 -y 8000
sqlcmd -S %serverName% -d %databaseName% -U %username% -P %password% -Q "EXEC dbo.GetInsertsGammeCPO" -o "%tempOutputFile2%" -h -1 -r1

