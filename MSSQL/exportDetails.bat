@echo off
SETLOCAL ENABLEDELAYEDEXPANSION


SET serverName=SRV-APP-ANOD\SQLEXPRESS
SET databaseName=ANODISATION
:: Remplacer par le nom de la base de données
SET username=sa
:: Remplacer par le nom d'utilisateur SQL Server
SET password=Jeff_nenette

SET tempOutputFile=exportGammes.sql

SET firstNumfiche=00087661

IF EXIST "%tempOutputFile%" DEL "%tempOutputFile%"


sqlcmd -S %serverName% -d %databaseName% -U %username% -P %password% -Q "EXEC dbo.GetInsertsGammeCPO" -o "%tempOutputFile%" -h -1 -r1

