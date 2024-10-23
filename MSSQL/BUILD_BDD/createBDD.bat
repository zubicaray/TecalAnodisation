:: @echo off
setlocal

:: Configuration des variables
set ServerName=ZUBI-STUDIO\SQLEXPRESS
set DatabaseName=ANODISATION
set UserName=sa
set Password=Jeff_nenette

:: Commande SQL pour générer le script de suppression des tables
set SQLScript=DECLARE @sql NVARCHAR(MAX) = N''; 
set SQLScript=%SQLScript% SELECT @sql += 'DROP TABLE ' + QUOTENAME(SCHEMA_NAME(schema_id)) + '.' + QUOTENAME(name) + '; ' 
set SQLScript=%SQLScript% FROM sys.tables; 
set SQLScript=%SQLScript% EXEC sp_executesql @sql;

:: Exécution de la commande SQL pour supprimer les tables
sqlcmd -S %ServerName% -d %DatabaseName% -U %UserName% -P %Password% -Q "%SQLScript%"


for %%G in (dbo.*.sql) do sqlcmd -S %ServerName% -d %DatabaseName% -U %UserName% -P %Password% /i"%%G"

endlocal


