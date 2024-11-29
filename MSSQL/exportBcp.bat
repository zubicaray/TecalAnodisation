@echo off
SETLOCAL ENABLEDELAYEDEXPANSION


SET serverName=SRV-APP-ANOD\SQLEXPRESS
SET databaseName=ANODISATION
:: Remplacer par le nom de la base de données
SET username=sa
:: Remplacer par le nom d'utilisateur SQL Server
SET password=Jeff_nenette


bcp "select * from ANODISATION.dbo.DetailsFichesProduction where NumFicheProduction > '%firstNumfiche%'" queryout "DetailsFichesProduction.bcp" -n -S "tcp:%serverName%,1433;TrustServerCertificate=yes" -U %username% -P %password%
bcp "select * from ANODISATION.dbo.DetailsChargesProduction where NumFicheProduction > '%firstNumfiche%'" queryout "DetailsChargesProduction.bcp" -n -S "tcp:%serverName%,1433;TrustServerCertificate=yes" -U %username% -P %password%
bcp "select * from ANODISATION.dbo.DetailsPhasesProduction where NumFicheProduction > '%firstNumfiche%'" queryout "DetailsPhasesProduction.bcp" -n -S "tcp:%serverName%,1433;TrustServerCertificate=yes" -U %username% -P %password%
bcp "select * from ANODISATION.dbo.DetailsGammesProduction where NumFicheProduction > '%firstNumfiche%'" queryout "DetailsGammesProduction.bcp" -n -S "tcp:%serverName%,1433;TrustServerCertificate=yes" -U %username% -P %password%



