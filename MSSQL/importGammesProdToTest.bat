sqlcmd -S ZUBI-STUDIO\SQLEXPRESS -d ANODISATION -U sa -P Jeff_nenette  -Q "EXEC dbo.deleteTables"

sqlcmd -S VB\SQLEXPRESSANO -d ANODISATION -U sa -P Jeff_nénette  -Q "EXEC dbo.importProdTables"