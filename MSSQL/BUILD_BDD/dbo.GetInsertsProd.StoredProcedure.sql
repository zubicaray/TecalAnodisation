USE [ANODISATION]
GO
/****** Object:  StoredProcedure [dbo].[GetInsertsProd]    Script Date: 21/10/2024 17:52:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


CREATE   PROCEDURE [dbo].[GetInsertsProd] @NumFicheProduction nvarchar(30)
AS
	DECLARE @WHERE  nvarchar(300);

	SET @WHERE = ' from DetailsFichesProduction WHERE NumFicheProduction > ' + @NumFicheProduction + ' order by ClePrimaire';

	EXEC dbo.sp_generate_inserts @table_name='DetailsFichesProduction' , @ommit_identity = 1, 
		@from=@WHERE;

	SET @WHERE = ' from DetailsChargesProduction WHERE NumFicheProduction > ' + @NumFicheProduction + ' order by ClePrimaire';
	EXEC dbo.sp_generate_inserts @table_name='DetailsChargesProduction' , @ommit_identity = 1, 
		@from=@WHERE;

	SET @WHERE = ' from DetailsGammesProduction WHERE NumFicheProduction > ' + @NumFicheProduction + ' order by ClePrimaire';
	EXEC dbo.sp_generate_inserts @table_name='DetailsGammesProduction' , @ommit_identity = 1, 
		@from=@WHERE;

	SET @WHERE = ' from DetailsPhasesProduction WHERE NumFicheProduction > ' + @NumFicheProduction + ' order by ClePrimaire';
	EXEC dbo.sp_generate_inserts @table_name='DetailsPhasesProduction' , @ommit_identity = 1, 
		@from=@WHERE;
		
GO
