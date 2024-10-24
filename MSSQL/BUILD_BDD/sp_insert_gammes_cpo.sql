USE [ANODISATION]
GO
/****** Object:  StoredProcedure [dbo].[GetInsertsGammeCPO]    Script Date: 24/10/2024 11:20:51 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


CREATE OR ALTER   PROCEDURE [dbo].[GetInsertsGammeCPO] 
AS
	DECLARE @WHERE  nvarchar(300);

	SET @WHERE = ' from DetailsGammesAnodisation ';

	EXEC dbo.sp_generate_inserts @table_name='DetailsGammesAnodisation' , @ommit_identity = 1, 
		@from=@WHERE;

	SET @WHERE = ' from TempsDeplacements ';
	EXEC dbo.sp_generate_inserts @table_name='TempsDeplacements' , @ommit_identity = 1, 
		@from=@WHERE;

	SET @WHERE = ' from GammesAnodisation ';
	EXEC dbo.sp_generate_inserts @table_name='GammesAnodisation' , @ommit_identity = 1, 
		@from=@WHERE;

	SET @WHERE = ' from CalibrageTempsGammes ';
	EXEC dbo.sp_generate_inserts @table_name='CalibrageTempsGammes' , @ommit_identity = 1, 
		@from=@WHERE;
		

