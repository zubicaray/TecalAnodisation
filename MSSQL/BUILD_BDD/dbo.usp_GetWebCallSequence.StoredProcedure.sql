USE [ANODISATION]
GO
/****** Object:  StoredProcedure [dbo].[usp_GetWebCallSequence]    Script Date: 21/10/2024 17:52:31 ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO

CREATE PROC [dbo].[usp_GetWebCallSequence] AS DECLARE @WebCallSequence bigint; 
SET NOCOUNT ON
DECLARE @Today DATETIME
DECLARE @bddJOUR as DATETIME

SET @bddJour = (SELECT jour  FROM WebCallSequence)


SET @Today = dateadd(dd, datediff(dd, 0, getdate()), 0)

if @Today > @bddJour
BEGIN
	UPDATE dbo.WebCallSequence SET @WebCallSequence = WebCallSequence = 1
	UPDATE dbo.WebCallSequence SET jour = @Today
END
ELSE
	UPDATE dbo.WebCallSequence SET @WebCallSequence = WebCallSequence = WebCallSequence + 1

SELECT @WebCallSequence AS WebCallSequence;
GO
