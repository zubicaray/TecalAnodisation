USE [ANODISATION]
GO
/****** Object:  Table [dbo].[TempsDeplacements]    Script Date: 21/10/2024 17:52:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[TempsDeplacements](
	[depart] [smallint] NOT NULL,
	[arrivee] [smallint] NOT NULL,
	[lent] [smallint] NOT NULL,
	[normal] [smallint] NOT NULL,
	[rapide] [smallint] NOT NULL
) ON [PRIMARY]
GO