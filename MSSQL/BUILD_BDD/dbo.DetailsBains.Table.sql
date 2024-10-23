USE [ANODISATION]
GO
/****** Object:  Table [dbo].[DetailsBains]    Script Date: 21/10/2024 17:52:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[DetailsBains](
	[ClePrimaire] [int] IDENTITY(1,1) NOT NULL,
	[NomBain] [varchar](10) NOT NULL,
	[NumMatiere] [smallint] NOT NULL
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[DetailsBains] ADD  DEFAULT ('') FOR [NomBain]
GO
ALTER TABLE [dbo].[DetailsBains] ADD  DEFAULT ((0)) FOR [NumMatiere]
GO
