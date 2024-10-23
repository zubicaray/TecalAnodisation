USE [ANODISATION]
GO
/****** Object:  Table [dbo].[Bains]    Script Date: 21/10/2024 17:52:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Bains](
	[NumBain] [smallint] NOT NULL,
	[NomBain] [varchar](10) NOT NULL,
	[LibelleBain] [varchar](50) NOT NULL
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[Bains] ADD  DEFAULT ((0)) FOR [NumBain]
GO
ALTER TABLE [dbo].[Bains] ADD  DEFAULT ('') FOR [NomBain]
GO
ALTER TABLE [dbo].[Bains] ADD  DEFAULT ('') FOR [LibelleBain]
GO
