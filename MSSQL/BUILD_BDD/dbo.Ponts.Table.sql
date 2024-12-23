USE [ANODISATION]
GO
/****** Object:  Table [dbo].[Ponts]    Script Date: 21/10/2024 17:52:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Ponts](
	[NumPont] [smallint] NOT NULL,
	[NomPont] [varchar](10) NOT NULL,
	[LibellePont] [varchar](50) NOT NULL
) ON [PRIMARY]
GO
INSERT [dbo].[Ponts] ([NumPont], [NomPont], [LibellePont]) VALUES (1, N'P1', N'Pont 1')
INSERT [dbo].[Ponts] ([NumPont], [NomPont], [LibellePont]) VALUES (2, N'P2', N'Pont 2')
ALTER TABLE [dbo].[Ponts] ADD  DEFAULT ((0)) FOR [NumPont]
GO
ALTER TABLE [dbo].[Ponts] ADD  DEFAULT ('') FOR [NomPont]
GO
ALTER TABLE [dbo].[Ponts] ADD  DEFAULT ('') FOR [LibellePont]
GO
