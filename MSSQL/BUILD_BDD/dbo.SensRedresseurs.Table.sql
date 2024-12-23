USE [ANODISATION]
GO
/****** Object:  Table [dbo].[SensRedresseurs]    Script Date: 21/10/2024 17:52:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[SensRedresseurs](
	[SensRedresseur] [smallint] NOT NULL,
	[LibelleSensRedresseur] [varchar](20) NOT NULL,
	[OrdrePourAffichage] [smallint] NOT NULL
) ON [PRIMARY]
GO
INSERT [dbo].[SensRedresseurs] ([SensRedresseur], [LibelleSensRedresseur], [OrdrePourAffichage]) VALUES (0, N'ANODIQUE', 1)
INSERT [dbo].[SensRedresseurs] ([SensRedresseur], [LibelleSensRedresseur], [OrdrePourAffichage]) VALUES (1, N'CATHODIQUE', 2)
ALTER TABLE [dbo].[SensRedresseurs] ADD  DEFAULT ((0)) FOR [SensRedresseur]
GO
ALTER TABLE [dbo].[SensRedresseurs] ADD  DEFAULT ('') FOR [LibelleSensRedresseur]
GO
ALTER TABLE [dbo].[SensRedresseurs] ADD  DEFAULT ((0)) FOR [OrdrePourAffichage]
GO
