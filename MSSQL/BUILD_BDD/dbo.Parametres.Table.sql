USE [ANODISATION]
GO
/****** Object:  Table [dbo].[Parametres]    Script Date: 21/10/2024 17:52:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Parametres](
	[libellé] [nchar](20) NOT NULL,
	[valeur] [int] NOT NULL
) ON [PRIMARY]
GO
INSERT [dbo].[Parametres] ([libellé], [valeur]) VALUES (N'DISTANCE_SECURITE   ', 6300)
INSERT [dbo].[Parametres] ([libellé], [valeur]) VALUES (N'DEBUG_MODE          ', 1)
ALTER TABLE [dbo].[Parametres] ADD  DEFAULT ('') FOR [libellé]
GO
ALTER TABLE [dbo].[Parametres] ADD  DEFAULT ((0)) FOR [valeur]
GO
