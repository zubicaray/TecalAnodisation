USE [ANODISATION]
GO
/****** Object:  Table [dbo].[DelaisActionneurs]    Script Date: 21/10/2024 17:52:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[DelaisActionneurs](
	[ClePrimaire] [int] IDENTITY(1,1) NOT NULL,
	[Actionneur] [varchar](50) NOT NULL,
	[Delai] [int] NOT NULL
) ON [PRIMARY]
GO
SET IDENTITY_INSERT [dbo].[DelaisActionneurs] ON 

INSERT [dbo].[DelaisActionneurs] ([ClePrimaire], [Actionneur], [Delai]) VALUES (1, N'Ouverture/Fermeture couvercles', 10)
INSERT [dbo].[DelaisActionneurs] ([ClePrimaire], [Actionneur], [Delai]) VALUES (2, N'Arrêt en position centrale cuve de nickel', 5)
INSERT [dbo].[DelaisActionneurs] ([ClePrimaire], [Actionneur], [Delai]) VALUES (3, N'Ouverture/Fermeture crochets', 2)
INSERT [dbo].[DelaisActionneurs] ([ClePrimaire], [Actionneur], [Delai]) VALUES (4, N'Niveau bas vers niveau intermédiaire et inverse', 3)
INSERT [dbo].[DelaisActionneurs] ([ClePrimaire], [Actionneur], [Delai]) VALUES (5, N'Niveau bas vers niveau haut et inverse', 10)
INSERT [dbo].[DelaisActionneurs] ([ClePrimaire], [Actionneur], [Delai]) VALUES (6, N'Transversal', 13)
INSERT [dbo].[DelaisActionneurs] ([ClePrimaire], [Actionneur], [Delai]) VALUES (7, N'1 mètre de parcours du pont', 3)
INSERT [dbo].[DelaisActionneurs] ([ClePrimaire], [Actionneur], [Delai]) VALUES (8, N'8,60 mètres de parcours du pont', 18)
SET IDENTITY_INSERT [dbo].[DelaisActionneurs] OFF
ALTER TABLE [dbo].[DelaisActionneurs] ADD  DEFAULT ('') FOR [Actionneur]
GO
ALTER TABLE [dbo].[DelaisActionneurs] ADD  DEFAULT ((0)) FOR [Delai]
GO
