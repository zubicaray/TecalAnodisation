USE [ANODISATION]
GO
/****** Object:  Table [dbo].[Matieres]    Script Date: 21/10/2024 17:52:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Matieres](
	[Matiere] [varchar](30) NOT NULL,
	[TypeMatiere] [varchar](30) NOT NULL,
	[CompositionMatiere] [varchar](50) NOT NULL,
	[OrdrePourAffichage] [smallint] NOT NULL
) ON [PRIMARY]
GO
INSERT [dbo].[Matieres] ([Matiere], [TypeMatiere], [CompositionMatiere], [OrdrePourAffichage]) VALUES (N'1050', N'A5', N'', 1)
INSERT [dbo].[Matieres] ([Matiere], [TypeMatiere], [CompositionMatiere], [OrdrePourAffichage]) VALUES (N'1070/1080', N'A7/A8', N'', 2)
INSERT [dbo].[Matieres] ([Matiere], [TypeMatiere], [CompositionMatiere], [OrdrePourAffichage]) VALUES (N'2011/2030', N'AU4/5PB', N'', 4)
INSERT [dbo].[Matieres] ([Matiere], [TypeMatiere], [CompositionMatiere], [OrdrePourAffichage]) VALUES (N'2017/2024', N'AU 4 G', N'', 5)
INSERT [dbo].[Matieres] ([Matiere], [TypeMatiere], [CompositionMatiere], [OrdrePourAffichage]) VALUES (N'3003/3005', N'AM/G', N'', 6)
INSERT [dbo].[Matieres] ([Matiere], [TypeMatiere], [CompositionMatiere], [OrdrePourAffichage]) VALUES (N'5005', N'AG06', N'', 7)
INSERT [dbo].[Matieres] ([Matiere], [TypeMatiere], [CompositionMatiere], [OrdrePourAffichage]) VALUES (N'5083/8086', N'AG 4/5', N'', 8)
INSERT [dbo].[Matieres] ([Matiere], [TypeMatiere], [CompositionMatiere], [OrdrePourAffichage]) VALUES (N'5754', N'AG3M', N'', 9)
INSERT [dbo].[Matieres] ([Matiere], [TypeMatiere], [CompositionMatiere], [OrdrePourAffichage]) VALUES (N'6005', N'ASG 0,5', N'', 10)
INSERT [dbo].[Matieres] ([Matiere], [TypeMatiere], [CompositionMatiere], [OrdrePourAffichage]) VALUES (N'6060/6061', N'AGS', N'', 11)
INSERT [dbo].[Matieres] ([Matiere], [TypeMatiere], [CompositionMatiere], [OrdrePourAffichage]) VALUES (N'6082', N'ASGM 0,7', N'', 12)
INSERT [dbo].[Matieres] ([Matiere], [TypeMatiere], [CompositionMatiere], [OrdrePourAffichage]) VALUES (N'7020', N'AZ5G', N'', 13)
INSERT [dbo].[Matieres] ([Matiere], [TypeMatiere], [CompositionMatiere], [OrdrePourAffichage]) VALUES (N'7075', N'AZ5GU', N'', 14)
INSERT [dbo].[Matieres] ([Matiere], [TypeMatiere], [CompositionMatiere], [OrdrePourAffichage]) VALUES (N'A4/A9', N'Plaqué', N'', 3)
INSERT [dbo].[Matieres] ([Matiere], [TypeMatiere], [CompositionMatiere], [OrdrePourAffichage]) VALUES (N'AG3T', N'Sable ou coquille', N'', 15)
INSERT [dbo].[Matieres] ([Matiere], [TypeMatiere], [CompositionMatiere], [OrdrePourAffichage]) VALUES (N'AS 12/13', N'Coquille ou sous pression', N'', 18)
INSERT [dbo].[Matieres] ([Matiere], [TypeMatiere], [CompositionMatiere], [OrdrePourAffichage]) VALUES (N'AS2GT', N'Sable ou coquille', N'', 16)
INSERT [dbo].[Matieres] ([Matiere], [TypeMatiere], [CompositionMatiere], [OrdrePourAffichage]) VALUES (N'AS7G06', N'Sable ou coquille', N'', 17)
INSERT [dbo].[Matieres] ([Matiere], [TypeMatiere], [CompositionMatiere], [OrdrePourAffichage]) VALUES (N'AS9U3', N'Pression', N'', 19)
INSERT [dbo].[Matieres] ([Matiere], [TypeMatiere], [CompositionMatiere], [OrdrePourAffichage]) VALUES (N'AU5GT', N'Sable ou coquille', N'', 20)
ALTER TABLE [dbo].[Matieres] ADD  DEFAULT ('') FOR [Matiere]
GO
ALTER TABLE [dbo].[Matieres] ADD  DEFAULT ('') FOR [TypeMatiere]
GO
ALTER TABLE [dbo].[Matieres] ADD  DEFAULT ('') FOR [CompositionMatiere]
GO
ALTER TABLE [dbo].[Matieres] ADD  DEFAULT ((0)) FOR [OrdrePourAffichage]
GO
