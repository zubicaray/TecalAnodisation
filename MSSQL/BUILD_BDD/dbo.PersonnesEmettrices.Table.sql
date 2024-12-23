USE [ANODISATION]
GO
/****** Object:  Table [dbo].[PersonnesEmettrices]    Script Date: 21/10/2024 17:52:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[PersonnesEmettrices](
	[ClePrimaire] [int] IDENTITY(1,1) NOT NULL,
	[Nom] [varchar](30) NOT NULL,
	[Prenom] [varchar](30) NOT NULL,
	[NomComplet] [varchar](50) NOT NULL
) ON [PRIMARY]
GO
SET IDENTITY_INSERT [dbo].[PersonnesEmettrices] ON 

INSERT [dbo].[PersonnesEmettrices] ([ClePrimaire], [Nom], [Prenom], [NomComplet]) VALUES (1, N'VERBRUGGE', N'Jean-François', N'Jean-François VERBRUGGE')
INSERT [dbo].[PersonnesEmettrices] ([ClePrimaire], [Nom], [Prenom], [NomComplet]) VALUES (2, N'TARDIF', N'Catherine', N'Catherine TARDIF')
INSERT [dbo].[PersonnesEmettrices] ([ClePrimaire], [Nom], [Prenom], [NomComplet]) VALUES (3, N'DELY', N'Amandine', N'Amandine DELY')
INSERT [dbo].[PersonnesEmettrices] ([ClePrimaire], [Nom], [Prenom], [NomComplet]) VALUES (4, N'PERISSINOTTO', N'Joseph', N'Joseph PERISSINOTTO')
INSERT [dbo].[PersonnesEmettrices] ([ClePrimaire], [Nom], [Prenom], [NomComplet]) VALUES (5, N'AUBOINE', N'Claude', N'Claude AUBOINE')
INSERT [dbo].[PersonnesEmettrices] ([ClePrimaire], [Nom], [Prenom], [NomComplet]) VALUES (6, N'BRIET', N'Sylvie', N'Sylvie BRIET')
INSERT [dbo].[PersonnesEmettrices] ([ClePrimaire], [Nom], [Prenom], [NomComplet]) VALUES (7, N'SARTRE', N'Eric', N'Eric SARTRE')
SET IDENTITY_INSERT [dbo].[PersonnesEmettrices] OFF
ALTER TABLE [dbo].[PersonnesEmettrices] ADD  DEFAULT ('') FOR [Nom]
GO
ALTER TABLE [dbo].[PersonnesEmettrices] ADD  DEFAULT ('') FOR [Prenom]
GO
ALTER TABLE [dbo].[PersonnesEmettrices] ADD  DEFAULT ('') FOR [NomComplet]
GO
