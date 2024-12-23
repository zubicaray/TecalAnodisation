USE [ANODISATION]
GO
/****** Object:  Table [dbo].[Cuves]    Script Date: 21/10/2024 17:52:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Cuves](
	[NumCuve] [smallint] NOT NULL,
	[NomCuve] [varchar](10) NOT NULL,
	[LibelleCuve] [varchar](50) NOT NULL,
	[GestionAPI] [bit] NOT NULL,
	[PresencePompe] [bit] NOT NULL,
	[NbrChauffages] [smallint] NOT NULL,
	[PresenceRefroidissementBain] [bit] NOT NULL,
	[PresenceNiveauBas] [bit] NOT NULL,
	[PresenceNiveauHaut] [bit] NOT NULL,
	[PresenceEVEau] [bit] NOT NULL,
	[PresenceAnalyseurAnodisation] [bit] NOT NULL
) ON [PRIMARY]
GO
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (1, N'C00', N'Dégraissage', 1, 0, 1, 0, 1, 1, 1, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (2, N'DEC', N'Décapage', 1, 0, 1, 0, 1, 1, 1, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (3, N'SAT', N'Satinage', 0, 0, 0, 0, 1, 1, 1, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (4, N'C03', N'Rinçage soude', 0, 0, 0, 0, 0, 0, 0, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (5, N'C04', N'Rinçage soude/dégraissage ', 0, 0, 0, 0, 0, 0, 0, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (6, N'C05', N'Dégraissage acide', 0, 0, 1, 0, 1, 1, 1, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (7, N'C06', N'Rinçage Mt brillantage', 0, 0, 1, 0, 1, 1, 1, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (8, N'C07', N'Brillantage', 1, 0, 1, 0, 1, 1, 1, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (9, N'C08', N'Rinçage brillantage', 0, 0, 0, 0, 0, 0, 0, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (10, N'C09', N'Rinçage brillantage', 0, 0, 0, 0, 0, 0, 0, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (11, N'C10', N'Blanchiment', 0, 0, 0, 0, 0, 0, 0, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (12, N'C11', N'Rinçage', 0, 0, 0, 0, 0, 0, 0, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (13, N'C12', N'Rinçage', 0, 0, 0, 0, 0, 0, 0, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (14, N'C13', N'Anodisation 1', 1, 0, 1, 1, 1, 1, 1, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (15, N'C14', N'Anodisation 2', 1, 0, 1, 1, 1, 1, 1, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (16, N'C15', N'Anodisation 3', 1, 0, 1, 1, 1, 1, 1, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (17, N'C16', N'Anodisation 4', 0, 0, 1, 1, 0, 0, 0, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (18, N'C17', N'Rinçage anodisation', 0, 0, 0, 0, 0, 0, 0, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (19, N'C18', N'Rinçage anodisation', 0, 0, 0, 0, 0, 0, 0, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (20, N'C19', N'Spectrocoloration', 0, 0, 1, 0, 1, 1, 1, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (21, N'C20', N'Rinçage', 0, 0, 0, 0, 0, 0, 0, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (22, N'C21', N'Rinçage', 0, 0, 0, 0, 0, 0, 0, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (23, N'C22', N'Coloration or', 1, 0, 1, 0, 1, 1, 1, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (24, N'C23', N'Coloration orange', 0, 0, 0, 0, 0, 0, 0, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (25, N'C24', N'RESERVE 2', 0, 0, 0, 0, 0, 0, 0, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (26, N'C25', N'Imprégnation à froid', 0, 0, 0, 0, 0, 0, 0, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (27, N'C26', N'Rinçage imprégnation', 0, 0, 0, 0, 0, 0, 0, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (28, N'C27', N'Imprégnation à froid', 1, 0, 1, 0, 1, 1, 1, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (29, N'C28', N'Coloration noire', 1, 0, 1, 0, 1, 1, 1, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (30, N'C29', N'Rinçage noir', 0, 0, 0, 0, 0, 0, 0, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (31, N'C30', N'Rinçage final', 0, 0, 0, 0, 0, 0, 0, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (32, N'C31', N'Colmatage à chaud', 1, 0, 1, 0, 1, 1, 1, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (33, N'C32', N'Colmatage chaud', 1, 0, 1, 0, 1, 1, 1, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (34, N'C33', N'Conversion chimique', 0, 0, 1, 0, 1, 1, 1, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (35, N'C34', N'Rinçage final', 0, 0, 0, 0, 0, 0, 0, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (36, N'C35', N'Rinçage final', 0, 0, 0, 0, 0, 0, 0, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (37, N'C36	', N'Réserve', 0, 0, 0, 0, 0, 0, 0, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (38, N'C37	', N'Etuve', 0, 0, 0, 0, 0, 0, 0, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (39, N'C38', N'Basculeur', 0, 0, 0, 0, 0, 0, 0, 0)
INSERT [dbo].[Cuves] ([NumCuve], [NomCuve], [LibelleCuve], [GestionAPI], [PresencePompe], [NbrChauffages], [PresenceRefroidissementBain], [PresenceNiveauBas], [PresenceNiveauHaut], [PresenceEVEau], [PresenceAnalyseurAnodisation]) VALUES (40, N'C02', N'Réserve', 0, 0, 0, 0, 0, 0, 0, 0)
ALTER TABLE [dbo].[Cuves] ADD  DEFAULT ((0)) FOR [NumCuve]
GO
ALTER TABLE [dbo].[Cuves] ADD  DEFAULT ('') FOR [NomCuve]
GO
ALTER TABLE [dbo].[Cuves] ADD  DEFAULT ('') FOR [LibelleCuve]
GO
ALTER TABLE [dbo].[Cuves] ADD  DEFAULT ((0)) FOR [GestionAPI]
GO
ALTER TABLE [dbo].[Cuves] ADD  DEFAULT ((0)) FOR [PresencePompe]
GO
ALTER TABLE [dbo].[Cuves] ADD  DEFAULT ((0)) FOR [NbrChauffages]
GO
ALTER TABLE [dbo].[Cuves] ADD  DEFAULT ((0)) FOR [PresenceRefroidissementBain]
GO
ALTER TABLE [dbo].[Cuves] ADD  DEFAULT ((0)) FOR [PresenceNiveauBas]
GO
ALTER TABLE [dbo].[Cuves] ADD  DEFAULT ((0)) FOR [PresenceNiveauHaut]
GO
ALTER TABLE [dbo].[Cuves] ADD  DEFAULT ((0)) FOR [PresenceEVEau]
GO
ALTER TABLE [dbo].[Cuves] ADD  DEFAULT ((0)) FOR [PresenceAnalyseurAnodisation]
GO
