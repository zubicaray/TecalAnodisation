USE [ANODISATION]
GO
/****** Object:  Table [dbo].[Actions]    Script Date: 21/10/2024 17:52:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Actions](
	[ClePrimaire] [int] IDENTITY(1,1) NOT NULL,
	[NumAction] [smallint] NOT NULL,
	[CodeAction] [varchar](20) NOT NULL,
	[LibelleAction] [varchar](100) NOT NULL,
	[ParametreOuiNon] [bit] NOT NULL,
	[LibelleParametre] [varchar](50) NOT NULL
) ON [PRIMARY]
GO
SET IDENTITY_INSERT [dbo].[Actions] ON 

INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (1, 1, N'CHGT1', N'Poste de chargement 1', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (2, 2, N'CHGT2', N'Poste de chargement 2', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (5, 4, N'C00', N'Dégraissage', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (6, 5, N'DEC', N'Décapage', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (7, 6, N'SAT', N'Satinage', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (8, 8, N'C03', N'Rinçage soude', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (9, 9, N'C04', N'Rinçage dégraissage', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (10, 10, N'C05', N'Brillantage n°1', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (11, 11, N'C06', N'Rinçage Mt brillantage', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (12, 12, N'C07', N'Dérochage acide', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (13, 13, N'C08', N'Rinçage brillantage', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (14, 14, N'C09', N'Rinçage brillantage', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (15, 15, N'C10', N'Neutralisation', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (16, 16, N'C11', N'Rinçage blanchiment', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (17, 17, N'C12', N'Blanchiment', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (18, 18, N'C13', N'Anodisation 1', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (19, 19, N'C14', N'Anodisation 2', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (20, 20, N'C15', N'Anodisation 3', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (21, 21, N'C16', N'Anodisation 4', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (22, 22, N'C17', N'Rinçage anodisation', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (23, 23, N'C18', N'Rinçage anodisation', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (24, 24, N'C19', N'Spectrocoloration', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (25, 25, N'C20', N'Rinçage', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (27, 26, N'C21', N'Rinçage', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (28, 27, N'C22', N'Coloration or', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (29, 28, N'C23', N'RESERVE 1', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (30, 29, N'C24', N'RESERVE 2', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (31, 30, N'C25', N'Imprégnation à froid', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (32, 31, N'C26', N'RESERVE 4', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (33, 32, N'C27', N'Imprégnation à froid', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (34, 33, N'C28', N'Coloration noire', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (35, 34, N'C29', N'Rinçage noir', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (36, 35, N'C30', N'Rinçage eau dure/imprégnation', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (37, 36, N'C31', N'Colmatage à chaud', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (38, 37, N'C32', N'Colmatage chaud', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (39, 38, N'C33', N'Conversion chimique', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (40, 39, N'C34', N'Rinçage totale', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (41, 40, N'C35', N'Rinçage totale', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (42, 41, N'D1', N'Poste de déchargement 1', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (43, 42, N'D2', N'Poste de déchargement 2', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (44, 201, N'NB', N'Niveau bas', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (45, 202, N'NI', N'Niveau intermédiaire', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (47, 215, N'NH', N'Niveau haut', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (48, 220, N'FMOH', N'Forcer la montée au niveau haut avec contrôle sur capteur', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (57, 230, N'FREFLEV', N'Forcer la référence du levage', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (66, 240, N'BP', N'Attente de l''appui sur un bouton poussoir', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (68, 260, N'TEMPO', N'Temporisation en SECONDES sur PARAMETRE', 1, N'Temps en SECONDES')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (69, 270, N'TEMPO_EGOUT', N'Temporisation d''égouttage en SECONDES sur PARAMETRE', 1, N'Temps en SECONDES')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (70, 280, N'TEMPO_STAB', N'Temporisation de stabilisation en SECONDES sur PARAMETRE', 1, N'Temps en SECONDES')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (71, 600, N'MOAC', N'Montée des accroches', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (72, 610, N'DEAC', N'Descente des accroches', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (73, 800, N'AEVPUL', N'Arrêt de l''électro-vanne de pulvérisation', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (75, 810, N'MEVPUL', N'Marche de l''électro-vanne de pulvérisation', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (76, 2000, N'LGAMRED', N'Lancer la gamme redresseur', 1, N'Numéro du poste ou se trouve le redresseur')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (78, 2010, N'AFINANOD', N'Attente de la fin de l''anodisation', 1, N'Numéro du poste ou se trouve le redresseur')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (79, 2020, N'AARRETRED', N'Attente de l''arrêt d''un redresseur', 1, N'Numéro du poste ou se trouve le redresseur')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (91, 2500, N'ASECHOIR', N'Arrêt du séchoir', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (92, 2510, N'MSECHOIR', N'Marche du séchoir', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (93, 8000, N'FCY', N'Fin de cycle', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (94, 0, N'NOP', N'Pas d''opération', 0, N'')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (95, 500, N'OCO', N'Ouverture des couvercles', 1, N'Numéro du poste ou se trouve les couvercles')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (96, 510, N'FCO', N'Fermeture des couvercles', 1, N'Numéro du poste ou se trouve les couvercles')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (97, 700, N'MAGIT', N'Marche de l''agitation', 1, N'Numéro du poste ou se trouve l''agitation')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (98, 710, N'AAGIT', N'Arrêt de l''agitation', 1, N'Numéro du poste ou se trouve l''agitation')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (99, 720, N'CTRLARRETAGIT', N'Contrôle de l''arrêt de l''agitation', 1, N'Numéro du poste ou se trouve l''agitation')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (104, 43, N'C37', N'etuve', 0, N'Etuve')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (105, 44, N'C38', N'basculeur', 0, N'Basculeur')
INSERT [dbo].[Actions] ([ClePrimaire], [NumAction], [CodeAction], [LibelleAction], [ParametreOuiNon], [LibelleParametre]) VALUES (106, 7, N'C02', N'Réserve', 0, N'')
SET IDENTITY_INSERT [dbo].[Actions] OFF
ALTER TABLE [dbo].[Actions] ADD  DEFAULT ((0)) FOR [NumAction]
GO
ALTER TABLE [dbo].[Actions] ADD  DEFAULT ('') FOR [CodeAction]
GO
ALTER TABLE [dbo].[Actions] ADD  DEFAULT ('') FOR [LibelleAction]
GO
ALTER TABLE [dbo].[Actions] ADD  DEFAULT ((0)) FOR [ParametreOuiNon]
GO
ALTER TABLE [dbo].[Actions] ADD  DEFAULT ('') FOR [LibelleParametre]
GO
