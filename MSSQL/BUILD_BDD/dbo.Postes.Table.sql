USE [ANODISATION]
GO
/****** Object:  Table [dbo].[Postes]    Script Date: 21/10/2024 17:52:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Postes](
	[NumPoste] [smallint] NOT NULL,
	[NomPoste] [varchar](6) NOT NULL,
	[LibellePoste] [varchar](50) NOT NULL,
	[AvecTemps] [bit] NOT NULL,
	[RespectTempsObligatoire] [bit] NOT NULL,
	[AvecEgouttage] [bit] NOT NULL,
	[PresenceCouvercles] [bit] NOT NULL,
	[PresenceRedresseur] [bit] NOT NULL,
	[PresenceAgitationBain] [bit] NOT NULL,
	[XAxePosteLigne] [int] NOT NULL,
	[XAxePosteSynoptique] [int] NOT NULL,
	[XInferieurPosteSynoptique] [int] NOT NULL,
	[YInferieurPosteSynoptique] [int] NOT NULL,
	[XSuperieurPosteSynoptique] [int] NOT NULL,
	[YSuperieurPosteSynoptique] [int] NOT NULL,
	[XInferieurLibellePosteSynoptique] [int] NOT NULL,
	[YInferieurLibellePosteSynoptique] [int] NOT NULL,
	[XSuperieurLibellePosteSynoptique] [int] NOT NULL,
	[YSuperieurLibellePosteSynoptique] [int] NOT NULL
) ON [PRIMARY]
GO
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (1, N'CHGT1', N'Poste de chargement 1', 0, 0, 0, 0, 0, 0, 4000, 1834, 1813, 95, 1855, 153, 1820, 162, 1850, 180)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (2, N'CHGT2', N'Poste de chargement 2', 0, 0, 0, 0, 0, 0, 5539, 1793, 1772, 95, 1814, 153, 1777, 162, 1807, 180)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (4, N'C02', N'Réserve', 0, 0, 0, 0, 0, 0, 8379, 1593, 1574, 95, 1616, 153, 1577, 163, 1609, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (5, N'C00', N'Dégraissage', 1, 0, 1, 0, 0, 0, 7412, 1717, 1691, 95, 1733, 153, 1702, 163, 1732, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (6, N'DEC', N'Décapage', 1, 0, 1, 0, 0, 0, 8991, 1674, 1652, 95, 1694, 153, 1659, 163, 1689, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (7, N'SAT', N'Satinage', 1, 0, 1, 0, 0, 0, 10690, 1634, 1602, 95, 1638, 153, 1618, 163, 1648, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (8, N'C03', N'Rinçage soude', 1, 0, 1, 0, 0, 0, 13866, 1551, 1530, 95, 1572, 153, 1536, 163, 1566, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (9, N'C04', N'Rinçage dégraissage', 1, 0, 1, 0, 0, 0, 15165, 1511, 1490, 95, 1532, 153, 1496, 163, 1526, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (10, N'C05', N'Dégraissage acide', 1, 1, 1, 0, 0, 0, 16496, 1469, 1448, 95, 1490, 153, 1454, 163, 1484, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (11, N'C06', N'Rinçage Mt brillantage', 1, 0, 1, 0, 0, 0, 18129, 1429, 1408, 95, 1450, 153, 1414, 163, 1444, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (12, N'C07', N'Brillantage', 1, 1, 1, 0, 0, 0, 19671, 1390, 1369, 95, 1411, 153, 1375, 163, 1405, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (13, N'C08', N'Rinçage brillantage', 1, 0, 1, 0, 0, 0, 21203, 1348, 1327, 95, 1369, 153, 1333, 163, 1363, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (14, N'C09', N'Rinçage brillantage', 1, 0, 1, 0, 0, 0, 22433, 1305, 1284, 95, 1326, 153, 1290, 163, 1320, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (15, N'C10', N'Neutralisation', 1, 0, 1, 0, 0, 0, 23718, 1264, 1243, 95, 1285, 153, 1249, 163, 1279, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (16, N'C11', N'Rinçage', 1, 0, 1, 0, 0, 0, 24937, 1224, 1203, 95, 1245, 153, 1209, 163, 1239, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (17, N'C12', N'Rinçage', 1, 1, 1, 0, 0, 0, 26154, 1183, 1162, 95, 1204, 153, 1168, 163, 1198, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (18, N'C13', N'Anodisation 1', 1, 1, 1, 0, 1, 0, 27682, 1143, 1122, 95, 1164, 153, 1128, 163, 1158, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (19, N'C14', N'Anodisation 2', 1, 1, 1, 0, 1, 0, 29349, 1102, 1081, 95, 1123, 153, 1087, 163, 1117, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (20, N'C15', N'Anodisation 3', 1, 1, 1, 0, 1, 0, 31008, 1061, 1040, 95, 1082, 153, 1046, 163, 1076, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (21, N'C16', N'Anodisation 4', 1, 0, 1, 0, 1, 0, 32587, 1017, 996, 95, 1038, 153, 1002, 163, 1032, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (22, N'C17', N'Rinçage anodisation', 1, 1, 1, 0, 0, 0, 34187, 977, 956, 95, 998, 153, 962, 163, 992, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (23, N'C18', N'Rinçage anodisation', 1, 1, 1, 0, 0, 0, 35455, 936, 915, 95, 957, 153, 921, 163, 951, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (24, N'C19', N'Spectrocoloration', 1, 0, 1, 0, 0, 0, 36833, 895, 874, 95, 916, 153, 880, 163, 910, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (25, N'C20', N'Rinçage ', 1, 0, 1, 0, 0, 0, 38149, 855, 834, 95, 876, 153, 840, 163, 870, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (26, N'C21', N'Rinçage ', 1, 0, 1, 0, 0, 0, 39373, 814, 793, 95, 835, 153, 799, 163, 829, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (27, N'C22', N'Coloration or', 1, 0, 1, 0, 0, 0, 40805, 771, 750, 95, 792, 153, 756, 163, 786, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (28, N'C23', N'Coloration orange', 1, 0, 1, 0, 0, 0, 42000, 732, 711, 95, 753, 153, 717, 163, 747, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (29, N'C24', N'RESERVE 2', 1, 0, 1, 0, 0, 0, 42937, 689, 668, 95, 710, 153, 674, 163, 704, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (30, N'C25', N'Imprégnation à froid', 1, 0, 1, 0, 0, 0, 43963, 649, 628, 95, 670, 153, 634, 163, 664, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (31, N'C26', N'Rinçage ', 1, 0, 1, 0, 0, 0, 45525, 607, 586, 95, 628, 153, 592, 163, 622, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (32, N'C27', N'Imprégnation à froid', 1, 0, 1, 0, 0, 0, 47020, 566, 545, 95, 587, 153, 551, 163, 581, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (33, N'C28', N'Coloration noire', 1, 0, 1, 0, 0, 0, 48466, 524, 503, 95, 545, 153, 509, 163, 539, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (34, N'C29', N'Rinçage noir', 1, 0, 1, 0, 0, 0, 49938, 483, 462, 95, 504, 153, 468, 163, 498, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (35, N'C30', N'Eau dure / imprégnation', 1, 0, 1, 0, 0, 0, 51323, 445, 424, 95, 466, 153, 430, 163, 460, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (36, N'C31', N'Colmatage à chaud', 1, 0, 1, 0, 0, 0, 53358, 402, 381, 95, 423, 153, 387, 163, 417, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (37, N'C32', N'Colmatage chaud', 1, 0, 1, 0, 0, 0, 54990, 361, 340, 95, 382, 153, 346, 163, 376, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (38, N'C33', N'Conversion chimique', 1, 1, 1, 0, 0, 0, 56636, 319, 298, 95, 340, 153, 305, 163, 335, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (39, N'C34', N'Rinçage totale', 1, 1, 1, 0, 0, 0, 57800, 278, 257, 95, 299, 153, 263, 163, 293, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (40, N'C35', N'Rinçage totale', 1, 1, 1, 0, 0, 0, 59233, 238, 217, 95, 259, 153, 223, 163, 253, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (41, N'D1', N'Poste de déchargement 1', 0, 0, 0, 0, 0, 0, 61404, 156, 135, 95, 177, 153, 145, 163, 174, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (42, N'D2', N'Poste de déchargement 2', 0, 0, 0, 0, 0, 0, 62925, 115, 94, 95, 136, 153, 104, 163, 130, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (43, N'C37', N'Etuve', 1, 0, 0, 0, 0, 0, 64990, 75, 54, 95, 96, 153, 61, 163, 91, 178)
INSERT [dbo].[Postes] ([NumPoste], [NomPoste], [LibellePoste], [AvecTemps], [RespectTempsObligatoire], [AvecEgouttage], [PresenceCouvercles], [PresenceRedresseur], [PresenceAgitationBain], [XAxePosteLigne], [XAxePosteSynoptique], [XInferieurPosteSynoptique], [YInferieurPosteSynoptique], [XSuperieurPosteSynoptique], [YSuperieurPosteSynoptique], [XInferieurLibellePosteSynoptique], [YInferieurLibellePosteSynoptique], [XSuperieurLibellePosteSynoptique], [YSuperieurLibellePosteSynoptique]) VALUES (44, N'C38', N'Basculeur', 1, 0, 0, 0, 0, 0, 66700, 34, 13, 95, 55, 153, 20, 163, 50, 178)
ALTER TABLE [dbo].[Postes] ADD  DEFAULT ((0)) FOR [NumPoste]
GO
ALTER TABLE [dbo].[Postes] ADD  DEFAULT ('') FOR [NomPoste]
GO
ALTER TABLE [dbo].[Postes] ADD  DEFAULT ('') FOR [LibellePoste]
GO
ALTER TABLE [dbo].[Postes] ADD  DEFAULT ((0)) FOR [AvecTemps]
GO
ALTER TABLE [dbo].[Postes] ADD  DEFAULT ((0)) FOR [RespectTempsObligatoire]
GO
ALTER TABLE [dbo].[Postes] ADD  DEFAULT ((0)) FOR [AvecEgouttage]
GO
ALTER TABLE [dbo].[Postes] ADD  DEFAULT ((0)) FOR [PresenceCouvercles]
GO
ALTER TABLE [dbo].[Postes] ADD  DEFAULT ((0)) FOR [PresenceRedresseur]
GO
ALTER TABLE [dbo].[Postes] ADD  DEFAULT ((0)) FOR [PresenceAgitationBain]
GO
ALTER TABLE [dbo].[Postes] ADD  DEFAULT ((0)) FOR [XAxePosteLigne]
GO
ALTER TABLE [dbo].[Postes] ADD  DEFAULT ((0)) FOR [XAxePosteSynoptique]
GO
ALTER TABLE [dbo].[Postes] ADD  DEFAULT ((0)) FOR [XInferieurPosteSynoptique]
GO
ALTER TABLE [dbo].[Postes] ADD  DEFAULT ((0)) FOR [YInferieurPosteSynoptique]
GO
ALTER TABLE [dbo].[Postes] ADD  DEFAULT ((0)) FOR [XSuperieurPosteSynoptique]
GO
ALTER TABLE [dbo].[Postes] ADD  DEFAULT ((0)) FOR [YSuperieurPosteSynoptique]
GO
ALTER TABLE [dbo].[Postes] ADD  DEFAULT ((0)) FOR [XInferieurLibellePosteSynoptique]
GO
ALTER TABLE [dbo].[Postes] ADD  DEFAULT ((0)) FOR [YInferieurLibellePosteSynoptique]
GO
ALTER TABLE [dbo].[Postes] ADD  DEFAULT ((0)) FOR [XSuperieurLibellePosteSynoptique]
GO
ALTER TABLE [dbo].[Postes] ADD  DEFAULT ((0)) FOR [YSuperieurLibellePosteSynoptique]
GO
