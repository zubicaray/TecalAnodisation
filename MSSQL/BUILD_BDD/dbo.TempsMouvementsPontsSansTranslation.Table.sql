USE [ANODISATION]
GO
/****** Object:  Table [dbo].[TempsMouvementsPontsSansTranslation]    Script Date: 21/10/2024 17:52:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[TempsMouvementsPontsSansTranslation](
	[ClePrimaire] [int] IDENTITY(1,1) NOT NULL,
	[NumPont] [smallint] NOT NULL,
	[TempsAccrochesChargeVersHaut] [float] NOT NULL,
	[TempsAccrochesChargeVersBas] [float] NOT NULL,
	[TempsDescenteHautVersBas] [float] NOT NULL,
	[TempsDescenteIntermediaireVersBas] [float] NOT NULL,
	[TempsMonteeBasVersIntermediaire] [float] NOT NULL,
	[TempsMonteeBasVersHaut] [float] NOT NULL
) ON [PRIMARY]
GO
SET IDENTITY_INSERT [dbo].[TempsMouvementsPontsSansTranslation] ON 

INSERT [dbo].[TempsMouvementsPontsSansTranslation] ([ClePrimaire], [NumPont], [TempsAccrochesChargeVersHaut], [TempsAccrochesChargeVersBas], [TempsDescenteHautVersBas], [TempsDescenteIntermediaireVersBas], [TempsMonteeBasVersIntermediaire], [TempsMonteeBasVersHaut]) VALUES (1, 1, 2, 1, 14, 5, 3, 16)
INSERT [dbo].[TempsMouvementsPontsSansTranslation] ([ClePrimaire], [NumPont], [TempsAccrochesChargeVersHaut], [TempsAccrochesChargeVersBas], [TempsDescenteHautVersBas], [TempsDescenteIntermediaireVersBas], [TempsMonteeBasVersIntermediaire], [TempsMonteeBasVersHaut]) VALUES (2, 2, 2, 1, 13, 5, 2, 16)
SET IDENTITY_INSERT [dbo].[TempsMouvementsPontsSansTranslation] OFF
ALTER TABLE [dbo].[TempsMouvementsPontsSansTranslation] ADD  DEFAULT ((0)) FOR [NumPont]
GO
ALTER TABLE [dbo].[TempsMouvementsPontsSansTranslation] ADD  DEFAULT ((0)) FOR [TempsAccrochesChargeVersHaut]
GO
ALTER TABLE [dbo].[TempsMouvementsPontsSansTranslation] ADD  DEFAULT ((0)) FOR [TempsAccrochesChargeVersBas]
GO
ALTER TABLE [dbo].[TempsMouvementsPontsSansTranslation] ADD  DEFAULT ((0)) FOR [TempsDescenteHautVersBas]
GO
ALTER TABLE [dbo].[TempsMouvementsPontsSansTranslation] ADD  DEFAULT ((0)) FOR [TempsDescenteIntermediaireVersBas]
GO
ALTER TABLE [dbo].[TempsMouvementsPontsSansTranslation] ADD  DEFAULT ((0)) FOR [TempsMonteeBasVersIntermediaire]
GO
ALTER TABLE [dbo].[TempsMouvementsPontsSansTranslation] ADD  DEFAULT ((0)) FOR [TempsMonteeBasVersHaut]
GO
