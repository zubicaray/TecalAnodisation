USE [ANODISATION]
GO
/****** Object:  Table [dbo].[TempsMouvementsTranslationPonts]    Script Date: 21/10/2024 17:52:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[TempsMouvementsTranslationPonts](
	[ClePrimaire] [int] IDENTITY(1,1) NOT NULL,
	[NumPont] [smallint] NOT NULL,
	[NumPosteDepart] [smallint] NOT NULL,
	[NumPosteArrivee] [smallint] NOT NULL,
	[TempsTranslation] [float] NOT NULL
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[TempsMouvementsTranslationPonts] ADD  DEFAULT ((0)) FOR [NumPont]
GO
ALTER TABLE [dbo].[TempsMouvementsTranslationPonts] ADD  DEFAULT ((0)) FOR [NumPosteDepart]
GO
ALTER TABLE [dbo].[TempsMouvementsTranslationPonts] ADD  DEFAULT ((0)) FOR [NumPosteArrivee]
GO
ALTER TABLE [dbo].[TempsMouvementsTranslationPonts] ADD  DEFAULT ((0)) FOR [TempsTranslation]
GO