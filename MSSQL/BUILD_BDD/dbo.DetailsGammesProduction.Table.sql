USE [ANODISATION]
GO
/****** Object:  Table [dbo].[DetailsGammesProduction]    Script Date: 21/10/2024 17:52:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[DetailsGammesProduction](
	[ClePrimaire] [int] IDENTITY(1,1) NOT NULL,
	[NumFicheProduction] [varchar](8) NOT NULL,
	[NumLigne] [tinyint] NOT NULL,
	[NumZone] [smallint] NOT NULL,
	[TempsAuPosteTexte] [varchar](12) NOT NULL,
	[TempsEgouttageTexte] [varchar](5) NOT NULL,
	[TempsAuPosteSecondes] [int] NOT NULL,
	[TempsEgouttageSecondes] [smallint] NOT NULL,
	[DecompteDuTempsAuPosteReelSecondes] [varchar](8) NOT NULL,
	[NumPosteReel] [smallint] NOT NULL,
 CONSTRAINT [PK_DetailsGammesProduction_ID] PRIMARY KEY CLUSTERED 
(
	[ClePrimaire] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[DetailsGammesProduction] ADD  DEFAULT ('') FOR [NumFicheProduction]
GO
ALTER TABLE [dbo].[DetailsGammesProduction] ADD  DEFAULT ((0)) FOR [NumLigne]
GO
ALTER TABLE [dbo].[DetailsGammesProduction] ADD  DEFAULT ((0)) FOR [NumZone]
GO
ALTER TABLE [dbo].[DetailsGammesProduction] ADD  DEFAULT ('') FOR [TempsAuPosteTexte]
GO
ALTER TABLE [dbo].[DetailsGammesProduction] ADD  DEFAULT ('') FOR [TempsEgouttageTexte]
GO
ALTER TABLE [dbo].[DetailsGammesProduction] ADD  DEFAULT ((0)) FOR [TempsAuPosteSecondes]
GO
ALTER TABLE [dbo].[DetailsGammesProduction] ADD  DEFAULT ((0)) FOR [TempsEgouttageSecondes]
GO
ALTER TABLE [dbo].[DetailsGammesProduction] ADD  DEFAULT ('') FOR [DecompteDuTempsAuPosteReelSecondes]
GO
ALTER TABLE [dbo].[DetailsGammesProduction] ADD  DEFAULT ((0)) FOR [NumPosteReel]
GO
