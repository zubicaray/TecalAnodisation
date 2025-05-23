USE [ANODISATION]
GO
/****** Object:  Table [dbo].[DetailsGammesAnodisation]    Script Date: 21/10/2024 17:52:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[DetailsGammesAnodisation](
	[ClePrimaire] [int] IDENTITY(1,1) NOT NULL,
	[NumGamme] [varchar](6) NOT NULL,
	[NumLigne] [smallint] NOT NULL,
	[NumZone] [smallint] NOT NULL,
	[TempsAuPosteTexte] [varchar](8) NOT NULL,
	[TempsAlerteTexte] [varchar](8) NOT NULL,
	[TempsEgouttageTexte] [varchar](5) NOT NULL,
	[TempsAuPosteSecondes] [int] NOT NULL,
	[TempsAlerteSecondes] [int] NOT NULL,
	[TempsEgouttageSecondes] [smallint] NOT NULL,
 CONSTRAINT [PK_DetailsGammesAnodisation2] PRIMARY KEY CLUSTERED 
(
	[ClePrimaire] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[DetailsGammesAnodisation] ADD  DEFAULT ('') FOR [NumGamme]
GO
ALTER TABLE [dbo].[DetailsGammesAnodisation] ADD  DEFAULT ((0)) FOR [NumLigne]
GO
ALTER TABLE [dbo].[DetailsGammesAnodisation] ADD  DEFAULT ((0)) FOR [NumZone]
GO
ALTER TABLE [dbo].[DetailsGammesAnodisation] ADD  DEFAULT ('') FOR [TempsAuPosteTexte]
GO
ALTER TABLE [dbo].[DetailsGammesAnodisation] ADD  DEFAULT ('') FOR [TempsAlerteTexte]
GO
ALTER TABLE [dbo].[DetailsGammesAnodisation] ADD  DEFAULT ('') FOR [TempsEgouttageTexte]
GO
ALTER TABLE [dbo].[DetailsGammesAnodisation] ADD  DEFAULT ((0)) FOR [TempsAuPosteSecondes]
GO
ALTER TABLE [dbo].[DetailsGammesAnodisation] ADD  DEFAULT ((0)) FOR [TempsAlerteSecondes]
GO
ALTER TABLE [dbo].[DetailsGammesAnodisation] ADD  DEFAULT ((0)) FOR [TempsEgouttageSecondes]
GO
