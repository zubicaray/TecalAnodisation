USE [ANODISATION]
GO
/****** Object:  Table [dbo].[DetailsChargesProduction]    Script Date: 21/10/2024 17:52:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[DetailsChargesProduction](
	[ClePrimaire] [int] IDENTITY(1,1) NOT NULL,
	[NumCommandeInterne] [varchar](8) NOT NULL,
	[NbrReparations] [varchar](1) NOT NULL,
	[DateEntreeEnLigne] [datetime] NOT NULL,
	[DateArriveeAuDechargement] [datetime] NOT NULL,
	[NumBarre] [smallint] NOT NULL,
	[NumLigne] [tinyint] NOT NULL,
	[CodeClient] [varchar](30) NOT NULL,
	[NbrPieces] [numeric](10, 0) NOT NULL,
	[Designation] [varchar](255) NOT NULL,
	[NumLignesReferencesClient] [varchar](50) NOT NULL,
	[Matiere] [varchar](30) NOT NULL,
	[NumGammeAnodisation] [varchar](6) NOT NULL,
	[RefGammeAnodisation] [varchar](18) NOT NULL,
	[TempsAnodisationTexte] [varchar](8) NOT NULL,
	[NumFicheProduction] [varchar](8) NOT NULL,
	[ChargePrioritaire] [smallint] NOT NULL,
	[AlarmesLigne] [text] NOT NULL,
	[ControleColmatage] [smallint] NOT NULL,
	[ControleEpaisseurAnodisation] [smallint] NOT NULL,
	[ControleColoration] [varchar](20) NOT NULL,
	[ControleObservations] [varchar](50) NOT NULL,
 CONSTRAINT [PK_DetailsChargesProduction_ID] PRIMARY KEY CLUSTERED 
(
	[ClePrimaire] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
ALTER TABLE [dbo].[DetailsChargesProduction] ADD  DEFAULT ('') FOR [NumCommandeInterne]
GO
ALTER TABLE [dbo].[DetailsChargesProduction] ADD  DEFAULT ('') FOR [NbrReparations]
GO
ALTER TABLE [dbo].[DetailsChargesProduction] ADD  DEFAULT (getdate()) FOR [DateEntreeEnLigne]
GO
ALTER TABLE [dbo].[DetailsChargesProduction] ADD  DEFAULT (getdate()) FOR [DateArriveeAuDechargement]
GO
ALTER TABLE [dbo].[DetailsChargesProduction] ADD  DEFAULT ((0)) FOR [NumBarre]
GO
ALTER TABLE [dbo].[DetailsChargesProduction] ADD  DEFAULT ((0)) FOR [NumLigne]
GO
ALTER TABLE [dbo].[DetailsChargesProduction] ADD  DEFAULT ('') FOR [CodeClient]
GO
ALTER TABLE [dbo].[DetailsChargesProduction] ADD  DEFAULT ('') FOR [Designation]
GO
ALTER TABLE [dbo].[DetailsChargesProduction] ADD  DEFAULT ('') FOR [NumLignesReferencesClient]
GO
ALTER TABLE [dbo].[DetailsChargesProduction] ADD  DEFAULT ('') FOR [Matiere]
GO
ALTER TABLE [dbo].[DetailsChargesProduction] ADD  DEFAULT ('') FOR [NumGammeAnodisation]
GO
ALTER TABLE [dbo].[DetailsChargesProduction] ADD  DEFAULT ('') FOR [RefGammeAnodisation]
GO
ALTER TABLE [dbo].[DetailsChargesProduction] ADD  DEFAULT ('') FOR [TempsAnodisationTexte]
GO
ALTER TABLE [dbo].[DetailsChargesProduction] ADD  DEFAULT ('') FOR [NumFicheProduction]
GO
ALTER TABLE [dbo].[DetailsChargesProduction] ADD  DEFAULT ((0)) FOR [ChargePrioritaire]
GO
ALTER TABLE [dbo].[DetailsChargesProduction] ADD  DEFAULT ('') FOR [AlarmesLigne]
GO
ALTER TABLE [dbo].[DetailsChargesProduction] ADD  DEFAULT ((0)) FOR [ControleColmatage]
GO
ALTER TABLE [dbo].[DetailsChargesProduction] ADD  DEFAULT ((0)) FOR [ControleEpaisseurAnodisation]
GO
ALTER TABLE [dbo].[DetailsChargesProduction] ADD  DEFAULT ('') FOR [ControleColoration]
GO
ALTER TABLE [dbo].[DetailsChargesProduction] ADD  DEFAULT ('') FOR [ControleObservations]
GO
