USE [ANODISATION]
GO
/****** Object:  Table [dbo].[ListeDefauts]    Script Date: 21/10/2024 17:52:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ListeDefauts](
	[NumDefaut] [smallint] NOT NULL,
	[SignalerOuiNon] [smallint] NOT NULL,
	[GyrophareOuiNon] [smallint] NOT NULL,
	[KlaxonOuiNon] [smallint] NOT NULL,
	[MessageVocalOuiNon] [smallint] NOT NULL,
	[AfficheurOuiNon] [smallint] NOT NULL,
	[InformationsAPI] [varchar](30) NOT NULL,
	[LibelleDefaut] [varchar](100) NOT NULL,
	[LibelleDefautAfficheur] [varchar](100) NOT NULL,
	[NumIntervenant1] [smallint] NOT NULL,
	[NumIntervenant2] [smallint] NOT NULL,
	[NumIntervenant3] [smallint] NOT NULL,
	[NumIntervenant4] [smallint] NOT NULL,
	[NumIntervenant5] [smallint] NOT NULL
) ON [PRIMARY]

GO
ALTER TABLE [dbo].[ListeDefauts] ADD  DEFAULT ((0)) FOR [NumDefaut]
GO
ALTER TABLE [dbo].[ListeDefauts] ADD  DEFAULT ((0)) FOR [SignalerOuiNon]
GO
ALTER TABLE [dbo].[ListeDefauts] ADD  DEFAULT ((0)) FOR [GyrophareOuiNon]
GO
ALTER TABLE [dbo].[ListeDefauts] ADD  DEFAULT ((0)) FOR [KlaxonOuiNon]
GO
ALTER TABLE [dbo].[ListeDefauts] ADD  DEFAULT ((0)) FOR [MessageVocalOuiNon]
GO
ALTER TABLE [dbo].[ListeDefauts] ADD  DEFAULT ((0)) FOR [AfficheurOuiNon]
GO
ALTER TABLE [dbo].[ListeDefauts] ADD  DEFAULT ('') FOR [InformationsAPI]
GO
ALTER TABLE [dbo].[ListeDefauts] ADD  DEFAULT ('') FOR [LibelleDefaut]
GO
ALTER TABLE [dbo].[ListeDefauts] ADD  DEFAULT ('') FOR [LibelleDefautAfficheur]
GO
ALTER TABLE [dbo].[ListeDefauts] ADD  DEFAULT ((0)) FOR [NumIntervenant1]
GO
ALTER TABLE [dbo].[ListeDefauts] ADD  DEFAULT ((0)) FOR [NumIntervenant2]
GO
ALTER TABLE [dbo].[ListeDefauts] ADD  DEFAULT ((0)) FOR [NumIntervenant3]
GO
ALTER TABLE [dbo].[ListeDefauts] ADD  DEFAULT ((0)) FOR [NumIntervenant4]
GO
ALTER TABLE [dbo].[ListeDefauts] ADD  DEFAULT ((0)) FOR [NumIntervenant5]
GO
