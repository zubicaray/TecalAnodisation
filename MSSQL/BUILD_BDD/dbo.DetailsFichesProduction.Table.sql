USE [ANODISATION]
GO
/****** Object:  Table [dbo].[DetailsFichesProduction]    Script Date: 21/10/2024 17:52:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[DetailsFichesProduction](
	[ClePrimaire] [int] IDENTITY(1,1) NOT NULL,
	[NumFicheProduction] [varchar](8) NOT NULL,
	[NumLigne] [tinyint] NOT NULL,
	[NumPoste] [smallint] NOT NULL,
	[DateEntreePoste] [datetime] NOT NULL,
	[DateSortiePoste] [datetime] NOT NULL,
	[DateDebutEgouttage] [datetime] NOT NULL,
	[DateFinEgouttage] [datetime] NOT NULL,
	[TemperatureEnEntree] [float] NOT NULL,
	[TemperatureEnSortie] [float] NOT NULL,
	[GrapheTemperature] [text] NOT NULL,
	[URedresseur] [float] NOT NULL,
	[IRedresseur] [float] NOT NULL,
	[SensRedresseur] [smallint] NOT NULL,
	[GrapheRedresseur] [text] NOT NULL,
	[AnalyseurEnEntree] [float] NOT NULL,
	[AnalyseurEnSortie] [float] NOT NULL,
	[GrapheAnalyseur] [text] NOT NULL,
	[AlarmesPoste] [text] NOT NULL,
 CONSTRAINT [PK_DetailsFichesProduction_ID] PRIMARY KEY CLUSTERED 
(
	[ClePrimaire] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
ALTER TABLE [dbo].[DetailsFichesProduction] ADD  DEFAULT ('') FOR [NumFicheProduction]
GO
ALTER TABLE [dbo].[DetailsFichesProduction] ADD  DEFAULT ((0)) FOR [NumLigne]
GO
ALTER TABLE [dbo].[DetailsFichesProduction] ADD  DEFAULT ((0)) FOR [NumPoste]
GO
ALTER TABLE [dbo].[DetailsFichesProduction] ADD  DEFAULT (getdate()) FOR [DateEntreePoste]
GO
ALTER TABLE [dbo].[DetailsFichesProduction] ADD  DEFAULT (getdate()) FOR [DateSortiePoste]
GO
ALTER TABLE [dbo].[DetailsFichesProduction] ADD  DEFAULT (getdate()) FOR [DateDebutEgouttage]
GO
ALTER TABLE [dbo].[DetailsFichesProduction] ADD  DEFAULT (getdate()) FOR [DateFinEgouttage]
GO
ALTER TABLE [dbo].[DetailsFichesProduction] ADD  DEFAULT ((0)) FOR [TemperatureEnEntree]
GO
ALTER TABLE [dbo].[DetailsFichesProduction] ADD  DEFAULT ((0)) FOR [TemperatureEnSortie]
GO
ALTER TABLE [dbo].[DetailsFichesProduction] ADD  DEFAULT ('') FOR [GrapheTemperature]
GO
ALTER TABLE [dbo].[DetailsFichesProduction] ADD  DEFAULT ((0)) FOR [URedresseur]
GO
ALTER TABLE [dbo].[DetailsFichesProduction] ADD  DEFAULT ((0)) FOR [IRedresseur]
GO
ALTER TABLE [dbo].[DetailsFichesProduction] ADD  DEFAULT ((0)) FOR [SensRedresseur]
GO
ALTER TABLE [dbo].[DetailsFichesProduction] ADD  DEFAULT ('') FOR [GrapheRedresseur]
GO
ALTER TABLE [dbo].[DetailsFichesProduction] ADD  DEFAULT ((0)) FOR [AnalyseurEnEntree]
GO
ALTER TABLE [dbo].[DetailsFichesProduction] ADD  DEFAULT ((0)) FOR [AnalyseurEnSortie]
GO
ALTER TABLE [dbo].[DetailsFichesProduction] ADD  DEFAULT ('') FOR [GrapheAnalyseur]
GO
ALTER TABLE [dbo].[DetailsFichesProduction] ADD  DEFAULT ('') FOR [AlarmesPoste]
GO
