USE [ANODISATION]
GO
/****** Object:  Table [dbo].[DetailsPhasesProduction]    Script Date: 21/10/2024 17:52:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[DetailsPhasesProduction](
	[ClePrimaire] [int] IDENTITY(1,1) NOT NULL,
	[NumFicheProduction] [varchar](8) NOT NULL,
	[NumRedresseur] [tinyint] NOT NULL,
	[ModeUouI] [tinyint] NOT NULL,
	[NumPhase] [tinyint] NOT NULL,
	[TempsPhase] [smallint] NOT NULL,
	[UPhase] [float] NOT NULL,
	[IPhase] [float] NOT NULL,
 CONSTRAINT [PK_DetailsPhasesProduction_ID] PRIMARY KEY CLUSTERED 
(
	[ClePrimaire] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[DetailsPhasesProduction] ADD  DEFAULT ('') FOR [NumFicheProduction]
GO
ALTER TABLE [dbo].[DetailsPhasesProduction] ADD  DEFAULT ((0)) FOR [NumRedresseur]
GO
ALTER TABLE [dbo].[DetailsPhasesProduction] ADD  DEFAULT ((0)) FOR [ModeUouI]
GO
ALTER TABLE [dbo].[DetailsPhasesProduction] ADD  DEFAULT ((0)) FOR [NumPhase]
GO
ALTER TABLE [dbo].[DetailsPhasesProduction] ADD  DEFAULT ((0)) FOR [TempsPhase]
GO
ALTER TABLE [dbo].[DetailsPhasesProduction] ADD  DEFAULT ((0)) FOR [UPhase]
GO
ALTER TABLE [dbo].[DetailsPhasesProduction] ADD  DEFAULT ((0)) FOR [IPhase]
GO
