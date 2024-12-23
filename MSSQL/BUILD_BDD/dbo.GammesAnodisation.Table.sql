USE [ANODISATION]
GO
/****** Object:  Table [dbo].[GammesAnodisation]    Script Date: 21/10/2024 17:52:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[GammesAnodisation](
	[NumGamme] [varchar](6) NOT NULL,
	[RefGamme] [varchar](30) NOT NULL,
	[DateCreationGamme] [datetime] NOT NULL,
	[NomGamme] [varchar](50) NOT NULL,
	[Designation] [varchar](255) NOT NULL,
	[PassageAnodisation] [bit] NOT NULL,
	[PassageSpectro] [bit] NOT NULL,
	[PassageOr] [bit] NOT NULL,
	[PassageNoir] [bit] NOT NULL,
	[Matiere1] [varchar](30) NOT NULL,
	[Matiere2] [varchar](30) NOT NULL,
	[Matiere3] [varchar](30) NOT NULL,
	[Matiere4] [varchar](30) NOT NULL,
	[Matiere5] [varchar](30) NOT NULL,
	[Matiere6] [varchar](30) NOT NULL,
	[Matiere7] [varchar](30) NOT NULL,
	[Matiere8] [varchar](30) NOT NULL,
	[Matiere9] [varchar](30) NOT NULL,
	[Matiere10] [varchar](30) NOT NULL,
	[TempsAvantPostePrincipalTexte] [varchar](8) NOT NULL,
	[TempsPostePrincipalTexte] [varchar](8) NOT NULL,
	[TempsApresPostePrincipalTexte] [varchar](8) NOT NULL,
	[TempsTotalPostesTexte] [varchar](8) NOT NULL,
	[TempsTotalEgouttagesTexte] [varchar](8) NOT NULL,
	[TempsTotalGammeTexte] [varchar](8) NOT NULL,
	[TempsAvantPostePrincipalSecondes] [int] NOT NULL,
	[TempsPostePrincipalSecondes] [int] NOT NULL,
	[TempsApresPostePrincipalSecondes] [int] NOT NULL,
	[TempsTotalPostesSecondes] [int] NOT NULL,
	[TempsTotalEgouttagesSecondes] [int] NOT NULL,
	[TempsTotalGammeSecondes] [int] NOT NULL,
	[ModeUouI] [tinyint] NOT NULL,
	[TempsPhase1] [smallint] NOT NULL,
	[UPhase1] [float] NOT NULL,
	[IPhase1] [float] NOT NULL,
	[TempsPhase2] [smallint] NOT NULL,
	[UPhase2] [float] NOT NULL,
	[IPhase2] [float] NOT NULL,
	[TempsPhase3] [smallint] NOT NULL,
	[UPhase3] [float] NOT NULL,
	[IPhase3] [float] NOT NULL,
	[TempsPhase4] [smallint] NOT NULL,
	[UPhase4] [float] NOT NULL,
	[IPhase4] [float] NOT NULL,
 CONSTRAINT [PK_GammesAnodisation2] PRIMARY KEY CLUSTERED 
(
	[NumGamme] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ('') FOR [NumGamme]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ('') FOR [RefGamme]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT (getdate()) FOR [DateCreationGamme]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ('') FOR [NomGamme]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ('') FOR [Designation]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ((0)) FOR [PassageAnodisation]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ((0)) FOR [PassageSpectro]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ((0)) FOR [PassageOr]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ((0)) FOR [PassageNoir]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ('') FOR [Matiere1]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ('') FOR [Matiere2]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ('') FOR [Matiere3]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ('') FOR [Matiere4]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ('') FOR [Matiere5]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ('') FOR [Matiere6]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ('') FOR [Matiere7]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ('') FOR [Matiere8]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ('') FOR [Matiere9]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ('') FOR [Matiere10]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ('') FOR [TempsAvantPostePrincipalTexte]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ('') FOR [TempsPostePrincipalTexte]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ('') FOR [TempsApresPostePrincipalTexte]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ('') FOR [TempsTotalPostesTexte]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ('') FOR [TempsTotalEgouttagesTexte]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ('') FOR [TempsTotalGammeTexte]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ((0)) FOR [TempsAvantPostePrincipalSecondes]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ((0)) FOR [TempsPostePrincipalSecondes]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ((0)) FOR [TempsApresPostePrincipalSecondes]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ((0)) FOR [TempsTotalPostesSecondes]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ((0)) FOR [TempsTotalEgouttagesSecondes]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ((0)) FOR [TempsTotalGammeSecondes]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ((0)) FOR [ModeUouI]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ((0)) FOR [TempsPhase1]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ((0)) FOR [UPhase1]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ((0)) FOR [IPhase1]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ((0)) FOR [TempsPhase2]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ((0)) FOR [UPhase2]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ((0)) FOR [IPhase2]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ((0)) FOR [TempsPhase3]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ((0)) FOR [UPhase3]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ((0)) FOR [IPhase3]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ((0)) FOR [TempsPhase4]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ((0)) FOR [UPhase4]
GO
ALTER TABLE [dbo].[GammesAnodisation] ADD  DEFAULT ((0)) FOR [IPhase4]
GO
