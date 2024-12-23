USE [ANODISATION]
GO
/****** Object:  Table [dbo].[OuiNon]    Script Date: 21/10/2024 17:52:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[OuiNon](
	[ClePrimaire] [int] IDENTITY(1,1) NOT NULL,
	[OuiNonNumerique] [smallint] NOT NULL,
	[OuiNonTexte] [varchar](3) NOT NULL
) ON [PRIMARY]
GO
SET IDENTITY_INSERT [dbo].[OuiNon] ON 

INSERT [dbo].[OuiNon] ([ClePrimaire], [OuiNonNumerique], [OuiNonTexte]) VALUES (1, 0, N'NON')
INSERT [dbo].[OuiNon] ([ClePrimaire], [OuiNonNumerique], [OuiNonTexte]) VALUES (2, 1, N'OUI')
INSERT [dbo].[OuiNon] ([ClePrimaire], [OuiNonNumerique], [OuiNonTexte]) VALUES (3, -1, N'OUI')
SET IDENTITY_INSERT [dbo].[OuiNon] OFF
ALTER TABLE [dbo].[OuiNon] ADD  DEFAULT ((0)) FOR [OuiNonNumerique]
GO
ALTER TABLE [dbo].[OuiNon] ADD  DEFAULT ('') FOR [OuiNonTexte]
GO
