USE [ANODISATION]
GO
/****** Object:  Table [dbo].[Redresseurs]    Script Date: 21/10/2024 17:52:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Redresseurs](
	[NumRedresseur] [smallint] NOT NULL,
	[NomRedresseur] [varchar](20) NOT NULL,
	[LibelleRedresseur] [varchar](50) NOT NULL,
	[UMaxiRedresseur] [int] NOT NULL,
	[IMaxiRedresseur] [int] NOT NULL,
	[UMaxiProduction] [int] NOT NULL,
	[IMaxiProduction] [int] NOT NULL
) ON [PRIMARY]
GO
INSERT [dbo].[Redresseurs] ([NumRedresseur], [NomRedresseur], [LibelleRedresseur], [UMaxiRedresseur], [IMaxiRedresseur], [UMaxiProduction], [IMaxiProduction]) VALUES (1, N'Redresseur C13', N'Redresseur d''anodisation C13', 20, 3000, 20, 3000)
INSERT [dbo].[Redresseurs] ([NumRedresseur], [NomRedresseur], [LibelleRedresseur], [UMaxiRedresseur], [IMaxiRedresseur], [UMaxiProduction], [IMaxiProduction]) VALUES (2, N'Redresseur C14', N'Redresseur d''anodisation C14', 20, 3000, 20, 3000)
INSERT [dbo].[Redresseurs] ([NumRedresseur], [NomRedresseur], [LibelleRedresseur], [UMaxiRedresseur], [IMaxiRedresseur], [UMaxiProduction], [IMaxiProduction]) VALUES (3, N'Redresseur C15', N'Redresseur d''anodisation C15', 20, 3000, 20, 3000)
INSERT [dbo].[Redresseurs] ([NumRedresseur], [NomRedresseur], [LibelleRedresseur], [UMaxiRedresseur], [IMaxiRedresseur], [UMaxiProduction], [IMaxiProduction]) VALUES (4, N'Redresseur C16', N'Redresseur d''anodisation C16', 20, 3000, 20, 3000)
ALTER TABLE [dbo].[Redresseurs] ADD  DEFAULT ((0)) FOR [NumRedresseur]
GO
ALTER TABLE [dbo].[Redresseurs] ADD  DEFAULT ('') FOR [NomRedresseur]
GO
ALTER TABLE [dbo].[Redresseurs] ADD  DEFAULT ('') FOR [LibelleRedresseur]
GO
ALTER TABLE [dbo].[Redresseurs] ADD  DEFAULT ((0)) FOR [UMaxiRedresseur]
GO
ALTER TABLE [dbo].[Redresseurs] ADD  DEFAULT ((0)) FOR [IMaxiRedresseur]
GO
ALTER TABLE [dbo].[Redresseurs] ADD  DEFAULT ((0)) FOR [UMaxiProduction]
GO
ALTER TABLE [dbo].[Redresseurs] ADD  DEFAULT ((0)) FOR [IMaxiProduction]
GO
