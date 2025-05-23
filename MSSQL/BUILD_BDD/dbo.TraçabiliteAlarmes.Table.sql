USE [ANODISATION]
GO
/****** Object:  Table [dbo].[TraçabiliteAlarmes]    Script Date: 21/10/2024 17:52:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[TraçabiliteAlarmes](
	[ClePrimaire] [int] IDENTITY(1,1) NOT NULL,
	[NumDefaut] [int] NOT NULL,
	[ComplementDefaut] [varchar](30) NOT NULL,
	[DateDetectionDefaut] [datetime] NULL,
	[DateCorrectionDefaut] [datetime] NULL
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[TraçabiliteAlarmes] ADD  DEFAULT ((0)) FOR [NumDefaut]
GO
ALTER TABLE [dbo].[TraçabiliteAlarmes] ADD  DEFAULT ('') FOR [ComplementDefaut]
GO
