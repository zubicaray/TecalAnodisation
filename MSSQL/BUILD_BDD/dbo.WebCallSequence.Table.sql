USE [ANODISATION]
GO
/****** Object:  Table [dbo].[WebCallSequence]    Script Date: 21/10/2024 17:52:27 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[WebCallSequence](
	[WebCallSequenceID] [int] NOT NULL,
	[WebCallSequence] [bigint] NOT NULL,
	[jour] [datetime] NOT NULL
) ON [PRIMARY]
GO
INSERT [dbo].[WebCallSequence] ([WebCallSequenceID], [WebCallSequence], [jour]) VALUES (1, 41, CAST(N'2024-07-18T00:00:00.000' AS DateTime))
ALTER TABLE [dbo].[WebCallSequence] ADD  DEFAULT ((0)) FOR [WebCallSequenceID]
GO
ALTER TABLE [dbo].[WebCallSequence] ADD  DEFAULT ((0)) FOR [WebCallSequence]
GO
ALTER TABLE [dbo].[WebCallSequence] ADD  DEFAULT (getdate()) FOR [jour]
GO
