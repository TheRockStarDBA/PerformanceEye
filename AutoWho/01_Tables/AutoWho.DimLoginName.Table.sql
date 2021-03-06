SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [AutoWho].[DimLoginName](
	[DimLoginNameID] [smallint] IDENTITY(30,1) NOT NULL,
	[login_name] [nvarchar](128) NOT NULL,
	[original_login_name] [nvarchar](128) NOT NULL,
	[TimeAdded] [datetime] NOT NULL,
 CONSTRAINT [PK_DimLoginName] PRIMARY KEY CLUSTERED 
(
	[DimLoginNameID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
SET ANSI_PADDING ON

GO
CREATE UNIQUE NONCLUSTERED INDEX [AK_allattributes] ON [AutoWho].[DimLoginName]
(
	[login_name] ASC,
	[original_login_name] ASC
)
INCLUDE ( 	[DimLoginNameID],
	[TimeAdded]) WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO
ALTER TABLE [AutoWho].[DimLoginName] ADD  CONSTRAINT [DF_DimLoginName_TimeAdded]  DEFAULT (getdate()) FOR [TimeAdded]
GO
