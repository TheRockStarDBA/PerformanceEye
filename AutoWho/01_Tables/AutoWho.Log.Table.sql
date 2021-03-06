SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [AutoWho].[Log](
	[LogDT] [datetime2](7) NOT NULL,
	[TraceID] [int] NULL,
	[ErrorCode] [int] NOT NULL,
	[LocationTag] [nvarchar](50) NOT NULL,
	[LogMessage] [nvarchar](max) NOT NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO
CREATE CLUSTERED INDEX [Clus_SnapshotDT] ON [AutoWho].[Log]
(
	[LogDT] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO
