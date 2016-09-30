SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [AutoWho].[BlockingGraphs](
	[SPIDCaptureTime] [datetime] NOT NULL,
	[session_id] [smallint] NOT NULL,
	[request_id] [smallint] NOT NULL,
	[exec_context_id] [smallint] NULL,
	[calc__blocking_session_Id] [smallint] NULL,
	[wait_type] [nvarchar](60) NULL,
	[wait_duration_ms] [bigint] NULL,
	[resource_description] [nvarchar](500) NULL,
	[FKInputBufferStoreID] [bigint] NULL,
	[FKSQLStmtStoreID] [bigint] NULL,
	[sort_value] [nvarchar](400) NULL,
	[block_group] [smallint] NULL,
	[levelindc] [smallint] NOT NULL,
	[rn] [smallint] NOT NULL
) ON [PRIMARY]

GO
CREATE CLUSTERED INDEX [CL_SPIDCaptureTime] ON [AutoWho].[BlockingGraphs]
(
	[SPIDCaptureTime] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO
