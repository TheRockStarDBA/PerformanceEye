SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [CorePE].[QueryPlanStmtStore](
	[PKQueryPlanStmtStoreID] [bigint] IDENTITY(1,1) NOT NULL,
	[AWStmtPlanHash] [varbinary](64) NOT NULL,
	[plan_handle] [varbinary](64) NOT NULL,
	[statement_start_offset] [int] NOT NULL,
	[statement_end_offset] [int] NOT NULL,
	[dbid] [smallint] NOT NULL,
	[objectid] [int] NOT NULL,
	[fail_to_obtain] [bit] NOT NULL,
	[query_plan] [nvarchar](max) NOT NULL,
	[Insertedby_SPIDCaptureTime] [datetime] NOT NULL,
	[LastTouchedBy_SPIDCaptureTime] [datetime] NOT NULL,
 CONSTRAINT [PK_QueryPlanStmtStore] PRIMARY KEY CLUSTERED 
(
	[PKQueryPlanStmtStoreID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO
SET ANSI_PADDING OFF
GO
SET ANSI_PADDING ON

GO
CREATE UNIQUE NONCLUSTERED INDEX [AK_AWStmtPlanHash_handle_offsets] ON [CorePE].[QueryPlanStmtStore]
(
	[AWStmtPlanHash] ASC,
	[plan_handle] ASC,
	[statement_start_offset] ASC,
	[statement_end_offset] ASC
)
INCLUDE ( 	[dbid],
	[objectid]) WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO
SET ANSI_PADDING ON

GO
CREATE NONCLUSTERED INDEX [ncl_LastTouched] ON [CorePE].[QueryPlanStmtStore]
(
	[LastTouchedBy_SPIDCaptureTime] ASC
)
INCLUDE ( 	[PKQueryPlanStmtStoreID],
	[plan_handle],
	[statement_start_offset],
	[statement_end_offset],
	[dbid],
	[objectid]) WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO
