SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [CorePE].[CaptureOrdinalCache](
	[Utility] [nvarchar](30) NOT NULL,
	[StartTime] [datetime] NOT NULL,
	[EndTime] [datetime] NOT NULL,
	[Ordinal] [int] NOT NULL,
	[OrdinalNegative] [int] NOT NULL,
	[CaptureTime] [datetime] NOT NULL,
	[TimePopulated] [datetime] NOT NULL,
 CONSTRAINT [PK_CaptureOrdinalCache] PRIMARY KEY CLUSTERED 
(
	[Utility] ASC,
	[StartTime] ASC,
	[EndTime] ASC,
	[Ordinal] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
CREATE UNIQUE NONCLUSTERED INDEX [NCL_OrdinalNegative] ON [CorePE].[CaptureOrdinalCache]
(
	[StartTime] ASC,
	[EndTime] ASC,
	[OrdinalNegative] ASC
)
INCLUDE ( 	[Ordinal],
	[CaptureTime],
	[TimePopulated]) WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, 
		IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO
ALTER TABLE [CorePE].[CaptureOrdinalCache] ADD  CONSTRAINT [DF_AutoWho_CaptureOrdinalCache_TimePopulated]  DEFAULT (getdate()) FOR [TimePopulated]
GO
