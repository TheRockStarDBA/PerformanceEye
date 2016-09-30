CREATE TYPE [dbo].[CorePEFiltersType] AS TABLE(
	[FilterType] [tinyint] NOT NULL,
	[FilterID] [int] NOT NULL,
	[FilterName] [nvarchar](255) NULL
)
GO
