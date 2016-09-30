SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TRIGGER [AutoWho].[trgUPD_AutoWhoOptions] ON [AutoWho].[Options]

FOR UPDATE
AS 	BEGIN

INSERT INTO AutoWho.Options_History 
(RowID, AutoWhoEnabled, BeginTime, EndTime, IntervalLength, IncludeIdleWithTran, IncludeIdleWithoutTran, DurationFilter, IncludeDBs, ExcludeDBs, HighTempDBThreshold, CollectSystemSpids, HideSelf, ObtainBatchText, ObtainQueryPlanForStatement, ObtainQueryPlanForBatch, ObtainLocksForBlockRelevantThreshold, InputBufferThreshold, ParallelWaitsThreshold, QueryPlanThreshold, QueryPlanThresholdBlockRel, BlockingChainThreshold, BlockingChainDepth, TranDetailsThreshold, MediumDurationThreshold, HighDurationThreshold, BatchDurationThreshold, LongTransactionThreshold, Retention_IdleSPIDs_NoTran, Retention_IdleSPIDs_WithShortTran, Retention_IdleSPIDs_WithLongTran, Retention_IdleSPIDs_HighTempDB, Retention_ActiveLow, Retention_ActiveMedium, Retention_ActiveHigh, Retention_ActiveBatch, Retention_CaptureTimes, DebugSpeed, ThresholdFilterRefresh, SaveBadDims, Enable8666, ResolvePageLatches, ResolveLockWaits, 
HistoryInsertDate,
TriggerAction,
LastModifiedUser)
SELECT 
RowID, AutoWhoEnabled, BeginTime, EndTime, IntervalLength, IncludeIdleWithTran, IncludeIdleWithoutTran, DurationFilter, IncludeDBs, ExcludeDBs, HighTempDBThreshold, CollectSystemSpids, HideSelf, ObtainBatchText, ObtainQueryPlanForStatement, ObtainQueryPlanForBatch, ObtainLocksForBlockRelevantThreshold, InputBufferThreshold, ParallelWaitsThreshold, QueryPlanThreshold, QueryPlanThresholdBlockRel, BlockingChainThreshold, BlockingChainDepth, TranDetailsThreshold, MediumDurationThreshold, HighDurationThreshold, BatchDurationThreshold, LongTransactionThreshold, Retention_IdleSPIDs_NoTran, Retention_IdleSPIDs_WithShortTran, Retention_IdleSPIDs_WithLongTran, Retention_IdleSPIDs_HighTempDB, Retention_ActiveLow, Retention_ActiveMedium, Retention_ActiveHigh, Retention_ActiveBatch, Retention_CaptureTimes, DebugSpeed, ThresholdFilterRefresh, SaveBadDims, Enable8666, ResolvePageLatches, ResolveLockWaits, 
getdate(),
'Update',
SUSER_SNAME()
FROM inserted

END
GO


