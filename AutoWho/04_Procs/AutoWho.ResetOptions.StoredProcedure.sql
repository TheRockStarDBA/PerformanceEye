SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [AutoWho].[ResetOptions] 
/*   
	PROCEDURE:		AutoWho.ResetOptions

	AUTHOR:			Aaron Morelli
					email@TBD.com
					@sqlcrossjoin
					sqlcrossjoin.wordpress.com
					https://github.com/amorelli005/PerformanceEye

	PURPOSE: Deletes the row in AutoWho.Options and re-inserts a row based on default values

	OUTSTANDING ISSUES: None at this time.

    CHANGE LOG:	
				2016-09-09	Aaron Morelli		Code complete

	MIT License

	Copyright (c) 2016 Aaron Morelli

	Permission is hereby granted, free of charge, to any person obtaining a copy
	of this software and associated documentation files (the "Software"), to deal
	in the Software without restriction, including without limitation the rights
	to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
	copies of the Software, and to permit persons to whom the Software is
	furnished to do so, subject to the following conditions:

	The above copyright notice and this permission notice shall be included in all
	copies or substantial portions of the Software.

	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
	IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
	FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
	AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
	LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
	SOFTWARE.

To Execute
------------------------
EXEC AutoWho.ResetOptions

SELECT * FROM AutoWho.Options
*/
AS
BEGIN
	SET NOCOUNT ON;

	INSERT INTO AutoWho.Options_History
	(
	RowID, AutoWhoEnabled, BeginTime, EndTime, IntervalLength, IncludeIdleWithTran, IncludeIdleWithoutTran, DurationFilter, IncludeDBs, ExcludeDBs, HighTempDBThreshold, CollectSystemSpids, HideSelf, ObtainBatchText, ObtainQueryPlanForStatement, ObtainQueryPlanForBatch, ObtainLocksForBlockRelevantThreshold, InputBufferThreshold, ParallelWaitsThreshold, QueryPlanThreshold, QueryPlanThresholdBlockRel, BlockingChainThreshold, BlockingChainDepth, TranDetailsThreshold, MediumDurationThreshold, HighDurationThreshold, BatchDurationThreshold, LongTransactionThreshold, Retention_IdleSPIDs_NoTran, Retention_IdleSPIDs_WithShortTran, Retention_IdleSPIDs_WithLongTran, Retention_IdleSPIDs_HighTempDB, Retention_ActiveLow, Retention_ActiveMedium, Retention_ActiveHigh, Retention_ActiveBatch, Retention_CaptureTimes, DebugSpeed, ThresholdFilterRefresh, SaveBadDims, Enable8666, ResolvePageLatches, ResolveLockWaits, 
	HistoryInsertDate,
	TriggerAction,
	LastModifiedUser)
	SELECT 
	RowID, AutoWhoEnabled, BeginTime, EndTime, IntervalLength, IncludeIdleWithTran, IncludeIdleWithoutTran, DurationFilter, IncludeDBs, ExcludeDBs, HighTempDBThreshold, CollectSystemSpids, HideSelf, ObtainBatchText, ObtainQueryPlanForStatement, ObtainQueryPlanForBatch, ObtainLocksForBlockRelevantThreshold, InputBufferThreshold, ParallelWaitsThreshold, QueryPlanThreshold, QueryPlanThresholdBlockRel, BlockingChainThreshold, BlockingChainDepth, TranDetailsThreshold, MediumDurationThreshold, HighDurationThreshold, BatchDurationThreshold, LongTransactionThreshold, Retention_IdleSPIDs_NoTran, Retention_IdleSPIDs_WithShortTran, Retention_IdleSPIDs_WithLongTran, Retention_IdleSPIDs_HighTempDB, Retention_ActiveLow, Retention_ActiveMedium, Retention_ActiveHigh, Retention_ActiveBatch, Retention_CaptureTimes, DebugSpeed, ThresholdFilterRefresh, SaveBadDims, Enable8666, ResolvePageLatches, ResolveLockWaits, 
	GETDATE(),
	'Delete',
	SUSER_NAME()
	FROM AutoWho.Options;

	DISABLE TRIGGER AutoWho.trgDEL_AutoWhoOptions ON AutoWho.Options;

	DELETE FROM AutoWho.Options;

	ENABLE TRIGGER AutoWho.trgDEL_AutoWhoOptions ON AutoWho.Options;

	INSERT INTO AutoWho.Options 
		DEFAULT VALUES;

	RETURN 0;
END



GO
