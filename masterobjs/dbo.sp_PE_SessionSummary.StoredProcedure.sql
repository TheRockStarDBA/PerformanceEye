USE [master]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[sp_PE_SessionSummary] 
/*   
	PROCEDURE:		sp_Pe_SessionSummary

	AUTHOR:			Aaron Morelli
					email@TBD.com
					@sqlcrossjoin
					sqlcrossjoin.wordpress.com
					https://github.com/amorelli005/PerformanceEye

	PURPOSE: Returns 1 row per AutoWho capture time, with many aggregation columns indicating
		what was occurring in the system at that time. The goal of this procedure is to help
		the user quickly find "problem times" for this SQL instance, since often end users
		complain in general/vague terms about the nature of the app problem and the time in which
		those problems occurred. Whereas sp_SessionViewer gives detailed info at a particular 
		point in time (the "SPID Capture Time"), and is very useful for determining the root cause
		for problems, it is a poor tool for scanning through a larger time window very quickly.
		Thus, sp_SessionSummary is a complementary tool to help the SQL Server professional 
		focus on the problematic time window quickly, without sifting through mountains of data.

		(Though some might argue that 70+ columns of output is a mountain of data!)

    CHANGE LOG:	2015-09-16  Aaron Morelli		Development Start
				2016-04-26	Aaron Morelli		Final run-through and commenting


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
exec sp_SessionSummary @start='2015-10-08',@end='2015-10-08 14:00',
	@savespace=N'N',@orderby=1, @orderdir=N'A', @help=N'N'

--debug: select * from PerformanceEye.AutoWho.CaptureSummary
*/
(
	@start			DATETIME=NULL,			--the start of the time window. If NULL, defaults to 4 hours ago.
	@end			DATETIME=NULL,			-- the end of the time window. If NULL, defaults to 1 second ago.
	@savespace		NCHAR(1)=N'N',			--shorter column header names
	@orderby		INT=1,					-- the column number to order by. Column #'s are part of the column name if @savespace=N'N'
	@orderdir		NCHAR(1)=N'A',			-- (A)scending or (D)escending
	@help			NVARCHAR(10)=N'N'		-- "params", "columns", or "all" (anything else <> "N" maps to "all")
)
AS
BEGIN
	SET NOCOUNT ON;
	SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED;
	SET ANSI_PADDING ON;

	DECLARE @scratch__int				INT,
			@numCaptures				INT,
			@numNeedPopulation			INT,
			@orderbyColumnName			NVARCHAR(256),
			@helpexec					NVARCHAR(4000),
			@err__msg					NVARCHAR(MAX),
			@DynSQL						NVARCHAR(MAX),
			@helpstr					NVARCHAR(MAX);

	--We always print out the exec syntax (whether help was requested or not) so that the user can switch over to the Messages
	-- tab and see what their options are.
	SET @helpexec = N'
exec sp_SessionSummary @start=''<start datetime>'', @end=''<end datetime>'', @savespace = N''N | Y'', 
	@orderby = <integer greater than 0>, @orderdir = N''A | D'', @Help = N''N | params | columns | all''

	';

	IF ISNULL(@Help,N'z') <> N'N'
	BEGIN
		GOTO helpbasic
	END

	IF @start IS NULL
	BEGIN
		SET @start = DATEADD(HOUR, -4, GETDATE());
		RAISERROR('Parameter @start set to 4 hours ago because a NULL value was supplied.', 10, 1) WITH NOWAIT;
	END

	IF @end IS NULL
	BEGIN
		SET @end = DATEADD(SECOND,-1, GETDATE());
		RAISERROR('Parameter @end set to 1 second ago because a NULL value was supplied.',10,1) WITH NOWAIT;
	END

	--Now that we have @start and @end values, replace our @helpexec string with them
	SET @helpexec = REPLACE(@helpexec,'<start datetime>', REPLACE(CONVERT(NVARCHAR(20), @start, 102),'.','-') + ' ' + CONVERT(NVARCHAR(20), @start, 108) + '.' + 
															RIGHT(CONVERT(NVARCHAR(20),N'000') + CONVERT(NVARCHAR(20),DATEPART(MILLISECOND, @start)),3)
							);
	SET @helpexec = REPLACE(@helpexec,'<end datetime>', REPLACE(CONVERT(NVARCHAR(20), @end, 102),'.','-') + ' ' + CONVERT(NVARCHAR(20), @end, 108) + '.' + 
															RIGHT(CONVERT(NVARCHAR(20),N'000') + CONVERT(NVARCHAR(20),DATEPART(MILLISECOND, @end)),3)
							);

	IF @start > GETDATE() OR @end > GETDATE()
	BEGIN
		RAISERROR(@helpexec,10,1);
		RAISERROR('Neither of the parameters @start or @end can be in the future.',16,1);
		RETURN -1;
	END
	
	IF @end <= @start
	BEGIN
		RAISERROR(@helpexec,10,1);
		RAISERROR('Parameter @end cannot be <= to parameter @start', 16, 1);
		RETURN -1;
	END

	IF UPPER(ISNULL(@savespace,N'z')) NOT IN (N'N', N'Y')
	BEGIN
		RAISERROR(@helpexec,10,1);
		RAISERROR('Parameter @savespace must be either "N" or "Y"',16,1);
		RETURN -1;
	END

	IF ISNULL(@orderby,-1) < 1 OR ISNULL(@orderby,999) > 82
	BEGIN
		RAISERROR(@helpexec,10,1);
		RAISERROR('Parameter @orderby must be an integer between 1 and 82.',16,1);
		RETURN -1;
	END


	IF UPPER(ISNULL(@orderdir,N'z')) NOT IN (N'A', N'D') 
	BEGIN
		RAISERROR(@helpexec,10,1);
		RAISERROR('Parameter @orderdir must be either A (ascending, default) or D (descending)',16,1);
		RETURN -1;
	END
	
	--If this is a summary run, check to see if there are any AutoWho.CaptureTime entries (in our @st/@et range) that haven't been 
	-- processed (into the AutoWho.CaptureSummary table) yet
	SELECT 
		@numCaptures = ss.numCaptures,
		@numNeedPopulation = ss.numNeedPopulation
	FROM (
		SELECT COUNT(*) as numCaptures,
			SUM(CASE WHEN t.CaptureSummaryPopulated = 0 THEN 1 ELSE 0 END) as numNeedPopulation
		FROM PerformanceEye.AutoWho.CaptureTimes t
		WHERE t.RunWasSuccessful = 1
		AND t.SPIDCaptureTime BETWEEN @start AND @end
	) ss
	;

	IF ISNULL(@numCaptures,0) = 0
	BEGIN
		RAISERROR(@helpexec,10,1);
		RAISERROR('
		***There is no capture data from AutoWho for the time range specified.',10,1) WITH NOWAIT;
		RETURN 0;
	END

	IF ISNULL(@numNeedPopulation,0) > 0 
	BEGIN
		SET @scratch__int = NULL;
		EXEC @scratch__int = PerformanceEye.AutoWho.PopulateCaptureSummary @StartTime = @start, @EndTime = @end;
			--returns 0 if successful
			--returns 1 if there were 0 capture times for this range, which shouldn't happen since we just found CaptureSummaryPopulated > 0 rows
			--returns -1 if an error occurred (the error is logged to dbo.AutoWhoLog

		IF @scratch__int IS NULL OR @scratch__int < 0
		BEGIN
			SET @err__msg = N'Unexpected error occurred while retrieving the AutoWho summary data. More info is available in the AutoWho log under the tag "SummCapturePopulation".'
			RAISERROR(@err__msg, 16, 1);
			RETURN -1;
		END
	END	--IF there are unprocessed capture times in the range


	--Convert our order by number to an order by column name
	SELECT @orderbyColumnName = CASE 
		WHEN @orderby = 1 THEN N'SPIDCaptureTime'
		WHEN @orderby = 2 THEN N'CapturedSPIDs'

		WHEN @orderby = 3 THEN N'Active'
		WHEN @orderby = 4 THEN N'ActLongest_ms'
		WHEN @orderby = 5 THEN N'ActAvg_ms'
		WHEN @orderby = 6 THEN N'Act0to1'
		WHEN @orderby = 7 THEN N'Act1to5'
		WHEN @orderby = 8 THEN N'Act5to10'
		WHEN @orderby = 9 THEN N'Act10to30'
		WHEN @orderby = 10 THEN N'Act30to60'
		WHEN @orderby = 11 THEN N'Act60to300'
		WHEN @orderby = 12 THEN N'Act300plus'

		WHEN @orderby = 13 THEN N'SPIDCaptureTime'

		WHEN @orderby = 14 THEN N'IdleWithOpenTran'
		WHEN @orderby = 15 THEN N'IdlOpTrnLongest_ms'
		WHEN @orderby = 16 THEN N'IdlOpTrnAvg_ms'
		WHEN @orderby = 17 THEN N'IdlOpTrn0to1'
		WHEN @orderby = 18 THEN N'IdlOpTrn1to5'
		WHEN @orderby = 19 THEN N'IdlOpTrn5to10'
		WHEN @orderby = 20 THEN N'IdlOpTrn10to30'
		WHEN @orderby = 21 THEN N'IdlOpTrn30to60'
		WHEN @orderby = 22 THEN N'IdlOpTrn60to300'
		WHEN @orderby = 23 THEN N'IdlOpTrn300plus'

		WHEN @orderby = 24 THEN N'SPIDCaptureTime'

		WHEN @orderby = 25 THEN N'WithOpenTran'
		WHEN @orderby = 26 THEN N'TranDurLongest_ms'
		WHEN @orderby = 27 THEN N'TranDurAvg_ms'
		WHEN @orderby = 28 THEN N'TranDur0to1'
		WHEN @orderby = 29 THEN N'TranDur1to5'
		WHEN @orderby = 30 THEN N'TranDur5to10'
		WHEN @orderby = 31 THEN N'TranDur10to30'
		WHEN @orderby = 32 THEN N'TranDur30to60'
		WHEN @orderby = 33 THEN N'TranDur60to300'
		WHEN @orderby = 34 THEN N'TranDur300plus'

		WHEN @orderby = 35 THEN N'SPIDCaptureTime'

		WHEN @orderby = 36 THEN N'Blocked'
		WHEN @orderby = 37 THEN N'BlockedLongest_ms'
		WHEN @orderby = 38 THEN N'BlockedAvg_ms'
		WHEN @orderby = 39 THEN N'Blocked0to1'
		WHEN @orderby = 40 THEN N'Blocked1to5'
		WHEN @orderby = 41 THEN N'Blocked5to10'
		WHEN @orderby = 42 THEN N'Blocked10to30'
		WHEN @orderby = 43 THEN N'Blocked30to60'
		WHEN @orderby = 44 THEN N'Blocked60to300'
		WHEN @orderby = 45 THEN N'Blocked300plus'

		WHEN @orderby = 46 THEN N'SPIDCaptureTime'
		
		WHEN @orderby = 47 THEN N'WaitingSPIDs'
		WHEN @orderby = 48 THEN N'WaitingTasks'
		WHEN @orderby = 49 THEN N'WaitingTaskLongest_ms'
		WHEN @orderby = 50 THEN N'WaitingTaskAvg_ms'
		WHEN @orderby = 51 THEN N'WaitingTask0to1'
		WHEN @orderby = 52 THEN N'WaitingTask1to5'
		WHEN @orderby = 53 THEN N'WaitingTask5to10'
		WHEN @orderby = 54 THEN N'WaitingTask10to30'
		WHEN @orderby = 55 THEN N'WaitingTask30to60'
		WHEN @orderby = 56 THEN N'WaitingTask60to300'
		WHEN @orderby = 57 THEN N'WaitingTask300plus'

		WHEN @orderby = 58 THEN N'SPIDCaptureTime'

		WHEN @orderby = 59 THEN N'TlogUsed_MB'
		WHEN @orderby = 60 THEN N'LargestLogWriter_MB'
		WHEN @orderby = 61 THEN N'QueryMemory_MB'
		WHEN @orderby = 62 THEN N'LargestMemoryGrant_MB'
		WHEN @orderby = 63 THEN N'TempDB_MB'
		WHEN @orderby = 64 THEN N'LargestTempDBConsumer_MB'
		WHEN @orderby = 65 THEN N'CPUused'
		WHEN @orderby = 66 THEN N'CPUDelta'
		WHEN @orderby = 67 THEN N'LargestCPUConsumer'
		WHEN @orderby = 68 THEN N'AllocatedTasks'
		
		WHEN @orderby = 69 THEN N'SPIDCaptureTime'

		WHEN @orderby = 70 THEN N'WritesDone'
		WHEN @orderby = 71 THEN N'WritesDelta'
		WHEN @orderby = 72 THEN N'LargestWriter'
		WHEN @orderby = 73 THEN N'LogicalReadsDone'
		WHEN @orderby = 74 THEN N'LogicalReadsDelta'
		WHEN @orderby = 75 THEN N'LargestLogicalReader'
		WHEN @orderby = 76 THEN N'PhysicalReadsDone'
		WHEN @orderby = 77 THEN N'PhysicalReadsDelta'
		WHEN @orderby = 78 THEN N'LargestPhysicalReader'
		WHEN @orderby = 79 THEN N'BlockingGraph'
		WHEN @orderby = 80 THEN N'LockDetails'
		WHEN @orderby = 81 THEN N'TranDetails'
		WHEN @orderby = 82 THEN N'SPIDCaptureTime'
		ELSE N'error' END;

	--	Group 1: Active spids
	SET @DynSQL = N'
	SELECT 
		SPIDCaptureTime' + CASE WHEN @savespace = N'Y' THEN N' as SCT' ELSE N' as [1_SPIDCaptureTime]' END + N' 
		,CapturedSPIDs' + CASE WHEN @savespace = N'Y' THEN N' as [#SPIDs]' ELSE N' as [2_TotCapturedSPIDs]' END + N' 
		,CASE WHEN Active=0 THEN N'''' ELSE CONVERT(nvarchar(20),Active) END' + CASE WHEN @savespace = N'Y' THEN N' as [Act]' ELSE N' as [3_Active]' END + N'
		,CASE WHEN ISNULL(ActLongest_ms,0) <= 0 THEN N'''' 
			ELSE (CASE WHEN ActLongest_ms > 863999999 THEN N''(!!) '' + CONVERT(nvarchar(20), (ActLongest_ms/1000) / 86400) + N''~'' +			--day
					REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),((ActLongest_ms/1000) % 86400)/3600)),1,2)) + N'':'' +			--hour
					REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((ActLongest_ms/1000) % 86400)%3600)/60)),1,2)) + N'':'' +			--minute
					REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((ActLongest_ms/1000) % 86400)%3600)%60)),1,2)) 					--second

			WHEN ActLongest_ms > 86399999 THEN N''(!) '' + REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(ActLongest_ms/1000) / 86400)),1,2)) + N''~'' +			--day
					REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),((ActLongest_ms/1000) % 86400)/3600)),1,2)) + N'':'' +			--hour
					REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((ActLongest_ms/1000) % 86400)%3600)/60)),1,2)) + N'':'' +			--minute
					REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((ActLongest_ms/1000) % 86400)%3600)%60)),1,2)) 			--second

			WHEN ActLongest_ms > 59999 THEN REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),((ActLongest_ms/1000) % 86400)/3600)),1,2)) + N'':'' +			--hour
					REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((ActLongest_ms/1000) % 86400)%3600)/60)),1,2)) + N'':'' +			--minute
					REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((ActLongest_ms/1000) % 86400)%3600)%60)),1,2)) 			--second
			ELSE SUBSTRING(ActLongest_ms_money, 1, CHARINDEX(''.'',ActLongest_ms_money)-1)
			END)
		END' + CASE WHEN @savespace = N'Y' THEN N' as [LngAct]' ELSE N' as [4_LongestActive]' END;

	SET @DynSQL = @DynSQL + N' 
	,CASE WHEN ISNULL(ActAvg_ms,0) <= 0 THEN N'''' 
			ELSE (CASE WHEN ActAvg_ms > 863999999 THEN N''(!!) '' + CONVERT(nvarchar(20), (ActAvg_ms/1000) / 86400) + N''~'' +			--day
					REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),((ActAvg_ms/1000) % 86400)/3600)),1,2)) + N'':'' +			--hour
					REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((ActAvg_ms/1000) % 86400)%3600)/60)),1,2)) + N'':'' +			--minute
					REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((ActAvg_ms/1000) % 86400)%3600)%60)),1,2)) 					--second

			WHEN ActAvg_ms > 86399999 THEN N''(!) '' + REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(ActAvg_ms/1000) / 86400)),1,2)) + N''~'' +			--day
					REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),((ActAvg_ms/1000) % 86400)/3600)),1,2)) + N'':'' +			--hour
					REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((ActAvg_ms/1000) % 86400)%3600)/60)),1,2)) + N'':'' +			--minute
					REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((ActAvg_ms/1000) % 86400)%3600)%60)),1,2)) 			--second

			WHEN ActAvg_ms > 59999 THEN REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),((ActAvg_ms/1000) % 86400)/3600)),1,2)) + N'':'' +			--hour
					REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((ActAvg_ms/1000) % 86400)%3600)/60)),1,2)) + N'':'' +			--minute
					REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((ActAvg_ms/1000) % 86400)%3600)%60)),1,2)) 			--second
			ELSE SUBSTRING(ActAvg_ms_money, 1, CHARINDEX(''.'',ActAvg_ms_money)-1)
			END)
		END' + CASE WHEN @savespace = N'Y' THEN N' as [AvgAct]' ELSE N' as [5_AvgActive]' END;

	SET @DynSQL = @DynSQL + N'
		,CASE WHEN ISNULL(Act0to1,0)=0 THEN N'''' ELSE CONVERT(nvarchar(20),Act0to1) END' + CASE WHEN @savespace=N'Y' THEN N' as [0to1]' ELSE N' as [6_Act0to1]' END + N'
		,CASE WHEN ISNULL(Act1to5,0)=0 THEN N'''' ELSE CONVERT(nvarchar(20),Act1to5) END' + CASE WHEN @savespace=N'Y' THEN N' as [1to5]' ELSE N' as [7_Act1to5]' END + N'
		,CASE WHEN ISNULL(Act5to10,0)=0 THEN N'''' ELSE CONVERT(nvarchar(20),Act5to10) END' + CASE WHEN @savespace=N'Y' THEN N' as [5to10]' ELSE N' as [8_Act5to10]' END + N'
		,CASE WHEN ISNULL(Act10to30,0)=0 THEN N'''' ELSE CONVERT(nvarchar(20),Act10to30) END' + CASE WHEN @savespace=N'Y' THEN N' as [10to30]' ELSE N' as [9_Act10to30]' END + N'
		,CASE WHEN ISNULL(Act30to60,0)=0 THEN N'''' ELSE CONVERT(nvarchar(20),Act30to60) END' + CASE WHEN @savespace=N'Y' THEN N' as [30to60]' ELSE N' as [10_Act30to60]' END + N'
		,CASE WHEN ISNULL(Act60to300,0)=0 THEN N'''' ELSE CONVERT(nvarchar(20),Act60to300) END' + CASE WHEN @savespace=N'Y' THEN N' as [60to300]' ELSE N' as [11_Act60to300]' END + N'
		,CASE WHEN ISNULL(Act300plus,0)=0 THEN N'''' ELSE CONVERT(nvarchar(20),Act300plus) END' + CASE WHEN @savespace=N'Y' THEN N' as [300plus]' ELSE N' as [12_Act300plus]' END + N'
		,SPIDCaptureTime' + CASE WHEN @savespace = N'Y' THEN N' as SCT' ELSE N' as [13_SPIDCaptureTime]' END + N'
	';


	--	Group 2: Idle spids with open trans
	SET @DynSQL = @DynSQL + N'
		,CASE WHEN ISNULL(IdleWithOpenTran,0) = 0 THEN N'''' ELSE CONVERT(nvarchar(20),IdleWithOpenTran) END' + CASE WHEN @savespace = N'Y' THEN N' as [Idle]' ELSE N' as [14_IdleWOpenTran]' END + N'
		,CASE WHEN ISNULL(IdlOpTrnLongest_ms,0) = 0 THEN N'''' 
			ELSE (CASE WHEN IdlOpTrnLongest_ms > 863999999 THEN N''(!!) '' + CONVERT(nvarchar(20), (IdlOpTrnLongest_ms/1000) / 86400) + N''~'' +			--day
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),((IdlOpTrnLongest_ms/1000) % 86400)/3600)),1,2)) + N'':'' +			--hour
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((IdlOpTrnLongest_ms/1000) % 86400)%3600)/60)),1,2)) + N'':'' +			--minute
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((IdlOpTrnLongest_ms/1000) % 86400)%3600)%60)),1,2)) 					--second

			WHEN IdlOpTrnLongest_ms > 86399999 THEN N''(!) '' + REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(IdlOpTrnLongest_ms/1000) / 86400)),1,2)) + N''~'' +			--day
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),((IdlOpTrnLongest_ms/1000) % 86400)/3600)),1,2)) + N'':'' +			--hour
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((IdlOpTrnLongest_ms/1000) % 86400)%3600)/60)),1,2)) + N'':'' +			--minute
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((IdlOpTrnLongest_ms/1000) % 86400)%3600)%60)),1,2)) 			--second

			WHEN IdlOpTrnLongest_ms > 59999 THEN REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),((IdlOpTrnLongest_ms/1000) % 86400)/3600)),1,2)) + N'':'' +			--hour
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((IdlOpTrnLongest_ms/1000) % 86400)%3600)/60)),1,2)) + N'':'' +			--minute
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((IdlOpTrnLongest_ms/1000) % 86400)%3600)%60)),1,2)) 			--second
			ELSE SUBSTRING(IdlOpTrnLongest_ms_money, 1, CHARINDEX(''.'',IdlOpTrnLongest_ms_money)-1)
			END)
		END' + CASE WHEN @savespace = N'Y' THEN N' as [LngIdl]' ELSE N' as [15_LongestIdleWTran]' END;

	SET @DynSQL = @DynSQL + N'
	,CASE WHEN ISNULL(IdlOpTrnAvg_ms,0) = 0 THEN N'''' 
			ELSE (CASE WHEN IdlOpTrnAvg_ms > 863999999 THEN N''(!!) '' + CONVERT(nvarchar(20), (IdlOpTrnAvg_ms/1000) / 86400) + N''~'' +			--day
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),((IdlOpTrnAvg_ms/1000) % 86400)/3600)),1,2)) + N'':'' +			--hour
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((IdlOpTrnAvg_ms/1000) % 86400)%3600)/60)),1,2)) + N'':'' +			--minute
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((IdlOpTrnAvg_ms/1000) % 86400)%3600)%60)),1,2)) 					--second

			WHEN IdlOpTrnAvg_ms > 86399999 THEN N''(!) '' + REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(IdlOpTrnAvg_ms/1000) / 86400)),1,2)) + N''~'' +			--day
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),((IdlOpTrnAvg_ms/1000) % 86400)/3600)),1,2)) + N'':'' +			--hour
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((IdlOpTrnAvg_ms/1000) % 86400)%3600)/60)),1,2)) + N'':'' +			--minute
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((IdlOpTrnAvg_ms/1000) % 86400)%3600)%60)),1,2)) 			--second

			WHEN IdlOpTrnAvg_ms > 59999 THEN REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),((IdlOpTrnAvg_ms/1000) % 86400)/3600)),1,2)) + N'':'' +			--hour
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((IdlOpTrnAvg_ms/1000) % 86400)%3600)/60)),1,2)) + N'':'' +			--minute
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((IdlOpTrnAvg_ms/1000) % 86400)%3600)%60)),1,2)) 			--second
			ELSE SUBSTRING(IdlOpTrnAvg_ms_money, 1, CHARINDEX(''.'',IdlOpTrnAvg_ms_money)-1)
			END)
		END' + CASE WHEN @savespace = N'Y' THEN N' as [AvgIdl]' ELSE N' as [16_AvgIdleWTran]' END;

	SET @DynSQL = @DynSQL + N'
		,CASE WHEN ISNULL(IdlOpTrn0to1,0)=0 THEN N'''' ELSE CONVERT(nvarchar(20),IdlOpTrn0to1) END' + CASE WHEN @savespace=N'Y' THEN N' as [0to1]' ELSE N' as [17_IdlOpTrn0to1]' END + N'
		,CASE WHEN ISNULL(IdlOpTrn1to5,0)=0 THEN N'''' ELSE CONVERT(nvarchar(20),IdlOpTrn1to5) END' + CASE WHEN @savespace=N'Y' THEN N' as [1to5]' ELSE N' as [18_IdlOpTrn1to5]' END + N'
		,CASE WHEN ISNULL(IdlOpTrn5to10,0)=0 THEN N'''' ELSE CONVERT(nvarchar(20),IdlOpTrn5to10) END' + CASE WHEN @savespace=N'Y' THEN N' as [5to10]' ELSE N' as [19_IdlOpTrn5to10]' END + N'
		,CASE WHEN ISNULL(IdlOpTrn10to30,0)=0 THEN N'''' ELSE CONVERT(nvarchar(20),IdlOpTrn10to30) END' + CASE WHEN @savespace=N'Y' THEN N' as [10to30]' ELSE N' as [20_IdlOpTrn10to30]' END + N'
		,CASE WHEN ISNULL(IdlOpTrn30to60,0)=0 THEN N'''' ELSE CONVERT(nvarchar(20),IdlOpTrn30to60) END' + CASE WHEN @savespace=N'Y' THEN N' as [30to60]' ELSE N' as [21_IdlOpTrn30to60]' END + N'
		,CASE WHEN ISNULL(IdlOpTrn60to300,0)=0 THEN N'''' ELSE CONVERT(nvarchar(20),IdlOpTrn60to300) END' + CASE WHEN @savespace=N'Y' THEN N' as [60to300]' ELSE N' as [22_IdlOpTrn60to300]' END + N'
		,CASE WHEN ISNULL(IdlOpTrn300plus,0)=0 THEN N'''' ELSE CONVERT(nvarchar(20),IdlOpTrn300plus) END' + CASE WHEN @savespace=N'Y' THEN N' as [300plus]' ELSE N' as [23_IdlOpTrn300plus]' END + N'
		,SPIDCaptureTime' + CASE WHEN @savespace = N'Y' THEN N' as SCT' ELSE N' as [24_SPIDCaptureTime]' END + N'
	';

	--	Group 3: Open Trans (both active and idle spids)
	SET @DynSQL = @DynSQL + N'
		,CASE WHEN ISNULL(WithOpenTran,0) = 0 THEN N'''' ELSE CONVERT(nvarchar(20),WithOpenTran) END' + CASE WHEN @savespace = N'Y' THEN N' as [OpenTran]' ELSE N' as [25_wOpenTran]' END + N'
		,CASE WHEN ISNULL(TranDurLongest_ms,0) = 0 THEN N'''' 
			ELSE (CASE WHEN TranDurLongest_ms > 863999999 THEN N''(!!) '' + CONVERT(nvarchar(20), (TranDurLongest_ms/1000) / 86400) + N''~'' +			--day
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),((TranDurLongest_ms/1000) % 86400)/3600)),1,2)) + N'':'' +			--hour
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((TranDurLongest_ms/1000) % 86400)%3600)/60)),1,2)) + N'':'' +			--minute
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((TranDurLongest_ms/1000) % 86400)%3600)%60)),1,2)) 					--second

			WHEN TranDurLongest_ms > 86399999 THEN N''(!) '' + REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(TranDurLongest_ms/1000) / 86400)),1,2)) + N''~'' +			--day
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),((TranDurLongest_ms/1000) % 86400)/3600)),1,2)) + N'':'' +			--hour
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((TranDurLongest_ms/1000) % 86400)%3600)/60)),1,2)) + N'':'' +			--minute
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((TranDurLongest_ms/1000) % 86400)%3600)%60)),1,2)) 			--second

			WHEN TranDurLongest_ms > 59999 THEN REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),((TranDurLongest_ms/1000) % 86400)/3600)),1,2)) + N'':'' +			--hour
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((TranDurLongest_ms/1000) % 86400)%3600)/60)),1,2)) + N'':'' +			--minute
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((TranDurLongest_ms/1000) % 86400)%3600)%60)),1,2)) 			--second
			ELSE SUBSTRING(TranDurLongest_ms_money, 1, CHARINDEX(''.'',TranDurLongest_ms_money)-1)
			END)
		END' + CASE WHEN @savespace = N'Y' THEN N' as [LngTrn]' ELSE N' as [26_LongestTran]' END;

	SET @DynSQL = @DynSQL + N'
	,CASE WHEN ISNULL(TranDurAvg_ms,0) = 0 THEN N'''' 
			ELSE (CASE WHEN TranDurAvg_ms > 863999999 THEN N''(!!) '' + CONVERT(nvarchar(20), (TranDurAvg_ms/1000) / 86400) + N''~'' +			--day
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),((TranDurAvg_ms/1000) % 86400)/3600)),1,2)) + N'':'' +			--hour
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((TranDurAvg_ms/1000) % 86400)%3600)/60)),1,2)) + N'':'' +			--minute
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((TranDurAvg_ms/1000) % 86400)%3600)%60)),1,2)) 					--second

			WHEN TranDurAvg_ms > 86399999 THEN N''(!) '' + REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(TranDurAvg_ms/1000) / 86400)),1,2)) + N''~'' +			--day
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),((TranDurAvg_ms/1000) % 86400)/3600)),1,2)) + N'':'' +			--hour
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((TranDurAvg_ms/1000) % 86400)%3600)/60)),1,2)) + N'':'' +			--minute
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((TranDurAvg_ms/1000) % 86400)%3600)%60)),1,2)) 			--second

			WHEN TranDurAvg_ms > 59999 THEN REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),((TranDurAvg_ms/1000) % 86400)/3600)),1,2)) + N'':'' +			--hour
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((TranDurAvg_ms/1000) % 86400)%3600)/60)),1,2)) + N'':'' +			--minute
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((TranDurAvg_ms/1000) % 86400)%3600)%60)),1,2)) 			--second
			ELSE SUBSTRING(TranDurAvg_ms_money, 1, CHARINDEX(''.'',TranDurAvg_ms_money)-1)
			END)
		END' + CASE WHEN @savespace = N'Y' THEN N' as [AvgTrn]' ELSE N' as [27_AvgTran]' END;

	SET @DynSQL = @DynSQL + N'
		,CASE WHEN ISNULL(TranDur0to1,0)=0 THEN N'''' ELSE CONVERT(nvarchar(20),TranDur0to1) END' + CASE WHEN @savespace=N'Y' THEN N' as [0to1]' ELSE N' as [28_TranDur0to1]' END + N'
		,CASE WHEN ISNULL(TranDur1to5,0)=0 THEN N'''' ELSE CONVERT(nvarchar(20),TranDur1to5) END' + CASE WHEN @savespace=N'Y' THEN N' as [1to5]' ELSE N' as [29_TranDur1to5]' END + N'
		,CASE WHEN ISNULL(TranDur5to10,0)=0 THEN N'''' ELSE CONVERT(nvarchar(20),TranDur5to10) END' + CASE WHEN @savespace=N'Y' THEN N' as [5to10]' ELSE N' as [30_TranDur5to10]' END + N'
		,CASE WHEN ISNULL(TranDur10to30,0)=0 THEN N'''' ELSE CONVERT(nvarchar(20),TranDur10to30) END' + CASE WHEN @savespace=N'Y' THEN N' as [10to30]' ELSE N' as [31_TranDur10to30]' END + N'
		,CASE WHEN ISNULL(TranDur30to60,0)=0 THEN N'''' ELSE CONVERT(nvarchar(20),TranDur30to60) END' + CASE WHEN @savespace=N'Y' THEN N' as [30to60]' ELSE N' as [32_TranDur30to60]' END + N'
		,CASE WHEN ISNULL(TranDur60to300,0)=0 THEN N'''' ELSE CONVERT(nvarchar(20),TranDur60to300) END' + CASE WHEN @savespace=N'Y' THEN N' as [60to300]' ELSE N' as [33_TranDur60to300]' END + N'
		,CASE WHEN ISNULL(TranDur300plus,0)=0 THEN N'''' ELSE CONVERT(nvarchar(20),TranDur300plus) END' + CASE WHEN @savespace=N'Y' THEN N' as [300plus]' ELSE N' as [34_TranDur300plus]' END + N'
		,SPIDCaptureTime' + CASE WHEN @savespace = N'Y' THEN N' as SCT' ELSE N' as [35_SPIDCaptureTime]' END + N'
	';

	--Group 4: Blocked spids
	SET @DynSQL = @DynSQL + N'
		,CASE WHEN ISNULL(Blocked,0) = 0 THEN N'''' ELSE CONVERT(nvarchar(20),Blocked) END' + CASE WHEN @savespace = N'Y' THEN N' as [Blkd]' ELSE N' as [36_Blocked]' END + N'
		,CASE WHEN ISNULL(BlockedLongest_ms,0) = 0 THEN N'''' 
			ELSE (CASE WHEN BlockedLongest_ms > 863999999 THEN N''(!!) '' + CONVERT(nvarchar(20), (BlockedLongest_ms/1000) / 86400) + N''~'' +			--day
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),((BlockedLongest_ms/1000) % 86400)/3600)),1,2)) + N'':'' +			--hour
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((BlockedLongest_ms/1000) % 86400)%3600)/60)),1,2)) + N'':'' +			--minute
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((BlockedLongest_ms/1000) % 86400)%3600)%60)),1,2)) 					--second

			WHEN BlockedLongest_ms > 86399999 THEN N''(!) '' + REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(BlockedLongest_ms/1000) / 86400)),1,2)) + N''~'' +			--day
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),((BlockedLongest_ms/1000) % 86400)/3600)),1,2)) + N'':'' +			--hour
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((BlockedLongest_ms/1000) % 86400)%3600)/60)),1,2)) + N'':'' +			--minute
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((BlockedLongest_ms/1000) % 86400)%3600)%60)),1,2)) 			--second

			WHEN BlockedLongest_ms > 59999 THEN REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),((BlockedLongest_ms/1000) % 86400)/3600)),1,2)) + N'':'' +			--hour
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((BlockedLongest_ms/1000) % 86400)%3600)/60)),1,2)) + N'':'' +			--minute
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((BlockedLongest_ms/1000) % 86400)%3600)%60)),1,2)) 			--second
			ELSE SUBSTRING(BlockedLongest_ms_money, 1, CHARINDEX(''.'',BlockedLongest_ms_money)-1)
			END)
		END' + CASE WHEN @savespace = N'Y' THEN N' as [LngBlk]' ELSE N' as [37_LongBlkTask]' END;

	SET @DynSQL = @DynSQL + N'
	,CASE WHEN ISNULL(BlockedAvg_ms,0) = 0 THEN N'''' 
			ELSE (CASE WHEN BlockedAvg_ms > 863999999 THEN N''(!!) '' + CONVERT(nvarchar(20), (BlockedAvg_ms/1000) / 86400) + N''~'' +			--day
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),((BlockedAvg_ms/1000) % 86400)/3600)),1,2)) + N'':'' +			--hour
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((BlockedAvg_ms/1000) % 86400)%3600)/60)),1,2)) + N'':'' +			--minute
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((BlockedAvg_ms/1000) % 86400)%3600)%60)),1,2)) 					--second

			WHEN BlockedAvg_ms > 86399999 THEN N''(!) '' + REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(BlockedAvg_ms/1000) / 86400)),1,2)) + N''~'' +			--day
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),((BlockedAvg_ms/1000) % 86400)/3600)),1,2)) + N'':'' +			--hour
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((BlockedAvg_ms/1000) % 86400)%3600)/60)),1,2)) + N'':'' +			--minute
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((BlockedAvg_ms/1000) % 86400)%3600)%60)),1,2)) 			--second

			WHEN BlockedAvg_ms > 59999 THEN REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),((BlockedAvg_ms/1000) % 86400)/3600)),1,2)) + N'':'' +			--hour
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((BlockedAvg_ms/1000) % 86400)%3600)/60)),1,2)) + N'':'' +			--minute
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((BlockedAvg_ms/1000) % 86400)%3600)%60)),1,2)) 			--second
			ELSE SUBSTRING(BlockedAvg_ms_money, 1, CHARINDEX(''.'',BlockedAvg_ms_money)-1)
			END)
		END' + CASE WHEN @savespace = N'Y' THEN N' as [AvgBlk]' ELSE N' as [38_AvgBlkTask]' END;

	SET @DynSQL = @DynSQL + N'
		,CASE WHEN ISNULL(Blocked0to1,0)=0 THEN N'''' ELSE CONVERT(nvarchar(20),Blocked0to1) END' + CASE WHEN @savespace=N'Y' THEN N' as [0to1]' ELSE N' as [39_Blocked0to1]' END + N'
		,CASE WHEN ISNULL(Blocked1to5,0)=0 THEN N'''' ELSE CONVERT(nvarchar(20),Blocked1to5) END' + CASE WHEN @savespace=N'Y' THEN N' as [1to5]' ELSE N' as [40_Blocked1to5]' END + N'
		,CASE WHEN ISNULL(Blocked5to10,0)=0 THEN N'''' ELSE CONVERT(nvarchar(20),Blocked5to10) END' + CASE WHEN @savespace=N'Y' THEN N' as [5to10]' ELSE N' as [41_Blocked5to10]' END + N'
		,CASE WHEN ISNULL(Blocked10to30,0)=0 THEN N'''' ELSE CONVERT(nvarchar(20),Blocked10to30) END' + CASE WHEN @savespace=N'Y' THEN N' as [10to30]' ELSE N' as [42_Blocked10to30]' END + N'
		,CASE WHEN ISNULL(Blocked30to60,0)=0 THEN N'''' ELSE CONVERT(nvarchar(20),Blocked30to60) END' + CASE WHEN @savespace=N'Y' THEN N' as [30to60]' ELSE N' as [43_Blocked30to60]' END + N'
		,CASE WHEN ISNULL(Blocked60to300,0)=0 THEN N'''' ELSE CONVERT(nvarchar(20),Blocked60to300) END' + CASE WHEN @savespace=N'Y' THEN N' as [60to300]' ELSE N' as [44_Blocked60to300]' END + N'
		,CASE WHEN ISNULL(Blocked300plus,0)=0 THEN N'''' ELSE CONVERT(nvarchar(20),Blocked300plus) END' + CASE WHEN @savespace=N'Y' THEN N' as [300plus]' ELSE N' as [45_Blocked300plus]' END + N'
		,SPIDCaptureTime' + CASE WHEN @savespace = N'Y' THEN N' as SCT' ELSE N' as [46_SPIDCaptureTime]' END + N'
	';

	--Group 5: waiting (Unlike SQL Server standard terminology, "waiting" here means not blocked by another spid, but not able to progress)
	SET @DynSQL = @DynSQL + N'
		,CASE WHEN ISNULL(WaitingSPIDs,0) = 0 THEN N'''' ELSE CONVERT(nvarchar(20),WaitingSPIDs) END' + CASE WHEN @savespace = N'Y' THEN N' as [WtSPIDs]' ELSE N' as [47_WaitingSPIDs]' END + N'
		,CASE WHEN ISNULL(WaitingTasks,0) = 0 THEN N'''' ELSE CONVERT(nvarchar(20), WaitingTasks) END' + CASE WHEN @savespace = N'Y' THEN N' as [WtTsk]' ELSE N' as [48_WaitingTasks]' END + N'
		,CASE WHEN ISNULL(WaitingTaskLongest_ms,0) = 0 THEN N'''' 
			ELSE (CASE WHEN WaitingTaskLongest_ms > 863999999 THEN N''(!!) '' + CONVERT(nvarchar(20), (WaitingTaskLongest_ms/1000) / 86400) + N''~'' +			--day
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),((WaitingTaskLongest_ms/1000) % 86400)/3600)),1,2)) + N'':'' +			--hour
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((WaitingTaskLongest_ms/1000) % 86400)%3600)/60)),1,2)) + N'':'' +			--minute
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((WaitingTaskLongest_ms/1000) % 86400)%3600)%60)),1,2)) 					--second

			WHEN WaitingTaskLongest_ms > 86399999 THEN N''(!) '' + REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(WaitingTaskLongest_ms/1000) / 86400)),1,2)) + N''~'' +			--day
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),((WaitingTaskLongest_ms/1000) % 86400)/3600)),1,2)) + N'':'' +			--hour
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((WaitingTaskLongest_ms/1000) % 86400)%3600)/60)),1,2)) + N'':'' +			--minute
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((WaitingTaskLongest_ms/1000) % 86400)%3600)%60)),1,2)) 			--second

			WHEN WaitingTaskLongest_ms > 59999 THEN REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),((WaitingTaskLongest_ms/1000) % 86400)/3600)),1,2)) + N'':'' +			--hour
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((WaitingTaskLongest_ms/1000) % 86400)%3600)/60)),1,2)) + N'':'' +			--minute
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((WaitingTaskLongest_ms/1000) % 86400)%3600)%60)),1,2)) 			--second
			ELSE SUBSTRING(WaitingTaskLongest_ms_money, 1, CHARINDEX(''.'',WaitingTaskLongest_ms_money)-1)
			END)
		END' + CASE WHEN @savespace = N'Y' THEN N' as [LngWtTsk]' ELSE N' as [49_LongestWaitTask]' END;

	SET @DynSQL = @DynSQL + N'
	,CASE WHEN ISNULL(WaitingTaskAvg_ms,0) = 0 THEN N'''' 
			ELSE (CASE WHEN WaitingTaskAvg_ms > 863999999 THEN N''(!!) '' + CONVERT(nvarchar(20), (WaitingTaskAvg_ms/1000) / 86400) + N''~'' +			--day
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),((WaitingTaskAvg_ms/1000) % 86400)/3600)),1,2)) + N'':'' +			--hour
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((WaitingTaskAvg_ms/1000) % 86400)%3600)/60)),1,2)) + N'':'' +			--minute
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((WaitingTaskAvg_ms/1000) % 86400)%3600)%60)),1,2)) 					--second

			WHEN WaitingTaskAvg_ms > 86399999 THEN N''(!) '' + REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(WaitingTaskAvg_ms/1000) / 86400)),1,2)) + N''~'' +			--day
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),((WaitingTaskAvg_ms/1000) % 86400)/3600)),1,2)) + N'':'' +			--hour
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((WaitingTaskAvg_ms/1000) % 86400)%3600)/60)),1,2)) + N'':'' +			--minute
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((WaitingTaskAvg_ms/1000) % 86400)%3600)%60)),1,2)) 			--second

			WHEN WaitingTaskAvg_ms > 59999 THEN REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),((WaitingTaskAvg_ms/1000) % 86400)/3600)),1,2)) + N'':'' +			--hour
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((WaitingTaskAvg_ms/1000) % 86400)%3600)/60)),1,2)) + N'':'' +			--minute
				REVERSE(SUBSTRING(REVERSE(N''0'' + CONVERT(nvarchar(20),(((WaitingTaskAvg_ms/1000) % 86400)%3600)%60)),1,2)) 			--second
			ELSE SUBSTRING(WaitingTaskAvg_ms_money, 1, CHARINDEX(''.'',WaitingTaskAvg_ms_money)-1)
			END)
		END' + CASE WHEN @savespace = N'Y' THEN N' as [AvgWtTsk]' ELSE N' as [50_AvgWaitTask]' END;

	SET @DynSQL = @DynSQL + N'
		,CASE WHEN ISNULL(WaitingTask0to1,0)=0 THEN N'''' ELSE CONVERT(nvarchar(20),WaitingTask0to1) END' + CASE WHEN @savespace=N'Y' THEN N' as [0to1]' ELSE N' as [51_WaitingTask0to1]' END + N'
		,CASE WHEN ISNULL(WaitingTask1to5,0)=0 THEN N'''' ELSE CONVERT(nvarchar(20),WaitingTask1to5) END' + CASE WHEN @savespace=N'Y' THEN N' as [1to5]' ELSE N' as [52_WaitingTask1to5]' END + N'
		,CASE WHEN ISNULL(WaitingTask5to10,0)=0 THEN N'''' ELSE CONVERT(nvarchar(20),WaitingTask5to10) END' + CASE WHEN @savespace=N'Y' THEN N' as [5to10]' ELSE N' as [53_WaitingTask5to10]' END + N'
		,CASE WHEN ISNULL(WaitingTask10to30,0)=0 THEN N'''' ELSE CONVERT(nvarchar(20),WaitingTask10to30) END' + CASE WHEN @savespace=N'Y' THEN N' as [10to30]' ELSE N' as [54_WaitingTask10to30]' END + N'
		,CASE WHEN ISNULL(WaitingTask30to60,0)=0 THEN N'''' ELSE CONVERT(nvarchar(20),WaitingTask30to60) END' + CASE WHEN @savespace=N'Y' THEN N' as [30to60]' ELSE N' as [55_WaitingTask30to60]' END + N'
		,CASE WHEN ISNULL(WaitingTask60to300,0)=0 THEN N'''' ELSE CONVERT(nvarchar(20),WaitingTask60to300) END' + CASE WHEN @savespace=N'Y' THEN N' as [60to300]' ELSE N' as [56_WaitingTask60to300]' END + N'
		,CASE WHEN ISNULL(WaitingTask300plus,0)=0 THEN N'''' ELSE CONVERT(nvarchar(20),WaitingTask300plus) END' + CASE WHEN @savespace=N'Y' THEN N' as [300plus]' ELSE N' as [57_WaitingTask300plus]' END + N'
		,SPIDCaptureTime' + CASE WHEN @savespace = N'Y' THEN N' as SCT' ELSE N' as [58_SPIDCaptureTime]' END + N'
	';

	--Group 6: Resources #1
	SET @DynSQL = @DynSQL + N'
		,CASE WHEN ISNULL(TlogUsed_MB,0) = 0 THEN N'''' ELSE CONVERT(nvarchar(20), TlogUsed_MB) END' + CASE WHEN @savespace = N'Y' THEN N' as [Tlog]' ELSE N' as [59_TLogUsed (MB)]' END + N'
		,CASE WHEN ISNULL(LargestLogWriter_MB,0) = 0 THEN N'''' ELSE CONVERT(nvarchar(20), LargestLogWriter_MB) END' + CASE WHEN @savespace = N'Y' THEN N' as [LargestTlog]' ELSE N' as [60_LargestLogWriter (MB)]' END + N' 
		,CASE WHEN ISNULL(QueryMemory_MB,0.0) = 0.0 THEN N'''' ELSE CONVERT(nvarchar(20),QueryMemory_MB) END' + CASE WHEN @savespace = N'Y' THEN N' as [Qmem]' ELSE N' as [61_QueryMem (MB)]' END + N'
		,CASE WHEN ISNULL(LargestMemoryGrant_MB,0.0) = 0.0 THEN N'''' ELSE CONVERT(nvarchar(20),CONVERT(money,LargestMemoryGrant_MB),1) END' + CASE WHEN @savespace = N'Y' THEN N' as [LargestQM]' ELSE N' as [62_LargestQueryMem (MB)]' END + N'
		,CASE WHEN ISNULL(TempDB_MB,0.0) = 0.0 THEN N'''' ELSE CONVERT(nvarchar(20),CONVERT(money,TempDB_MB),1) END' + CASE WHEN @savespace = N'Y' THEN N' as [Tdb]' ELSE N' as [63_TempDB (MB)]' END + N'
		,CASE WHEN ISNULL(LargestTempDBConsumer_MB,0.0) = 0.0 THEN N'''' ELSE CONVERT(nvarchar(20),CONVERT(money,LargestTempDBConsumer_MB),1) END' + CASE WHEN @savespace = N'Y' THEN N' as [LargestTdb]' ELSE N' as [64_LargestTempDB (MB)]' END + N'
		,CASE WHEN ISNULL(CPUused,0) = 0 THEN N'''' ELSE SUBSTRING(CPUused_money, 1, CHARINDEX(''.'',CPUused_money)-1) END' + CASE WHEN @savespace = N'Y' THEN N' as [CPU]' ELSE N' as [65_CPUused]' END + N'
		,CASE WHEN ISNULL(CPUDelta,0) = 0 THEN N'''' ELSE SUBSTRING(CPUDelta_money, 1, CHARINDEX(''.'',CPUDelta_money)-1) END' + CASE WHEN @savespace = N'Y' THEN N' as [CDelta]' ELSE N' as [66_CPUDelta]' END + N'
		,CASE WHEN ISNULL(LargestCPUConsumer,0) = 0 THEN N'''' ELSE SUBSTRING(LargestCPUConsumer_money, 1, CHARINDEX(''.'',LargestCPUConsumer_money)-1) END' + CASE WHEN @savespace = N'Y' THEN N' as [LrgCPU]' ELSE N' as [67_LargestCPUConsumer]' END + N'
		,CASE WHEN ISNULL(AllocatedTasks,0) = 0 THEN N'''' ELSE CONVERT(nvarchar(20),AllocatedTasks) END' + CASE WHEN @savespace = N'Y' THEN N' as [Tasks]' ELSE N' as [68_AllocatedTasks]' END


	--Group 7: Resources #2
	SET @DynSQL = @DynSQL + N'
		,SPIDCaptureTime' + CASE WHEN @savespace = N'Y' THEN N' as SCT' ELSE N' as [69_SPIDCaptureTime]' END + N'
		,CASE WHEN ISNULL(WritesDone,0) = 0 THEN N'''' ELSE SUBSTRING(WritesDone_money, 1, CHARINDEX(''.'',WritesDone_money)-1) END' + CASE WHEN @savespace = N'Y' THEN N' as [Wri]' ELSE N' as [70_WritesDone]' END + N'
		,CASE WHEN ISNULL(WritesDelta,0) = 0 THEN N'''' ELSE SUBSTRING(WritesDelta_money, 1, CHARINDEX(''.'',WritesDelta_money)-1) END' + CASE WHEN @savespace = N'Y' THEN N' as [WDelta]' ELSE N' as [71_WritesDelta]' END + N'
		,CASE WHEN ISNULL(LargestWriter,0) = 0 THEN N'''' ELSE SUBSTRING(LargestWriter_money, 1, CHARINDEX(''.'',LargestWriter_money)-1) END' + (CASE WHEN @savespace = N'Y' THEN N' as [LrgWri]' ELSE N' as [72_LargestWriter]' END) + N' 
		,CASE WHEN ISNULL(LogicalReadsDone,0) = 0 THEN N'''' ELSE SUBSTRING(LogicalReadsDone_money, 1, CHARINDEX(''.'',LogicalReadsDone_money)-1) END' + CASE WHEN @savespace = N'Y' THEN N' as [LRds]' ELSE N' as [73_LogicalReadsDone]' END + N'
		,CASE WHEN ISNULL(LogicalReadsDelta,0) = 0 THEN N'''' ELSE SUBSTRING(LogicalReadsDelta_money, 1, CHARINDEX(''.'',LogicalReadsDelta_money)-1) END' + CASE WHEN @savespace = N'Y' THEN N' as [LDelta]' ELSE N' as [74_LogicalReadsDelta]' END + N'
		,CASE WHEN ISNULL(LargestLogicalReader,0) = 0 THEN N'''' ELSE SUBSTRING(LargestLogicalReader_money, 1, CHARINDEX(''.'',LargestLogicalReader_money)-1) END' + CASE WHEN @savespace = N'Y' THEN N' as [LrgLRd]' ELSE N' as [75_LargestLogicalReader]' END + N'
		,CASE WHEN ISNULL(PhysicalReadsDone,0) = 0 THEN N'''' ELSE SUBSTRING(PhysicalReadsDone_money, 1, CHARINDEX(''.'',PhysicalReadsDone_money)-1) END' + CASE WHEN @savespace = N'Y' THEN N' as [PRds]' ELSE N' as [76_PhysicalReadsDone]' END + N'
		,CASE WHEN ISNULL(PhysicalReadsDelta,0) = 0 THEN N'''' ELSE SUBSTRING(PhysicalReadsDelta_money, 1, CHARINDEX(''.'',PhysicalReadsDelta_money)-1) END' + CASE WHEN @savespace = N'Y' THEN N' as [PDelta]' ELSE N' as [77_PhysicalReadsDelta]' END + N'
		,CASE WHEN ISNULL(LargestPhysicalReader,0) = 0 THEN N'''' ELSE SUBSTRING(LargestPhysicalReader_money, 1, CHARINDEX(''.'',LargestPhysicalReader_money)-1) END' + CASE WHEN @savespace = N'Y' THEN N' as [LrgPRd]' ELSE N' as [78_LargestPhysicalReader]' END + N'
		,CASE WHEN ISNULL(BlockingGraph,0) = 0 THEN N'''' ELSE CONVERT(nvarchar(20), BlockingGraph) END' + CASE WHEN @savespace = N'Y' THEN N' as [hasBG]' ELSE N' as [79_HasBlockingGraph]' END + N'
		,CASE WHEN ISNULL(LockDetails,0) = 0 THEN N'''' ELSE CONVERT(nvarchar(20), LockDetails) END ' + CASE WHEN @savespace = N'Y' THEN N' as [hasLck]' ELSE N' as [80_HasLockDetails]' END + N'
		,CASE WHEN ISNULL(TranDetails,0) = 0 THEN N'''' ELSE CONVERT(nvarchar(20), TranDetails) END ' + CASE WHEN @savespace = N'Y' THEN N' as [hasTrnD]' ELSE N' as [81_HasTranDetails]' END + N'
		,SPIDCaptureTime' + CASE WHEN @savespace = N'Y' THEN N' as SCT' ELSE N' as [82_SPIDCaptureTime]' END + N' 
	FROM (
		SELECT *,
			ActLongest_ms_money = CONVERT(nvarchar(20),CONVERT(money,ActLongest_ms),1), 
			ActAvg_ms_money = CONVERT(nvarchar(20),CONVERT(money,ActAvg_ms),1), 
			IdlOpTrnLongest_ms_money = CONVERT(nvarchar(20),CONVERT(money,IdlOpTrnLongest_ms),1),
			IdlOpTrnAvg_ms_money = CONVERT(nvarchar(20),CONVERT(money,IdlOpTrnAvg_ms),1), 
			TranDurLongest_ms_money = CONVERT(nvarchar(20),CONVERT(money,TranDurLongest_ms),1),
			TranDurAvg_ms_money = CONVERT(nvarchar(20),CONVERT(money,TranDurAvg_ms),1), 
			BlockedLongest_ms_money = CONVERT(nvarchar(20),CONVERT(money,BlockedLongest_ms),1),
			BlockedAvg_ms_money = CONVERT(nvarchar(20),CONVERT(money,BlockedAvg_ms),1), 
			WaitingTaskLongest_ms_money = CONVERT(nvarchar(20),CONVERT(money,WaitingTaskLongest_ms),1),
			WaitingTaskAvg_ms_money = CONVERT(nvarchar(20),CONVERT(money,WaitingTaskAvg_ms),1), 
			CPUused_money = CONVERT(nvarchar(20),CONVERT(money,CPUused),1),
			CPUDelta_money = CONVERT(nvarchar(20),CONVERT(money,CPUDelta),1),
			LargestCPUConsumer_money = CONVERT(nvarchar(20),CONVERT(money,LargestCPUConsumer),1),
			WritesDone_money = CONVERT(nvarchar(20),CONVERT(money,WritesDone),1),
			WritesDelta_money = CONVERT(nvarchar(20),CONVERT(money,WritesDelta),1),
			LargestWriter_money = CONVERT(nvarchar(20),CONVERT(money,LargestWriter),1),
			LogicalReadsDone_money = CONVERT(nvarchar(20),CONVERT(money,LogicalReadsDone),1),
			LogicalReadsDelta_money = CONVERT(nvarchar(20),CONVERT(money,LogicalReadsDelta),1),
			LargestLogicalReader_money = CONVERT(nvarchar(20),CONVERT(money,LargestLogicalReader),1),
			PhysicalReadsDone_money = CONVERT(nvarchar(20),CONVERT(money,PhysicalReadsDone),1),
			PhysicalReadsDelta_money =  CONVERT(nvarchar(20),CONVERT(money,PhysicalReadsDelta),1),
			LargestPhysicalReader_money = CONVERT(nvarchar(20),CONVERT(money,LargestPhysicalReader),1)
		FROM PerformanceEye.AutoWho.CaptureSummary t
		WHERE t.SPIDCaptureTime BETWEEN @start AND @end
		) t
	ORDER BY ' + ISNULL(@orderbyColumnName,'error') + N' ' + 
		CASE WHEN @orderdir = N'D' THEN N'desc' ELSE N'' END + 

		CASE WHEN @orderbyColumnName <> N'SPIDCaptureTime' 
			THEN N',SPIDCaptureTime ASC'
			ELSE N'' END + N'
	;
	';

	/* For debugging: 
		SELECT dyntxt, TxtLink
		from (SELECT @DynSQL AS dyntxt) t0
			cross apply (select TxtLink=(select [processing-instruction(q)]=dyntxt
                            for xml path(''),type)) F2
	*/
	EXEC sp_executesql @stmt = @DynSQL,	@params = N'@start DATETIME, @end DATETIME', @start = @start, @end = @end;

	--we always print out at least the EXEC command
	GOTO helpbasic


helpbasic:

	IF @Help <> N'N'
	BEGIN
		IF @Help NOT IN (N'params', N'columns', N'all')
		BEGIN
			--user may have typed gibberish... which is ok, give him/her all the help
			SET @Help = N'all'
		END
	END

	--Because we may have arrived here from a GOTO very early in the proc, we need to set @start & @end
	IF @start IS NULL 
	BEGIN
		SET @start = DATEADD(HOUR, -4, GETDATE());
	END

	IF @end IS NULL
	BEGIN
		SET @end = GETDATE();
	END

	SET @helpexec = REPLACE(@helpexec,'<start datetime>', REPLACE(CONVERT(NVARCHAR(20), @start, 102),'.','-') + ' ' + CONVERT(NVARCHAR(20), @start, 108) + '.' + 
															RIGHT(CONVERT(NVARCHAR(20),N'000') + CONVERT(NVARCHAR(20),DATEPART(MILLISECOND, @start)),3)
							);
	SET @helpexec = REPLACE(@helpexec,'<end datetime>', REPLACE(CONVERT(NVARCHAR(20), @end, 102),'.','-') + ' ' + CONVERT(NVARCHAR(20), @end, 108) + '.' + 
															RIGHT(CONVERT(NVARCHAR(20),N'000') + CONVERT(NVARCHAR(20),DATEPART(MILLISECOND, @end)),3)
							);



	SET @helpstr = @helpexec;
	RAISERROR(@helpstr,10,1) WITH NOWAIT;
	
	IF @Help = N'N'
	BEGIN
		--because the user is likely to use sp_SessionViewer next, if they haven't asked for help explicitly, we print out the syntax for 
		--the Session Viewer procedure

		SET @helpstr = '
EXEC sp_SessionViewer @start=''' + REPLACE(CONVERT(NVARCHAR(20), @start, 102),'.','-') + ' ' + CONVERT(NVARCHAR(20), @start, 108) + '.' + 
		RIGHT(CONVERT(NVARCHAR(20),N'000') + CONVERT(NVARCHAR(20),DATEPART(MILLISECOND, @start)),3) + ''',@end=''' + 
		REPLACE(CONVERT(NVARCHAR(20), @end, 102),'.','-') + ' ' + CONVERT(NVARCHAR(20), @end, 108) + '.' + 
		RIGHT(CONVERT(NVARCHAR(20),N'000') + CONVERT(NVARCHAR(20),DATEPART(MILLISECOND, @end)),3) + ''', --@offset=99999,
	@activity=1, @dur=0,@dbs=N'''',@xdbs=N'''',@spids=N'''',@xspids=N'''',
	@blockonly=N''N'',@attr=N''N'',@resources=N''N'',@batch=N''N'',@plan=N''none'',	--none, statement, full
	@ibuf=N''N'',@bchain=0,@tran=N''N'',@waits=0,		--bchain 0-10, waits 0-3
	@savespace=N''N'',@directives=N''''		--"query(ies)"
	';

		print @helpstr;
		GOTO exitloc
	END

	IF @Help NOT IN (N'params',N'all')
	BEGIN
		GOTO helpcolumns
	END

helpparams:
	SET @helpstr = N'
Parameters
-----------------------------------------------------------------------------------------------------------------------
@start			Valid Values: NULL, any datetime value in the past

				Defines the start time of the time window/range used to pull & display AutoWho capture summaries from
				the AutoWho database. The time cannot be in the future, and must be < @end. If NULL
				is passed, the time defaults to 4 hours before the current time [ DATEADD(hour, -4, GETDATE()) ]
	
@end			Valid Values: NULL, any datetime in the past

				Defines the end time of the time window/range used. The time cannot be in the future, and must be
				> @start. If NULL is passed, the time defaults to 1 second before the current time.
				[ DATEADD(second, -1, GETDATE()) ]

@savespace		Valid Values: ''N'', ''Y''

				If Y, instructs sp_SessionSummary to use abbreviated column names in the output. This is useful for 
				condensing the resulting data set so that more column data can be viewed at the same time.';
	RAISERROR(@helpstr,10,1);

	SET @helpstr = N'
@orderby		Valid Values: positive number (> 0), but <= the number of columns in the result set

				This integer is used in an ORDER BY in the query to define which column is used to order the result set.
				The left-most column in the result set is 1, the next one to the right is 2, etc. The value passed in 
				through @orderby cannot be greater than the number of columns in the table. When @savespace is "N", the
				column number is prepended to the column name in the output, to aid in choosing the correct column number.

@orderdir		Valid Values: A and D

				Determines whether the result set is ordered Ascending or Descending.

@Help			Valid Values: N, params, columns, all, or even gibberish

				If @Help=N''N'', then no help is printed. If =''params'', this section of Help
				is printed. If =''columns'', the section on result columns is prented. If @Help is
				passed anything else (even gibberish), it is set to ''all'', and all help content
				is printed.
	';

	RAISERROR(@helpstr,10,1);

	IF @Help = N'params'
	BEGIN
		GOTO exitloc
	END
	ELSE
	BEGIN
		SET @helpstr = N'
		';
		RAISERROR(@helpstr,10,1);
	END


helpcolumns:

	SET @helpstr = N'
NOTE: AutoWho can be configured to detect certain long-running spids that for whatever reason should not be counted with
other activity. These spids are completely ignored by sp_SessionSummary, which can account for apparent differences between
sp_SessionSummary output and sp_SessionViewer output.

NOTE 2: When the value for a given field is 0 or NULL or otherwise "N/A", an empty cell is returned to make the results less crowded.

Result Columns
-------------------------------------------------------------------------------------------------------------------------
1_SPIDCaptureTime						Short name: SCT

										Displays the datetime value (including milliseconds) of each successful AutoWho 
										capture that occurred during the time range specified by @start/@end. 
										There is always 1 row per capture time. Because of the width of output rows,
										SPIDCaptureTime is displayed multiple times, every few columns, so that the relevant
										time period can always be seen no matter what scrolling is done.

2_TotCapturedSPIDs						Short name: #SPIDs

										The total number of SPIDs captured by AutoWho. Note that this may not have been 
										the total # of spids connected to SQL Server at the time of the AutoWho execution, 
										since AutoWho can be set to filter by duration, database, state of the SPID (active, 
										idle with tran, idle), etc. 

3_Active								Short name: Act

										The # of captured SPIDs that were running a batch at the time of the AutoWho capture. 
										This will always be a subset of the spids represented by "2_TotCapturedSpids".

4_LongestActive							Short name: LngAct

										The MAX(duration) in seconds, of all Active SPIDs. Empty if no spids were captured.
										Format is SS.m for spids that have been active < 1 minute, HH:MM:SS for spids active
										less than a day, and <Day>~HH:MM:SS for spids active longer than a day.
';
	RAISERROR(@helpstr,10,1);

	SET @helpstr = N'
(Active histogram)						A bucket-ized set of columns that shows how many active batches had a duration of 0 to 1 second, 1 to 5 seconds,
										5 to 10 seconds, 10 to 30 seconds, 30 to 60 seconds, 60 to 300 seconds, and 300+ seconds. 
										Durations on the boundary edges are NOT double-counted (a category starts at 1 ms after the boundary point,
										except for the 0 to 1 second category). 

13_IdleWOpenTran						Short name: Idle

										The number of SPIDs that are NOT running a batch, but have dm_exec_sessions.open_transaction_count > 0.

14_LongestIdleWTran						Short name: LngIdl

										The idle duration of the spid that has been idle the longest (of the spids counted by 7_IdleWOpenTran) that also has
										an open tran. Note that this is the length of time that the spid has been IDLE, not the length of time of its
										longest open transaction. A spid that has been idle only a short time could have a long-running transaction.

(Idle w/tran histogram)					Similar to the "active histogram", except that the duration involved is the "idle duration", the amount of time since
										the last batch completed on this spid. Note that this is NOT the duration of the transaction. 
'
	RAISERROR(@helpstr,10,1);

	SET @helpstr = N'
23_wOpenTran							Short name: OpenTran

										The number of spids (whether running or idle) that have an open transaction. Both dm_exec_sessions.open_transaction_count
										and the presence of trans in the dm_tran* views are taken into consideration. Read-only trans will be counted if the
										isolation level is Repeatable Read or Serializable (i.e. able to hold locks for a longer time), or Snapshot Isolation 
										(able to hold open row versions in TempDB). Other trans (e.g. Read Committed) will only be captured if the spid has an
										active or idle w/tran duration >= the TranDetailsThreshold option.

24_LongestTran							Short name: LngTrn

										The duration of the oldest open tran (of spids counted by 11_wOpenTran). This is based on the value in
										dm_tran_active_transactions.transaction_begin_time.

(transaction histogram)					Similar to the "active histogram", except that the duration involved is the transaction length. 
	';

	RAISERROR(@helpstr,10,1);

	SET @helpstr = N'
33_Blocked								Short name: Blkd

										The # of spids that are actively running a batch but are blocked by another spid. A "blocked" spid here means
										that at least one of the tasks for the spid in dm_os_waiting_tasks has a non-null blocking_session_id value 
										which is also <> session_id. (Thus, CXPACKET waits are not considered "blocking"). Certain types of page latch
										waits and even RESOURCE_SEMAPHORE waits can fit this category. 

34_LongBlkTask							Short name: LngBlk

										The MAX(dm_os_waiting_tasks.wait_duration_ms) value for spids that qualify for the "16_Blocked" field. 

(blocked histogram)						Similar to the "active histogram", except that the duration involved is the MAX(dm_os_waiting_tasks.wait_duration_ms),
										per SPID, of blocked tasks.
'
	RAISERROR(@helpstr,10,1);

	SET @helpstr = N'
243_WaitingSPIDs							Short name: WtSPIDs

										The # of spids that are actively running a batch and are waiting but NOT blocked by another spid. Note
										that normally, SQL Server terminology defines blocking as one type (i.e. a subset) of waiting. However,
										in sp_SessionSummary, the two categories are non-overlapping, allowing the user to quickly see which type
										of "slowdown" has occurred. Note that CXPACKET waits are not considered waits. Thus, a parallel query
										with 16 tasks, 1 of which is running and the other 15 are waiting on CXPACKET, will not be considered to 
										be waiting (or blocked). If that 1 task then becomes blocked (and the other 15 are still waiting on CXPACKET),
										then the spid will be "blocked" but not waiting.

44_WaitingTasks							Short name: WtTsk

										The # of tasks that are waiting in actively-running spids. For example, a parallel query might have 9 tasks,
										4 of which are running, 3 are in CXPACKET waits, and 2 are waiting on PAGEIOLATCH waits. This spid would
										only increment the "20_WaitingSPIDs" field by 1 (one spid), but would increment the "21_WaitingTasks" field
										by 2 because of the 2 PAGEIOLATCH waits. (CXPACKET waits are not counted as "waits"). As mentioned above, 
										tasks that are blocked behind another spid do not count as waiting.

45_LongestWaitTask						Short name: LngWtTsk

										The MAX(dm_os_waiting_tasks.wait_duration_ms) value of tasks that are waiting. Waits due to being blocked
										behind another spid and CXPACKET waits are not counted as waiting.
										
	';

	RAISERROR(@helpstr,10,1);

	SET @helpstr = N'
(waiting histogram)						Similar to the "active histogram", except that the duration involved is the MAX(dm_os_waiting_tasks.wait_duration_ms)
										of tasks that are waiting, per-spid,(but not blocked). 

54_TLogUsed								Short name: Tlog

										The SUM() of the 2 dm_tran_database_transactions.database_transaction_log_bytes_reserved* columns for all user
										transactions. Note that logic is in place to prevent enlisted trans from being double-counted.
										
55_LargestLogWriter (MB)				Short name: LargestTlog

										The MAX() of the addition of the 2 dm_tran_database_transactions.database_transaction_log_bytes_reserved* columns
										for all user transactions.

56_QueryMem (MB)						Short name: Qmem

										The SUM() of dm_exec_memory_grants.requested_memory_kb. Note that this field uses requested_memory_kb, while
										the next field uses granted_memory_kb. The idea here is for field 28 to show how much memory is NEEDED across
										all queries, while field 29 shows what the largest requester got.

57_LargestQueryMem (MB)					Short name: LargestQM

										The MAX() of dm_exec_memory_grants.granted_memory_kb. See notes on "28_QueryMem (MB)".

58_TempDB (MB)							Short name: Tdb

										The SUM() of the various tempdb session and task allocation (minus deallocation) counters from 
										dm_db_session_space_usage and dm_db_task_space_usage. If a given "alloc - dealloc" pair yields a 
										negative number, this pair is "floored" at 0. 

59_LargestTempDB (MB)					Short name: LargestTdb

										The MAX() by spid of the tempdb allocation minus deallocation counters.
	';

	RAISERROR(@helpstr,10,1);

	SET @helpstr = N'
60_CPUused								Short name: CPU

										The SUM() of dm_exec_requests.cpu_time. Only active spids are counted so that long-idle spids do not influence
										(i.e. flatten) the up-and-down nature of this value over time.

61_CPUDelta								Short name: CDelta

										"32_CPUused" of the current row minus "32_CPUused" from the previous row. Negative values are ignored, leaving
										an empty cell.

62_LargestCPUConsumer					Short name: LrgCPU

										The MAX() of dm_exec_requests.cpu_time

63_AllocatedTasks						Short name: Tasks

										A SUM() of the # of tasks allocated for each spid. The # of tasks for a spid is calcuted as the COUNT(*) of
										records in sys.dm_db_task_space_usage.

65_WritesDone							Short name: Wri

										SUM() on dm_exec_requests.writes, otherwise similar to "32_CPUused"

66_WritesDelta							Short name: WDelta

										Similar to "33_CPUDelta", but for writes

67_LargestWriter						Short name: LrgWri

										MAX() of dm_exec_requests.writes
	';

	RAISERROR(@helpstr,10,1);

	SET @helpstr = N'
68_LogicalReadsDone						Short name: LRds

										SUM() on dm_exec_requests.logical_reads, otherwise similar to "32_CPUused"

69_LogicalReadsDelta					Short name: LDelta

										Similar to "33_CPUDelta", but for logical reads.

70_LargestLogicalReader					Short name: LrgLRd

										MAX() of dm_exec_requests.logical_reads

71_PhysicalReadsDone					Short name: PRds

										SUM() of dm_exec_requests.reads, otherwise similar to "32_CPUused"

72_PhysicalReadsDelta					Short name: PDelta

										Similar to "33_CPUDelta", but for physical reads

73_LargestPhysicalReader				Short name: LrgPRd

										MAX() of dm_exec_requests.reads

74_HasBlockingGraph						Short name: hasBG

										Indicates whether AutoWho constructed a blocking graph for its run at the time indicated by SPIDCaptureTime.

75_HasLockDetails						Short name: hasLck

										Indicates whether AutoWho collected details from dm_tran_locks about blocker & blockee spids during its run at the time
										indicated by SPIDCaptureTime.

76_HasTranDetails						Short name: hasTrnD

										Indicates whether AutoWho collected details from the dm_tran* (besides dm_tran_locks) views during its run at the time
										indicated by SPIDCaptureTime.
	';

	RAISERROR(@helpstr,10,1);

exitloc:

	RETURN 0;
END

GO
