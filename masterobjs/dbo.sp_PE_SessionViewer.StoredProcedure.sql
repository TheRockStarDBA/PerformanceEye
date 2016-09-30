USE [master]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[sp_PE_SessionViewer]
/*   
	PROCEDURE:		sp_PE_SessionViewer

	AUTHOR:			Aaron Morelli
					email@TBD.com
					@sqlcrossjoin
					sqlcrossjoin.wordpress.com
					https://github.com/amorelli005/PerformanceEye

	PURPOSE: sp_SessionViewer is a "DMV aggregator", and displays the contents of a number of session-focused DMVs
		in one concise, intuitive interface. A large number of proc parameters and "clickable-XML" fields allow for
		a "drilldown" experience when the user wants to see data points that don't fit well onto a smaller amount
		of SSMS grid real estate. When running this tool, the user can view all non-trivial sessions that are active
		or (in the case when running in historical mode) were active at the time that AutoWho ran a snapshot.

		*NOTE: as of 2016-04-26, sp_SessionViewer can *only* view historical data (collected by AutoWho); at some
		time in the future, this procedure will be able to view the current state of the system. 

		Detailed help documentation is available by running
			EXEC sp_PE_SessionViewer @help=N'Y'


	FUTURE ENHANCEMENTS: 
		When I implement the live version of this proc, add several directives:
		- one to cause page latch resolution
		- one to suppress system spids
		- one to "showself" (i.e. not suppress @@spid)
		- one to do the "debug speed" logic, or equivalent for the "Messages" tab

    CHANGE LOG:	2015-09-01  Aaron Morelli		Development Start
				2016-09-26	Aaron Morelli		Final run-through and commenting


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
EXEC dbo.sp_PE_SessionViewer @start='2015-11-02',@end='2015-11-03', --@offset=99999,	--99999
	@activity=1, @dur=0,@dbs=N'',@xdbs=N'',@spids=N'',@xspids=N'',
	@blockonly=N'N',@attr=N'N',@resources=N'N',@batch=N'N',@plan=N'none',	--none, statement, full
	@ibuf=N'N',@bchain=0,@tran=N'N',@waits=0,		--bchain 0-10, waits 0-3
	@savespace=N'N',@directives=N'', @help=N'N'		--"query(ies)", "basedata"

*/
(
	@start			DATETIME=NULL,			--if null, query live system (NOT YET IMPLEMENTED)
	@end			DATETIME=NULL,
	@offset			INT=99999,

	--filter variables
	@activity		TINYINT=1,				--0 = Running only, 1 = Active + idle-open-tran, 2 = everything
	@dur			INT = 0,				--milliseconds
	@dbs			NVARCHAR(512)=N'',		--these 4 variables are comma-separated lists
	@xdbs			NVARCHAR(512)=N'',
	@spids			NVARCHAR(100)=N'',
	@xspids			NVARCHAR(100)=N'',
	@blockonly		NCHAR(1)=N'N',			--only include blockers and blockees

	--auxiliary info
	@attr			NCHAR(1)=N'N',			--Session & Connection attributes
	@resources		NCHAR(1)=N'N',			--TempDB, memory, reads/writes, CPU info
	@batch			NCHAR(1)=N'N',			-- Include the full SQL batch text. For historical data, only available if AutoWho was set to collect it.
	@plan			NVARCHAR(20)=N'none',	--"none", "statement", "full"		For historical data, ""statement" and "full" are only available if AutoWho was set to collect it.
	@ibuf			NCHAR(1)=N'N',
	@bchain			TINYINT=0,				-- how many levels of the blocking chain data to show, if there is any to show. For historical data, only available if AutoWho captured it.
	@tran			NCHAR(1)=N'N',			-- Show transactions related to the spid.
	@waits			TINYINT=0,			--0 basic info; 1 adds more info about lock and latch waits; 
										-- 2 adds still more info for lock and latch waits; 3 displays aggregated dm_tran_locks data for SPIDs that were blockers or blockees, if available

	--Other options
	@savespace		NCHAR(1)=N'N',
	@directives		NVARCHAR(512)=N'',		--Allows extra directions to be passed to sp_SessionViewer, such as to alter formatting or include certain extra columns or column values.
											-- At this time, there are 2: "query/queries", which prints a number of potentially useful queries for ad-hoc analysis, and "basedata", which
											-- returns the collected data before it undergoes heavy formatting. This is useful for debugging.
	@help			NVARCHAR(10)=N'N'		--params, columns, all
)
AS
BEGIN
	SET NOCOUNT ON;
	SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED;
	SET ANSI_PADDING ON;

	DECLARE @lv__ViewCurrent						BIT,
			@lv__HistoricalSPIDCaptureTime			DATETIME,
			@lv__effectiveordinal					INT,
			@scratch__int							INT,
			@helpstr								NVARCHAR(MAX),
			@helpexec								NVARCHAR(4000),
			@scratch__nvarchar						NVARCHAR(MAX),
			@err__msg								NVARCHAR(MAX),
			@lv__DynSQL								NVARCHAR(MAX),
			@lv__OptionsHash_str					NVARCHAR(4000),
			@lv__OptionsHash						VARBINARY(64),
			@lv__LastOptionsHash					VARBINARY(64)
			;

	SET @helpexec = N'
EXEC dbo.sp_SessionViewer @start=''<start datetime>'',@end=''<end datetime>'', --@offset=99999,	--99999
	@activity=1, @dur=0,@dbs=N'''',@xdbs=N'''',@spids=N'''',@xspids=N'''',
	@blockonly=N''N'',@attr=N''N'',@resources=N''N'',@batch=N''N'',@plan=N''none'',	--none, statement, full
	@ibuf=N''N'',@bchain=0,@tran=N''N'',@waits=0,		--bchain 0-10, waits 0-3
	@savespace=N''N'',@directives=N'''', @help=N''N''		--"query(ies)"

	';

	IF ISNULL(@Help,N'z') <> N'N'
	BEGIN
		GOTO helpbasic
	END

	DECLARE @lv__SQLVersion NVARCHAR(10);
	SELECT @lv__SQLVersion = (
	SELECT CASE
			WHEN t.col1 LIKE N'8%' THEN N'2000'
			WHEN t.col1 LIKE N'9%' THEN N'2005'
			WHEN t.col1 LIKE N'10.5%' THEN N'2008R2'
			WHEN t.col1 LIKE N'10%' THEN N'2008'
			WHEN t.col1 LIKE N'11%' THEN N'2012'
			WHEN t.col1 LIKE N'12%' THEN N'2014'
			WHEN t.col1 LIKE N'13%' THEN N'2016'
		END AS val1
	FROM (SELECT CONVERT(SYSNAME, SERVERPROPERTY(N'ProductVersion')) AS col1) AS t);


	DECLARE @dir__shortcols BIT;
	SET @dir__shortcols = CONVERT(BIT,0);

	DECLARE @lv__UtilityName NVARCHAR(30);
	SET @lv__UtilityName = N'SessionViewer'

	IF @start IS NULL
	BEGIN
		IF @end IS NOT NULL
		BEGIN
			SET @helpexec = REPLACE(@helpexec,'<end datetime>', REPLACE(CONVERT(NVARCHAR(20), @end, 102),'.','-') + ' ' + CONVERT(NVARCHAR(20), @end, 108) + '.' + 
															RIGHT(CONVERT(NVARCHAR(20),N'000') + CONVERT(NVARCHAR(20),DATEPART(MILLISECOND, @end)),3)
							);
			RAISERROR(@helpexec,10,1);
			RAISERROR('Parameter @end cannot have a value when @start is NULL or unspecified',16,1);
			RETURN -1;
		END

		--ok, so both are NULL. This run will look at live SQL DMVs
		SET @lv__ViewCurrent = CONVERT(BIT,1);
		SET @lv__effectiveordinal = NULL;	--n/a to current runs
	END
	ELSE
	BEGIN
		--ok, @start is non-null

		--Put @start's value into our helpexec string
		SET @helpexec = REPLACE(@helpexec,'<start datetime>', REPLACE(CONVERT(NVARCHAR(20), @start, 102),'.','-') + ' ' + CONVERT(NVARCHAR(20), @start, 108) + '.' + 
														RIGHT(CONVERT(NVARCHAR(20),N'000') + CONVERT(NVARCHAR(20),DATEPART(MILLISECOND, @start)),3)
						);

		IF @end IS NULL
		BEGIN
			RAISERROR(@helpexec,10,1);
			RAISERROR('Parameter @end must have a value when @start has been specified.',16,1);
			RETURN -1;
		END

		--@end is also NOT NULL. Put it into our helpexec string
		SET @helpexec = REPLACE(@helpexec,'<end datetime>', REPLACE(CONVERT(NVARCHAR(20), @end, 102),'.','-') + ' ' + CONVERT(NVARCHAR(20), @end, 108) + '.' + 
															RIGHT(CONVERT(NVARCHAR(20),N'000') + CONVERT(NVARCHAR(20),DATEPART(MILLISECOND, @end)),3)
							);

		IF @start >= GETDATE() OR @end >= GETDATE()
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

		SET @lv__ViewCurrent = CONVERT(BIT,0);

		--@offset must be specified for historical runs
		IF @offset IS NULL
		BEGIN
			RAISERROR(@helpexec,10,1);
			RAISERROR('Parameter @offset must be non-null when @start is non-null.',16,1);
			RETURN -1;
		END
		ELSE
		BEGIN
			SET @lv__effectiveordinal = @offset;
		END
	END		--IF @start IS NULL


	IF @activity IS NULL
	BEGIN
		RAISERROR(@helpexec,10,1);
		RAISERROR('Parameter @activity cannot be NULL. Valid values are 0 (running SPIDs only), 1 (running plus idle w/tran), and 2 (all)',16,1);
		RETURN -1;
	END

	IF @activity NOT IN (0,1,2)
	BEGIN
		RAISERROR(@helpexec,10,1);
		RAISERROR('Parameter @activity must be 0 (running SPIDs only), 1 (running plus idle w/tran), or 2 (all)',16,1);
		RETURN -1;
	END

	IF ISNULL(@dur,-1) < 0
	BEGIN
		RAISERROR(@helpexec,10,1);
		RAISERROR('Parameter @dur cannot be NULL, and must be >= 0',16,1);
		RETURN -1;
	END

	IF ISNULL(@blockonly,N'z') NOT IN (N'Y', N'N')
	BEGIN
		RAISERROR(@helpexec,10,1);
		RAISERROR('Parameter @blockonly cannot be NULL, and must be either Y or N',16,1);
		RETURN -1;
	END

	--We don't validate the 4 CSV filtering variables, outside of checking for NULL. We leave that to the sub-procs
	IF @dbs IS NULL
	BEGIN
		RAISERROR(@helpexec,10,1);
		RAISERROR('Parameter @dbs cannot be NULL; it should either be an empty string or a comma-delimited string of database names.',16,1);
		RETURN -1;
	END

	IF @xdbs IS NULL
	BEGIN
		RAISERROR(@helpexec,10,1);
		RAISERROR('Parameter @xdbs cannot be NULL; it should either be an empty string or a comma-delimited string of database names.',16,1);
		RETURN -1;
	END

	IF @spids IS NULL
	BEGIN
		RAISERROR(@helpexec,10,1);
		RAISERROR('Parameter @spids cannot be NULL; it should either be an empty string or a comma-delimited string of session IDs.',16,1);
		RETURN -1;
	END

	IF @xspids IS NULL
	BEGIN
		RAISERROR(@helpexec,10,1);
		RAISERROR('Parameter @xspids cannot be NULL; it should either be an empty string or a comma-delimited string of session IDs.',16,1);
		RETURN -1;
	END

	IF UPPER(ISNULL(@attr,N'z')) NOT IN (N'N', N'Y')
	BEGIN
		RAISERROR(@helpexec,10,1);
		RAISERROR('Parameter @attr must be either Y or N',16,1);
		RETURN -1;
	END

	IF UPPER(ISNULL(@resources,N'z')) NOT IN (N'N', N'Y')
	BEGIN
		RAISERROR(@helpexec,10,1);
		RAISERROR('Parameter @resources must be either Y or N',16,1);
		RETURN -1;
	END

	IF UPPER(ISNULL(@batch,N'z')) NOT IN (N'N', N'Y')
	BEGIN
		RAISERROR(@helpexec,10,1);
		RAISERROR('Parameter @batch must be either Y or N',16,1);
		RETURN -1;
	END

	IF LOWER(ISNULL(@plan,N'z')) NOT IN (N'none', N'statement', N'full')
	BEGIN
		RAISERROR(@helpexec,10,1);
		RAISERROR('Parameter @plan must be either "none", "statement", or "full"',16,1);
		RETURN -1;
	END

	IF UPPER(ISNULL(@ibuf,N'z')) NOT IN (N'N', N'Y')
	BEGIN
		RAISERROR(@helpexec,10,1);
		RAISERROR('Parameter @ibuf must be either Y or N',16,1);
		RETURN -1;
	END

	IF @bchain IS NULL
	BEGIN
		RAISERROR(@helpexec,10,1);
		RAISERROR('Parameter @bchain cannot be NULL, and must be >= 0',16,1);
		RETURN -1;
	END

	IF @bchain < 0
	BEGIN
		RAISERROR(@helpexec,10,1);
		RAISERROR('Parameter @bchain must be >= 0',16,1);
		RETURN -1;
	END

	IF UPPER(ISNULL(@tran,N'z')) NOT IN (N'N', N'Y')
	BEGIN
		RAISERROR(@helpexec,10,1);
		RAISERROR('Parameter @tran must be either Y or N',16,1);
		RETURN -1;
	END

	IF ISNULL(@waits,255) NOT IN (0,1,2,3)
	BEGIN
		RAISERROR(@helpexec,10,1);
		RAISERROR('Parameter @waits must be either 0 (default), 1, 2, or 3.',16,1);
		RETURN -1;
	END

	IF UPPER(ISNULL(@savespace,N'z')) NOT IN (N'N', N'Y')
	BEGIN
		RAISERROR(@helpexec,10,1);
		RAISERROR('Parameter @savespace must be either Y or N',16,1);
		RETURN -1;
	END

	CREATE TABLE #HistoricalCaptureTimes (
		hct		DATETIME NOT NULL	PRIMARY KEY CLUSTERED
	);

	/* The below comment is only applicable to sp_SessionViewer in historical mode.
	
	We have two different, but related, caching tables. 
		CorePE.OrdinalCachePosition --> acts a bit like a cursor in the sense that it keeps track of which "position" the user is at currently w/sp_SessionViewer.
			An "ordinal cache" has a key of Utility name/Start Time/End Time/session_id (spid of of the user running sp_SessionViewer). As the user repeatedly
			presses F5, the position increments by 1 each time (or decrements, if @offset = -99999 instead of the default 99999), and this incrementation is stored
			in the OrdinalCachePosition table.

		CorePE.CaptureOrdinalCache --> When a user first enters a Utility/Start Time/End Time (in this case, Utility="SessionViewer" of course), the CaptureOrdinalCache
			is populated with all of the SPID Capture Time values between Start Time and End Time. That is, the CorePE.CaptureOrdinalCache will hold every run of
			the AutoWho.Collector between @start and @end inclusive. All of those capture times are numbered from 1 to X and -1 to -X, in time order ascending and descending.
			Thus, given a number (e.g. 5), the table can be used to find the SPID Capture Time that is the 5th one in the series of captures starting with the first one >= @Start time,
			and ending with the last one <= @End time. Or if the number is -5, the table can be used to obtain the SPID Capture Time that is 5th from @End, going backwards towards @Start.
			
			As mentioned above, each time the user hits F5 to execute the proc, the position in CorePE.OrdinalCachePosition is incremented and returned/store in @lv__effectiveordinal,
			and then this position is used to probe into CorePE.CaptureOrdinalCache to find the SPID Capture Time that corresponds to the @lv__effectiveordinal in the @start/@end
			time series that has been specified. 

		This complicated design behind the scenes is to give the user a relatively simple experience in moving through time when examining AutoWho data.
	*/

	--The below code block sets/updates CorePE.OrdinalCachePosition appropriately.
	-- If @offset = 0, we simply ignore the position logic completely. This has the nice effect of
	-- letting the user start with a position, switch to @offset=0 partway through, then switch back to the position they were on
	-- seamlessly.
	IF @lv__ViewCurrent = CONVERT(BIT,0) AND @offset <> 0
	BEGIN
		--this is a historical run, so let's get our "effective position". A first-time run creates a position marker entry in the cache table,
		-- a follow-up run modifies the position marker.

		--Aaron 2016-05-28: We create a string of all the options used and then hash it.
		-- We exclude @start & @end b/c they are keys, and we exclude @offset in case the only
		-- option changed was @offset (from 99999 to -99999 or vice versa)

		SET @lv__OptionsHash_str = 
			/*
			N'@start=' + 
			CASE WHEN @start IS NULL THEN N'NULL' 
			ELSE REPLACE(CONVERT(NVARCHAR(20), @start, 102),'.','-')+' '+CONVERT(NVARCHAR(20), @start, 108)+'.'+RIGHT(CONVERT(NVARCHAR(20),N'000')+CONVERT(NVARCHAR(20),DATEPART(MILLISECOND, @start)),3)
			END + N',@end=' + 
			CASE WHEN @end IS NULL THEN N'NULL' 
			ELSE REPLACE(CONVERT(NVARCHAR(20), @end, 102),'.','-')+' '+CONVERT(NVARCHAR(20), @end, 108)+'.'+RIGHT(CONVERT(NVARCHAR(20),N'000')+CONVERT(NVARCHAR(20),DATEPART(MILLISECOND, @end)),3)
			END + 
			N',@offset=' +		ISNULL(CONVERT(nvarchar(20), @offset),N'NULL') + 
			*/
			N',@activity=' +	ISNULL(CONVERT(nvarchar(20),@activity),N'NULL') + 
			N',@dur=' +			ISNULL(CONVERT(nvarchar(20),@dur),N'NULL') + 
			N',@dbs=' +			ISNULL(@dbs,N'NULL') + 
			N',@xdbs=' +		ISNULL(@xdbs,N'NULL') + 
			N',@xspids=' +		ISNULL(@xspids,N'NULL') + 
			N',@blockonly=' +	ISNULL(@blockonly,N'z') + 
			N',@attr='+			ISNULL(@attr,N'z') + 
			N',@resources=' +	ISNULL(@resources,N'z') + 
			N',@batch=' +		ISNULL(@batch,N'z') + 
			N',@plan=' +		ISNULL(@plan,N'NULL') + 
			N',@ibuf=' +		ISNULL(@ibuf,N'z') + 
			N',@bchain=' +		ISNULL(CONVERT(nvarchar(20),@bchain),N'NULL') + 
			N',@tran=' +		ISNULL(@tran,N'z') + 
			N',@waits=' +		ISNULL(CONVERT(nvarchar(20),@waits),N'NULL') + 
			N',@savespace=' +	ISNULL(@savespace,N'z') + 
			N',@directives=' +	ISNULL(@directives,N'NULL') + 
			N',@help=' +		ISNULL(@help,N'NULL')
			;

		IF @lv__SQLVersion IN (N'2016')
		BEGIN
			--SHA1 is deprecated in 2016
			SET @lv__OptionsHash = HASHBYTES('SHA2_256',@lv__OptionsHash_str); 
		END
		ELSE
		BEGIN
			SET @lv__OptionsHash = HASHBYTES('SHA1',@lv__OptionsHash_str); 
		END

		IF NOT EXISTS (
			SELECT * FROM PerformanceEye.CorePE.OrdinalCachePosition t WITH (NOLOCK)
			WHERE t.Utility = @lv__UtilityName
			AND t.StartTime = @start
			AND t.EndTime = @end 
			AND t.session_id = @@SPID 
			)
		BEGIN
			INSERT INTO PerformanceEye.CorePE.OrdinalCachePosition
			(Utility, StartTime, EndTime, session_id, CurrentPosition, LastOptionsHash)
			SELECT @lv__UtilityName, @start, @end, @@SPID, 
				CASE WHEN @offset = 99999 THEN 1
					WHEN @offset = -99999 THEN -1
					ELSE @offset END,
				@lv__OptionsHash;
		END
		ELSE
		BEGIN	--cache already exists, so someone has already run with this start/endtime before on this spid
			--If @offset = 99999, we want to increment by 1		(Aaron 2016-05-28: but only if the Options Hash is the same)
			--If @offset = -99999, we want to decrement by 1		""
			--If offset = a different value, we want to set the position to that value
			SELECT 
				@lv__LastOptionsHash = LastOptionsHash
			FROM PerformanceEye.CorePE.OrdinalCachePosition
			WHERE Utility = @lv__UtilityName
			AND StartTime = @start
			AND EndTime = @end 
			AND session_id = @@SPID
			;

			IF @lv__LastOptionsHash <> @lv__OptionsHash
			BEGIN
				--user changed the options in some way. Retain the some position if @offset wasn't explicit
				UPDATE PerformanceEye.CorePE.OrdinalCachePosition
				SET LastOptionsHash = @lv__OptionsHash, 
					CurrentPosition = CASE WHEN @offset IN (99999,-99999) THEN CurrentPosition
										ELSE @offset
										END
				WHERE Utility = @lv__UtilityName
				AND StartTime = @start
				AND EndTime = @end 
				AND session_id = @@SPID
				;
			END
			ELSE
			BEGIN
				--options stayed the same, in/decrement the position
				UPDATE PerformanceEye.CorePE.OrdinalCachePosition
				SET CurrentPosition = CASE 
					WHEN @offset = 99999 THEN CurrentPosition + 1
					WHEN @offset = -99999 THEN CurrentPosition - 1
					ELSE @offset 
					END
				WHERE Utility = @lv__UtilityName
				AND StartTime = @start
				AND EndTime = @end 
				AND session_id = @@SPID
				;
			END 
		END		--if cache already exists

		SELECT @lv__effectiveordinal = t.CurrentPosition
		FROM PerformanceEye.CorePE.OrdinalCachePosition t
		WHERE t.Utility = @lv__UtilityName 
		AND t.StartTime = @start
		AND t.EndTime = @end 
		AND t.session_id = @@SPID
		;
	END		--if historical run and @offset <> 0



	--is it a historical run?
	IF @lv__ViewCurrent = CONVERT(BIT, 0)
	BEGIN
		--Regardless of whether we are pulling for a specific ordinal or pulling for a range, we need to ensure
		-- that the capture summary table has all appropriate data in the time range. However, the way that CaptureSummary
		-- is populated differs slightly based on the value in @offset/@lv__effectiveordinal. 
		--If @offset=0, the user isn't using the position marker cache, so we don't need to interact at all with CorePE.OrdinalCachePosition.
		-- In fact, we don't need to interact at all with CorePE.CaptureOrdinalCache either, as we just need a list of all the 
		-- SPIDCaptureTime values between @start and @end. So, for @offset=0, we just check that CaptureSummary is up-to-date 
		-- based on AutoWho.CaptureTimes rows, and then we pull a list of datetime values.
		IF @lv__effectiveordinal = 0
		BEGIN
			IF EXISTS (SELECT * FROM PerformanceEye.AutoWho.CaptureTimes ct
						WHERE ct.RunWasSuccessful = 1
						AND ct.CaptureSummaryPopulated = 0
						AND ct.SPIDCaptureTime BETWEEN @start AND @end)
			BEGIN
				EXEC @scratch__int = PerformanceEye.AutoWho.PopulateCaptureSummary @StartTime = @start, @EndTime = @end; 
					--returns 1 if no rows were found in the range
					-- -1 if there was an unexpected exception
					-- 0 if success

				IF @scratch__int = 1
				BEGIN
					--no rows for this range. Return special code 2 and let the caller decide what to do
					RAISERROR(@helpexec,10,1);
					RAISERROR('
			There is no AutoWho data for the time range specified.',10,1);
					RETURN 1;
				END

				IF @scratch__int < 0
				BEGIN
					SET @err__msg = 'Unexpected error occurred while retrieving the AutoWho data. More info is available in the AutoWho log under the tag "SummCapturePopulation" or contact your administrator.'
					RAISERROR(@err__msg, 16, 1);
					RETURN -1;
				END
			END
		END
		ELSE
		BEGIN
			--However, if @offset is <> 0, then @lv__effectiveordinal is either < 0 or > 0. In that case, we DO need to 
			-- interact with the position cache and, since we're after just one SPID Capture Time value referenced by an offset, 
			-- the CaptureOrdinalCache table. But again, there's no guarantee that the CaptureSummary table is up-to-date.
			-- Thus, we rely on the fact that the procedure CorePE.RetrieveOrdinalCacheEntry will populate the CaptureSummary
			-- and the CaptureOrdinalCache if those are not yet populated. 

			SET @lv__HistoricalSPIDCaptureTime = NULL;

			--First, optimistically assume that the cache already exists, and grab the ordinal's HCT
			IF @lv__effectiveordinal < 0 
			BEGIN
				SELECT @lv__HistoricalSPIDCaptureTime = c.CaptureTime
				FROM PerformanceEye.CorePE.CaptureOrdinalCache c
				WHERE c.Utility = @lv__UtilityName
				AND c.StartTime = @start
				AND c.EndTime = @end
				AND c.OrdinalNegative = @lv__effectiveordinal;
			END
			ELSE IF @lv__effectiveordinal > 0
			BEGIN
				SELECT @lv__HistoricalSPIDCaptureTime = c.CaptureTime
				FROM PerformanceEye.CorePE.CaptureOrdinalCache c
				WHERE c.Utility = @lv__UtilityName
				AND c.StartTime = @start
				AND c.EndTime = @end
				AND c.Ordinal = @lv__effectiveordinal;
			END

			--If still NULL, the cache may not exist, or the ordinal is out of range. 
			IF @lv__HistoricalSPIDCaptureTime IS NULL
			BEGIN
				SET @scratch__int = NULL;
				SET @scratch__nvarchar = NULL;
				EXEC @scratch__int = PerformanceEye.CorePE.RetrieveOrdinalCacheEntry @ut = @lv__UtilityName, @st=@start, 
					@et=@end, @ord=@lv__effectiveordinal, @hct=@lv__HistoricalSPIDCaptureTime OUTPUT, @msg=@scratch__nvarchar OUTPUT;

					--returns 0 if successful,
					-- -1 if exception occurred
					-- 1 if the ordinal passed is out-of-range
					-- 2 or 3 if there is no AutoWho data for the time range specified

				IF @scratch__int IS NULL OR @scratch__int < 0
				BEGIN
					SET @err__msg = 'Unexpected error occurred while retrieving the AutoWho data. More info is available in the AutoWho log under the tag "RetrieveOrdinalCache", or contact your administrator.'
					RAISERROR(@err__msg, 16, 1);
					RETURN -1;
				END

				IF @scratch__int = 1
				BEGIN
					IF @scratch__nvarchar IS NULL
					BEGIN
						IF @lv__effectiveordinal < 0
						BEGIN
							SET @scratch__nvarchar = N'The value passed in for parameter @offset is out of range. Try a larger (closer to zero) value.';
						END
						ELSE
						BEGIN
							SET @scratch__nvarchar = N'The value passed in for parameter @offset is out of range. Try a smaller value.';
						END
					END

					RAISERROR(@helpexec,10,1);
					RAISERROR(@scratch__nvarchar,16,1);
					RETURN -1;
				END

				IF @scratch__int IN (2,3)
				BEGIN
					RAISERROR(@helpexec,10,1);
					RAISERROR('There is no AutoWho data for the time range specified.',10,1);
					RETURN 1;
				END
			END		--IF @lv__HistoricalSPIDCaptureTime IS NULL
		END		--IF @lv__effectiveordinal = 0

		--PRINT CONVERT(VARCHAR(20),@lv__HistoricalSPIDCaptureTime,102) + ' ' + CONVERT(VARCHAR(20), @lv__HistoricalSPIDCaptureTime,108);
		--return 0;


		IF @lv__effectiveordinal = 0
		BEGIN
			INSERT INTO #HistoricalCaptureTimes (
				hct
			)
			SELECT DISTINCT cs.SPIDCaptureTime
			FROM PerformanceEye.AutoWho.CaptureSummary cs
			WHERE cs.SPIDCaptureTime BETWEEN @start AND @end;

			DECLARE iterateHCTs CURSOR FOR
			SELECT hct 
			FROM #HistoricalCaptureTimes
			ORDER BY hct ASC;

			OPEN iterateHCTs;
			FETCH iterateHCTs INTO @lv__HistoricalSPIDCaptureTime;

			WHILE @@FETCH_STATUS = 0
			BEGIN
				EXEC PerformanceEye.AutoWho.ViewHistoricalSessions @hct = @lv__HistoricalSPIDCaptureTime, 
					@activity = @activity,
					@dur = @dur,
					@db = @dbs,
					@xdb = @xdbs,
					@spid = @spids,
					@xspid = @xspids,
					@blockonly = @blockonly,
					@attr = @attr,
					@resource = @resources,
					@batch = @batch,
					@plan = @plan,
					@ibuf = @ibuf,
					@bchain = @bchain,
					@waits = @waits,
					@tran = @tran,
					@savespace = @savespace,
					@effectiveordinal = @lv__effectiveordinal,
					@dir = @directives;

				FETCH iterateHCTs INTO @lv__HistoricalSPIDCaptureTime;
			END

			CLOSE iterateHCTs;
			DEALLOCATE iterateHCTs;
		END
		ELSE
		BEGIN
			--Just executing for 1 SPID Capture time
			EXEC PerformanceEye.AutoWho.ViewHistoricalSessions @hct = @lv__HistoricalSPIDCaptureTime, 
				@activity = @activity,
				@dur = @dur,
				@db = @dbs,
				@xdb = @xdbs,
				@spid = @spids,
				@xspid = @xspids,
				@blockonly = @blockonly,
				@attr = @attr,
				@resource = @resources,
				@batch = @batch,
				@plan = @plan,
				@ibuf = @ibuf,
				@bchain = @bchain,
				@waits = @waits,
				@tran = @tran,
				@savespace = @savespace,
				@effectiveordinal = @lv__effectiveordinal,
				@dir = @directives;
		END	--IF @lv__effectiveordinal = 0

		--we always print out at least the EXEC command
		GOTO helpbasic
	END
	ELSE
	BEGIN
		IF has_perms_by_name(null, null, 'VIEW SERVER STATE') <> 1
		BEGIN
			RAISERROR(@helpexec,10,1);
			RAISERROR(N'The VIEW SERVER STATE permission (or permissions/role membership that include VIEW SERVER STATE) is required to execute sp_SessionViewer. Exiting...', 11,1);
			RETURN -1;
		END
		ELSE
		BEGIN
			--TODO: query the live system

			--we always print out at least the EXEC command
			GOTO helpbasic
		END
		RETURN 0;
	END		--IF @lv__ViewCurrent = CONVERT(BIT,0)


helpbasic:

	IF @Help <> N'N'
	BEGIN
		IF @Help NOT IN (N'params', N'columns', N'all')
		BEGIN
			--user may have typed gibberish... which is ok, give him/her all the help
			SET @Help = N'all'
		END
	END

	--If the user DID enter @start/@end info, then we use those values to replace the <datetime> tags
	-- in the @helpexec string.
	IF @start IS NOT NULL
	BEGIN
		SET @helpexec = REPLACE(@helpexec,'<start datetime>', REPLACE(CONVERT(NVARCHAR(20), @start, 102),'.','-') + ' ' + CONVERT(NVARCHAR(20), @start, 108) + '.' + 
															RIGHT(CONVERT(NVARCHAR(20),N'000') + CONVERT(NVARCHAR(20),DATEPART(MILLISECOND, @start)),3)
							);
	END 

	IF @end IS NOT NULL 
	BEGIN
		SET @helpexec = REPLACE(@helpexec,'<end datetime>', REPLACE(CONVERT(NVARCHAR(20), @end, 102),'.','-') + ' ' + CONVERT(NVARCHAR(20), @end, 108) + '.' + 
															RIGHT(CONVERT(NVARCHAR(20),N'000') + CONVERT(NVARCHAR(20),DATEPART(MILLISECOND, @end)),3)
							);
	END 


	SET @helpstr = @helpexec;
	RAISERROR(@helpstr,10,1) WITH NOWAIT;

	IF @Help=N'N'
	BEGIN
		RETURN 0;
	END

	IF @Help NOT IN (N'params',N'all')
	BEGIN
		GOTO helpcolumns
	END

helpparams:
	SET @helpstr = N'
Parameters
---------------------------------------------------------------------------------------------------------------------------------------------------------------
@start			Valid Values: NULL, any datetime value in the past

				If NULL, directs sp_SessionViewer to pull its data from the DMVs (i.e. the live data for this SQL instance). If @start is NULL, 
				@end must also be NULL. If non-null, must be a time in the past, and must be < @end. Past times define the starting point of a 
				time window for which AutoWho data captures (i.e. historical state) will be displayed. 
				NOTE: see the help text for the @offset parameter below for more info on how this works.
				NOTE: "pulling from DMVs" has not yet been implemented.
	
@end			Valid Values: NULL, any datetime in the past

				If NULL, directs sp_SessionViewer to pull its data from the DMVs (i.e. current state). If @end is NULL, @start must also be NULL. 
				(NOTE: "pulling from DMVs" has not yet been implemented.) If non-null, must be a time in the past, and must be > @start. Past times 
				define the ending point of a time window for which AutoWho data captures will be displayed. 
				NOTE: see the @offset parameter for more info on how this works.
';
	RAISERROR(@helpstr,10,1);
	SET @helpstr = N'
@offset			Valid Values: NULL, any integer from -99999 to 99999

				When @start & @end are both NULL, @offset has no function and is ignored. When @start and @end times are specified, and thus 
				AutoWho/historical data is displayed, the @offset value defines which specific AutoWho capture time is displayed. For example, 
				if @start is "2015-01-01 04:00:00.000" and @end is "2015-01-01 05:00:00.000", every AutoWho capture time (default: roughly every 
				15 seconds) is numbered in order. Thus, in the example 1 hour timeframe just specified, "2015-01-01 04:00:00.000" would be @offset=1, 
				"2015-01-01 04:00:15.000" would be @offset=2, and so on all the way to "2015-01-01 05:00:00.000", @offset=240. (Note that since
				in the real world AutoWho does not run on perfect 15 second boundaries, the actual times are likely to be something like 
				"2015-01-01 04:00:01.463", "2015-01-01 04:00:16.137", and so on.)

				The user can thus move through the time range manually, by setting @offset first to 1, and then to 2, and so on, moving "forward through 
				time". He/she can also move backwards through time by using negative offsets. Thus, by starting with -1, then -2, etc. he/she would start 
				at "2015-01-01 05:00:00.000" and move to "2015-01-01 04:59:45.000", etc.
';
	RAISERROR(@helpstr,10,1);
	SET @helpstr = N'
				Because manually changing the @offset each time would be slow and tiresome, the special values 99999 and -99999 tell sp_SessionViewer to 
				go to the next @offset value forward or backward, respectively. The default value of @offset is 99999, and the default "position" starts at 
				1. Thus, when the user enters a @start and @end time range, he/she can just hit F5 and move through time without micro-managing offsets. This 
				is possible because the @offset value passed in is stored internally in a permanent table, specific to the time range specified for the SPID of 
				the user running sp_SessionViewer.

				The value "0" is also special. It directs sp_SessionViewer to execute its own logic once for each AutoWho capture time in the @start/@end time 
				window specified, in ascending time order. Thus, the output window in Management Studio will display a separate result set for each capture. 
				This is helpful for watching a specific SPID over time. Because it can generate a large number of result sets, it is recommended that users 
				only use @offset=0 when filtering the SPIDs returned to very specific criteria.
';
	RAISERROR(@helpstr,10,1);

	SET @helpstr = N'
@activity		Valid Values: 0, 1 (default), and 2

				Specifies which types of SPIDs the user wants to see. 0 means only SPIDs actively running a query/batch should be displayed. 1 is a superset of 0, 
				and includes SPIDs that are idle (not running a query), but have a transaction open. 2 is a superset of 1, including SPIDs that are idle with no 
				open transaction. Note that even if a SPID would normally be excluded, if it is blocking an active spid that IS included, the excluded spid will 
				still be included. Thus, when 0 is specified, idle spids that are blockers are included. When 1 is specified, an idle spid with NO tran can still 
				be included if it is blocking an active spid, for example when the active spid is attempting to put a DB into single-user mode. 

				NOTE 1: Any spid with >= 64000 pages (500 MB) of currently-allocated tempdb space will automatically appear in sp_SessionViewer output regardless of 
				whether it is actively running or has an open transaction.
				
				NOTE 2: when viewing historical/AutoWho data, an inclusive filter at display time is only as good as how inclusive AutoWho was when it collected the data. 
				For the @activity filter, its effectiveness is limited to what the AutoWho "IncludeIdleWithTran" and "IncludeIdleWithoutTran" settings are configured to. 

@dur			Valid values: 0 (no duration filtering, the default), or any positive integer

				Directs sp_SessionViewer to filter out spids with a duration shorter than @dur. For a spid running a query/batch, the duration is how long the batch has 
				been running ("active duration"), defined as DATEDIFF(ms, dm_exec_requests.start_time, Capture Time). For idle spids, this is defined as how long the spid 
				has been idle ("idle duration"), defined as DATEDIFF(ms, dm_exec_sessions.last_request_end_time, Capture Time). 

				NOTE: when viewing historical/AutoWho data, the @dur filter can only be as inclusive as the AutoWho "DurationFilter" option allows.
';
	RAISERROR(@helpstr,10,1);

	SET @helpstr = N'
@dbs			Valid values: empty string (default), or a comma-separated list of DB names

				Directs sp_SessionViewer to only include spids whose DB context (dm_exec_sessions.database_id) matches the list of DB names specified. An empty string 
				means that no filtering is done by DB name.

				NOTE: when viewing historical/AutoWho data, the @dbs filter can only be as inclusive as the AutoWho "IncludeDBs" option allows.

@xdbs			Valid values: empty string (default), or a comma-separated list of DB names

				Directs sp_SessionViewer to exclude spids whose DB context matches the any of the DB names specified. An empty string means that no filtering is done.

				NOTE: when viewing historical/AutoWho data, the @xdbs filter can only be as inclusive as the AutoWho "ExcludeDBs" option allows.

@spids			Valid values: empty string (default), or a comma-separated list of SPID numbers 

				Directs sp_SessionViewer to only include certain session_id values. An empty string means no filtering is done.

@xspids			Valid values: empty string (default), or a comma-separated list of SPID numbers

				Directs sp_SessionViewer to exclude certain session_id values. An empty string means no filtering is done. 

@blockonly		Valid values: Y or N (default)

				If "Y", only spids that are blockers or blockees are displayed. 
	';

	RAISERROR(@helpstr,10,1);

SET @helpstr = N'@attr			Valid values: Y or N (default)

				If "Y", causes several additional columns to be included in the result set. 
					Client	- displays the value of dm_exec_sessions.host_name
					IP		- displays the value of dm_exec_connections.client_net_address
					Program	- displays the value of dm_exec_sessions.program_name
					SessAttr	- a click-able XML value containing the remaining attribute values from dm_exec_sessions, dm_exec_connections, and several other odds and ends

@resources		Valid values: Y or N (default)

				If "Y", causes the "Resources" column to be included in the result set. This column is a click-able XML value that contains resource-usage values from 
				dm_db_session_space_usage, dm_db_task_space_usage, dm_exec_memory_grants, and several other odds and ends.

@batch			Valid values: Y or N (default)

				If "Y", causes the "BatchText" column to be included in the result set, for user spids that are currently executing a batch. This column is an XML value 
				containing the text of the complete T-SQL batch that was submitted to SQL Server, not just the current statement. (The currently-executing statment for active
				spids is always displayed, in the "Current_Command" column.)

				NOTE: when viewing historical/AutoWho data, this data is not available unless the AutoWho option "ObtainBatchText" is set to Y. Otherwise, this column is blank.
';

	RAISERROR(@helpstr,10,1);

SET @helpstr = N'
@plan			Valid values: "none" (default), "statement", "full"

				If "none", no query plan information is displayed. If "statement" or "full", the "QueryPlan" column is included in the result set. This column contains an XML 
				value that, when selected, will open a new window with the graphical query plan. The plan is only relevant to active user requests, of course.

				Note that when viewing historical data, if the AutoWho options "ObtainQueryPlanForStatement" and "ObtainQueryPlanForBatch" are both "N", there will be no query 
				plans stored in the database and thus the "QueryPlan" column will always be empty. Also, the "QueryPlanThreshold" and "QueryPlanThresholdBlockRel" options control 
				how long a request must be running before its query plan is obtained. 

@ibuf			Valid values: Y or N (default)

				If "Y", causes the "InputBuffer" column to be included in the result set. This column is an XML value containing the results of DBCC INPUTBUFFER for the SPID. 
				For historical data, the AutoWho option "InputBufferThreshold" controls how long a spid must be running or how long it must be idle before its input buffer is collected.
';

	RAISERROR(@helpstr,10,1);

SET @helpstr = N'
@bchain			Valid values: 0 (default) through 10

				Controls how many levels deep for which the "blocking chain" graphical representation is constructed. The blocking chain is an XML value that displays at the top of 
				the sp_SessionViewer result set, and contains a hierarchical structure (via indentation) of the root blockers and their blocking chains. Since very deep blocking 
				chains can be costly to construct without much payoff (usually the first few levels are the most important), this parameter allows the user to specify how deep
				to construct the chains. The parameter allows values of up to 10-deep, but of course the actual depth of the blocking chain could be much shallower than 10 levels. 
				
				NOTE: When viewing historical data, the AutoWho option "BlockingChainDepth" controls how many levels of blockees are CAPTURED. Thus, if the AutoWho option is 4, 
				then the max depth that can ever be displayed is 4, even if very deep blocking chains occurred in the SQL workload. The AutoWho logic to capture the blocking chain 
				data is only triggered once there has been a spid blocked for >= the "BlockingChainThreshold" option. Obviously, the sp_SessionViewer @bchain param can only be as 
				inclusive as the AutoWho option.

				For more information, see the help section "Special Rows" further down below, displayed when @help="columns" or @help="all".
';

	RAISERROR(@helpstr,10,1);

SET @helpstr = N'
@tran			Valid values: Y or N (default)

				If "Y", causes the "Transactions" column to be included in the result set. This column is an XML value that contains details on all of the transactions that a 
				spid has open. When viewing historical data, this information is collected by AutoWho for every spid that has dm_exec_sessions.open_transaction_count > 0 or 
				whose active or idle duration is >= the "TranDetailsThreshold" option in AutoWho.

@waits			Valid values: 0, 1, 2, or 3

				Values 0 through 2 control how much wait-type detail is displayed in the "Progress" column. (See the notes under the "Progress" column for more info). Value 3 
				controls whether the Lock Details XML value is displayed in the "Current_Command" column (see the help section "Special Rows" when @help="columns" or @help="all"). 

@savespace		Valid values: Y or N (default)

				When "Y", the column names are abbreviated to a shorter form, in order to condense the columns and reduce or eliminate the amount of horizontal scrolling needed. 
				A few other space-saving optimizations are done.

@directives		Valid values: "query/ies"

				Think of "directives" as a bit like SQL trace flags. They are not necessarily documented fully, and change the  behavior of sp_SessionViewer in various ways that can 
				change easily from release to release. They allow the developer of sp_SessionViewer to add in fringe features without committing to a change in the API. Over time,
				more valuable fringe features may be promoted via their own parameter. As of this release, there is only one directive (besides debugging-related directives): 
				adding the text "query" or "queries" to the @directives parameter when querying historical/AutoWho data will cause the Current_Command field in the very top row to 
				contain a number of queries against the underlying AutoWho tables. This allows follow-up ad-hoc analysis.

				For more information, see the section "Special Rows" when @help="columns" or @help="all".
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
Special Rows
---------------------------------------------------------------------------------------------------------------------------------------------------------------
Conceptually, each output row from sp_SessionViewer represents a session (aka spid). (To be completely accurate, each row represents a "request" that a spid is running, or the spid 
itself if it is idle. However, a spid will only have multiple rows if the MARS feature is in use, so most of the time it is sufficient to think of each row as a spid.

However, some information returned by sp_SessionViewer is not directly tied to one and only one spid. Therefore, this information is displayed on "special spids" 
(internally represented with negative session_id values). These special rows usually provide click-able XML values that can bring special insight in certain scenarios.

It may be helpful to think of sp_SessionViewer output as a "dashboard", with multiple sections. The top section contains "special rows" with special, non-spid-specific information. 
The next section contains system spids, if any qualify. The third section contains "active spids/requests", and the final section contains idle spids. The resulting rows are always
ordered by these sections; thus, idle SPIDs will always sort lower than active user SPIDs, which will always sort lower than system SPIDs, which will always sort lower than the "special rows".

The secondary "order by" column is effectively "by duration descending". Thus, an active spid that has been running for 5 minutes will be nearer to the top than an active spid that has 
been running for 7.2 seconds. An spid that has been idle for 13.6 seconds will be nearer to the top than a spid that has been idle for 15 minutes (because idle durations can be thought of
as "negative durations"). Future versions may reverse the idle sort order so that long-idle spids sort at the top of the idle section. Future versions may also give the user control over 
the sort order of the records (within, but not across, each of the four sections).
';

	RAISERROR(@helpstr,10,1);

	SET @helpstr = N'
In this build of sp_SessionViewer/AutoWho, the following "special rows" are possible: 

	CaptureTime --> This row always appears for historical runs of sp_SessionViewer (when @start & @end have values and data is pulled from the AutoWho tables), and never for 
					"current" runs, where data is pulled directly from the DMVs. It contains values for 3 columns:
						"SPID" --> The text "CapTime" is displayed, along with the current numerical offset: e.g. "CapTime[3]". See the notes on the @offset parameter.
						"SPIDContext" --> displays the exact capture time, down to the millisecond, when AutoWho captured the data. This allows easy copy-pasting into ad-hoc 
											queries against the underlying AutoWho tables.
						"CurrentCommand" --> is empty, unless certain "extra" data is available. For example, the value: "<?opt -- bchain trans waits3 --?>" indicates that Blocking Chain 
											data, Transaction Info data, and Lock Details data is available for this SPID Capture Time. This serves as an indicator to the user that they 
											can utilize the "bchain", "tran", and "waits=3" parameters to view this extra info.
	';

	RAISERROR(@helpstr,10,1);

	SET @helpstr = N'
	Blocking chains --> Displays an XML value (in the "CurrentCommand" column) that contains a textual, hierarchical representation of blocking chains. Each root blocker appears left-
						justified, and each of its blockees appears indented underneath. Data from dm_os_waiting_tasks is included to give more information about the locks involved.

	Lock Details	--> Displays a textual representation of the contents of dm_tran_locks, for relevant blocking/blockee spids. For historical data, relevant blocking spids are those
						who are blocked >= "ObtainLocksForBlockRelevantThreshold" (AutoWho option) milliseconds, and the spids that are blocking them. For current data, any spid that is 
						blocked or blocking is a relevant blocking spid. Note that in order to avoid displaying many rows of Row, Page, Key, or other detail-level locks, the dm_tran_locks 
						records are aggregated via COUNT(*) over all of the displayed columns, and the count is reported in the "#Rows" column.

						Displaying this value can be time-consuming; when viewing historical/AutoWho data (when the DMV capture logic has already happened previously and only display logic
						needs to occur), this is usually the cause of sp_SessionViewer executions that are longer than 1 or 2 seconds.

	Tasks W/O SPIDs	--> On very busy servers, the SQL Server thread pool can be exhausted and incoming requests can wait with the reason "THREADPOOL". These incoming requests do not yet 
						have their own SPID, so sp_SessionViewer aggregates them into a special row. The "#Tasks" column shows how many tasks are waiting on the THREADPOOL wait type.
	';

	RAISERROR(@helpstr,10,1);

	SET @helpstr = N'
Columns
---------------------------------------------------------------------------------------------------------------------------------------------------------------
SPID			A composite field that usually just displays the session_id value from sys.dm_exec_sessions. For "special rows", either "Cap Time", "Blk Chains", "Lock Details", 
				or "Tasks W/O spids" is displayed. For the remaining rows, the composite field is constructed as such:
					<blocking indicator> + <system spid indicator> + <session id> + <request ID>

				The session ID is always displayed. The other 3 indicators are only present if appropriate. Here are some examples: 
					72		--> a normal user session with the session_id value of "72". The spid may be idle or may be actively running a batch. Use the "Duration" column 
								to determine which.
					s14		--> the session_id is 14, and it is a system spid. System spids are only displayed when their dm_exec_requests.wait_type status is DIFFERENT than 
								their "typical" state. For example, the CHECKPOINT spid''s normal wait type is "CHECKPOINT_QUEUE", as it waits to be woken up for explicit or 
								automatic checkpoints. The sp_SessionViewer logic will only include the CHECKPOINT spid (dm_exec_requests.command="CHECKPOINT") when its wait 
								type is NOT "CHECKPOINT_QUEUE". Most other system spids are the same way, such that the user can be confident that the presence of a system spid 
								in the output is an indicator that the system spid in question is doing real work rather than its normal state of sleeping. ';

	RAISERROR(@helpstr,10,1);

	SET @helpstr = N'
					*  122		--> SPID 122 is not only blocking one or more other spids, but it is also blocked.
					(!)  103	--> SPID 103 is blocking one or more other spids, but it is not blocked by any spids. This means that it is a root blocker. The exclamation mark 
									can help the user quickly identify root blockers when much blocking is occurring. 
					(!) s14		--> The system spid is a root blocker. This is rare, and in the author''s experience usually happens when the CHECKPOINT system spid and a checkpoint 
									running from a backup are contending over the same database. 
					* 271:2		--> SPID 271 is a blocker (but not a root one) and is running a request with a non-zero request_id. Most requests run under sys.dm_exec_requests.request_id = 0, 
									but when MARS is used or at least enabled, request_ids can be > 0.';

	RAISERROR(@helpstr,10,1);

	SET @helpstr = N'
SPIDContext		For the "CapTime" special row, displays the datetime value of when the AutoWho collector captured DMV data. For other special rows, is empty. For true session rows, 
				displays the database context that the SPID is executing in. This is DB_NAME(sys.dm_exec_sessions.database_id). 

				Several caveats should be mentioned: First, some system SPIDs do not have a database context (database_id = 0), or have a DB context equal to the resource 
				database (32767). Empty strings are displayed for these. On occasion, a DB context cannot be established (for various reasons) for the spid, and empty strings are used 
				here as well. Finally, the SPIDContext is the result of running DB_NAME() on the session database_id at the time that sp_SessionViewer is called. Thus, if AutoWho collects 
				DMV data on Monday, and a database is detached and re-attached on Tuesday and receives a different database ID, and sp_SessionViewer is called on Wednesday for Monday''s data, 
				the results of DB_NAME(old DB ID) will often make little sense to the user. The design choice to store the DB ID instead of the name is due to the expense of calling 
				DB_NAME on systems that may have hundreds or thousands of SPIDs. This design choice may be re-evaluated in a later release.';

	RAISERROR(@helpstr,10,1);

	SET @helpstr = N'
Duration		Presents the active duration [DATEDIFF(millisecond, dm_exec_requests.start_time, AutoWho capture time)] or the idle duration 
				[DATEDIFF(millisecond, dm_exec_sessions.last_request_end_time, AutoWho capture time)] of a system or user spid. The format depends on the length of the duration:

					< 1 minute: <seconds>.<tenths of a second> --> 13.7
					< 1 day: HH:MM:SS --> 03:07:33
					>= 1 day: Day~HH:MM:SS --> 16~07:04:01	(this is most common with system spids and in those cases often approximates SQL instance uptime)

				In order to easily distinguish active from idle spids, the duration for idle spids is prepended with a "minus" sign. Thus, "-24.3", "-01:05:33", etc. In rare occasions, 
				the dm_exec_requests.start_time and/or dm_exec_sessions.last_request_end_time values can be "1900-01-01", and in these cases this field displays "???". Those scenarios 
				are usually very transient. 

				As mentioned above in the "Special Rows" section, the Duration column can be thought of as the "second" column in the ORDER BY clause for sp_SessionViewer output.
	';

	RAISERROR(@helpstr,10,1);

	SET @helpstr = N'
Blocker			If a SPID is blocked, the session_id of the blocking SPID is displayed, otherwise the field is blank. The blocking SPID # is not derived from dm_exec_requests, 
				which can be wrong for parallel queries. Instead, the correct SPID # is calculated from the dm_os_waiting_tasks data: the blocked task (blocking_session_id 
				IS NOT NULL and not = to the session id [e.g. a CXPACKET wait]) with the longest wait_duration_ms value is the task that defines the blocking session ID for the 
				whole SPID. Note that blocking can be due not only to lock waits (wait_type = LCK_M_*), but also to certain types of PAGELATCH or PAGEIOLATCH waits, and 
				even RESOURCE_SEMAPHORE waits. 
	';

	RAISERROR(@helpstr,10,1);

	SET @helpstr = N'
ObjectName		For system SPIDs, displays the value of dm_exec_requests.command, indicating what area of the SQL engine''s code the system spid is executing. For user spids that 
				are idle or executing ad-hoc T-SQL, an empty string is displayed. For user spids that are executing in a procedure, function, or trigger, the OBJECT_NAME(object_id) 
				value is displayed, along with schema information. (DB information is also included if the DB name of the object is different than the DB context presented in the 
				"SPIDContext" field). 

				The object_id value that is used to determine the proc/function/trigger name is obtained by calling the dm_exec_sql_text DMF using the dm_exec_requests.sql_handle value. 
				If the object_id value is NULL, the SPID is considered to be executing ad-hoc T-SQL, and the object_id value is coalesced internally into a special negative integer value. 
				When the sql_handle value is 0x0, but the spid is clearly active, this field displays the "?obj N/A?" value. 

				One exceptional case is encountered often enough to deserve special handling. Active spids sometimes have a dm_exec_requests.command value of "TM REQUEST". Often, a 
				tran-related wait type (such as "WRITELOG" or "DTC") is found in the dm_exec_requests wait_type field, and the dm_exec_requests.sql_handle value is 0x0. When
				the command is "TM REQUEST", therefore, the value "(TMRQ)" is appended to the object name (if any is present). Also, the wait type from dm_exec_requests is used in the 
				"Progress" field (see below) instead of using the one from dm_os_waiting_tasks. If this behavior occurs frequently enough, users may want to evaluate the Disk IO latency 
				for their T-log files, or the effects of distributed transactions on throughput performance. 
	';

	RAISERROR(@helpstr,10,1);

	SET @helpstr = N'
CurrentCommand	For special rows, displays a click-able XML value that provides important contextual info, such as blocking chains or lock details. For the CapTime special row, 
				displays (if applicable) which extra options are available to the user for the current capture time. For idle spids, this value is always an empty string.

				For active user SPID rows, displays the current statement that is executed, obtained by calling the dm_exec_sql_text DMF using the dm_exec_requests sql_handle, and 
				applying the statement_start_offset and statement_end_offset values. For historical data, at the bottom of the XML value is printed the surrogate key to the table that 
				holds a unique copy of all SQL text captured by AutoWho. This allows in depth research and debugging, as well as obtaining the sql_handle and offset values. For current
				data, the sql_handle, plan_handle, and offset values are appended to the bottom of the statement.

				Note that for active spids, whose dm_exec_requests.command value is "TM REQUEST", it is common for the sql_handle value to be 0x0. In these cases, the "CurrentCommand" 
				field is blank even for active spids.
	';

	RAISERROR(@helpstr,10,1);

SET @helpstr = N'
#Tasks			The number of tasks allocated to a specific running query. The special row "Tasks W/O SPIDs" holds a COUNT(*) from the dm_os_waiting_tasks DMV for tasks with a NULL session_id and 
				which are in the THREADPOOL wait. System spids and active spids running a serial query plan will always have 1 in this column. Idle spids will always have an empty string. 
				Active spids running a parallel query plan will have a fluctuating number of tasks.
				For active user SPIDs, this # is a COUNT(*) (grouped by session_id/request_id of course) of sys.dm_db_task_space_usage.

				The #Tasks column always reflects the # of tasks that were observed for a given active spid at the time that spid was observed. However, for historical data, that value can 
				differ from the number of task rows saved to the AutoWho.TasksAndWaits table for parallel queries for 2 reasons: 1) Unless a parallel query''s duration exceeds the 
				"ParallelWaitsThreshold" AutoWho option, only its top task/wait is saved to AutoWho.TasksAndWaits. This saves on disk space/processing time, since shorter-lived parallel queries 
				may not be of interest. 2) Even if a parallel query''s duration exceeds the "ParallelWaitsThreshold" option and all of its records from dm_os_tasks/dm_os_waiting_tasks are saved 
				to AutoWho.TasksAndWaits, the fact that 2 different DMVs were accessed at 2 different times means that subtle timing issues (between when dm_db_task_space_usage is
				accessed and when dm_os_tasks/dm_os_waiting_tasks are accessed) may come into play.
';


	RAISERROR(@helpstr,10,1);

SET @helpstr = N'
Progress		This complex field is composed of several conditional parts: Pct complete + session status + rqst status + wait info. Each section is described below.

					Pct Complete --> Exposes the dm_exec_requests.percent_complete, but only if the value is >= 0.001. Otherwise, an empty string.

					Session status --> The idea here is to show the dm_exec_sessions.status field if it is "interesting". When the session status is "interesting", it is displayed in 
									square brackets. For example, if an idle session has a status of "Sleeping", that is expected and thus not very interesting. Ditto for an active 
									session with a status of "Running". System spids are "not interesting" if they are either "Sleeping" or "Running". All other scenarios are considered 
									"interesting", such as "Preconnect" or "Dormant" statuses, and active spids in a "Sleeping" status or idle spids in a "Running" status. (Note that 
									these last 2 cases are usually due to timing-related issues. Even though the dm_exec_requests and dm_exec_sessions data is pulled in the same query, 
									the query plan gathers data from dm_exec_requests before it gathers the corresponding data from the sessions view. Thus, there could be micro- or 
									millisecond differences in the gather times, and therefore apparent inconsistencies. These inconsistencies are not common, and can usually be ignored.) 
';

	RAISERROR(@helpstr,10,1);

SET @helpstr = N'
					Request status --> This is very similar to the session status logic described above. However, the status is displayed in curly brackets if interesting, and comes 
									from dm_exec_requests.status. Nothing is ever displayed for idle spids. Active spids in "Running" or "Suspended" status have an "uninteresting" status, 
									as do system spids that are in "Background" status. This leaves "Runnable" and "Sleeping" for all spids, and "Background" for user spids. The author
									has not seen "Sleeping" or user-"Background" states, though "Runnable" is somewhat common and, when seen for many spids at the same time can be an 
									indicator of CPU contention.
';

	RAISERROR(@helpstr,10,1);

SET @helpstr = N'
					Wait info --> The "wait info" portion is more complex, but before we can adequately describe its sub-parts, we need to define the sources of "wait info" since they can 
								vary under different circumstances. "Wait Info" can be thought of as being derived from three different DMVs, (dm_exec_requests, dm_os_tasks, and 
								dm_os_waiting_tasks), depending on the status of the spid. Idle spids do not have data in any of these 3 DMVs, and therefore the "wait info" portion 
								is completely blank. Active spids executing a serial query plan have only 1 task, and thus only 1 row in dm_os_tasks. That single task may or may not be 
								waiting, and thus there may or may not be a corresponding row in dm_os_waiting_tasks. (Note that there even could be MULTIPLE rows in dm_os_waiting_tasks for 1 
								dm_os_tasks task_address value, in the event where SQL Server tracks multiple blockers for a given task.)

								An active spid executing a parallel query plan will usually have multiple sub-tasks, and thus multiple rows in dm_os_tasks. These tasks will often have multiple 
								records per task_address in dm_os_waiting_tasks. Thus, the question of "which wait information to display" becomes complex quite quickly. Because of limited space 
								in the sp_SessionViewer result set, the concept of a "most relevant task" is used and the sp_SessionViewer output uses that "top task" as the basis of the
								"wait info" portion. For active, serial queries, there is only 1 "top task" and it is therefore the only one captured (and in the case of AutoWho, stored). If 
								that serial plan task has multiple rows in dm_os_waiting_tasks (e.g. multiple blocker spids are being tracked), the one with the longest wait duration is used 
								to define what the wait_type and wait_duration_ms are. For parallel queries, the "top task" is defined using a scheme that prioritizes certain waits or
								statuses higher/more important than other waits or statuses. (In the case of AutoWho, "non-top tasks" and waits are only stored if the duration of the active batch 
								is >= the "ParallelWaitsThreshold" AutoWho option). 
';
	RAISERROR(@helpstr,10,1);

SET @helpstr = N'				Here is the priority order for defining the "top task": 
									1. LCK_M_* waits (waiting for a lock)
									2. Latch waits with a non-null blocking_session_id (that is <> the session_id column)
									3. PAGE/PAGEIO latch waits with a null blocking_session_id
									4. "Catch-all" bucket for anything not in categories 1 through 3, or 5 or 6.
									5. CXPACKET waits 
									6. A running task (i.e. not waiting)

								Thus, even if a parallel query has CXPACKET waits of 10 seconds and PAGEIO latch waits of 1 second, and a lock wait of 3 milliseconds, the lock wait will get 
								top priority, and that task/wait info will be shown. An active spid executing a parallel query will only have a blank "wait info" section if ALL tasks are running 
								(this is not common), and it will only show a CXPACKET wait if all tasks are either in CXPACKET waits or running. This reflects the fact that CXPACKET waits do 
								not, by themselves, indicate that a query is truly blocked, stuck, or otherwise unable to progress. 
';

	RAISERROR(@helpstr,10,1);

SET @helpstr = N'				The above logic based on finding the "top task" in dm_os_tasks and dm_os_waiting_tasks does have two important exceptions: For system spids, or a user spid with 
								a dm_exec_requests.command="TM REQUEST", then the wait type is taken from dm_exec_requests. This is because both types of activities usually involve shorter waits, 
								and it is possible, even likely, that the short time gap in between when dm_exec_requests and dm_os_waiting_tasks are accessed (separate queries) will produce 
								confusing and inconsistent data. Of course, when the dm_exec_requests.wait_type column is used, the dm_exec_requests.wait_duration column is used as well. 
								"TM REQUEST" spids often have a WRITELOG or DTC-related wait associated with them. 

								The "wait info" portion follows this format: "wait tag / duration indicator". Consider these 3 examples: 
									WAITFOR  = 8719ms
									WAITFOR  --> 14.6sec <--
									BACKUPTHREAD  = 7813ms
								In all three cases, the "wait tag" is simply just the wait_type value. In 2 of the 3 cases, the duration is less than 10000 milliseconds, and thus an "=" character 
								is used. For the wait that is >= 10000 milliseconds, "arrows" have been used to draw attention to the longer wait. 
';

	RAISERROR(@helpstr,10,1);

SET @helpstr = N'
								When the @waits parameter is 0, the "wait tag" is just wait_type. However, @waits=1 and @waits=2 add more information for some common
								wait types. Consider: 
									CXPACKET  = 103ms					(@waits=0)
									CXPACKET:PortOpen:3  = 103ms		(@waits=1 or 2)
								The increased detail gives us the "CXPACKET sub-type" and the node ID (3) that it is occurring on. Another example:
									PAGEIOLATCH_SH  = 0ms										(@waits=0)
									PAGEIOLATCH_SH{myDB1 1:*}  = 0ms							(@waits=1)
									PAGEIOLATCH_SH{myDB1 ObjId:955150448, IxId:1}  = 906ms		(@waits=2)
								The middle row gives us the DB name and the file ID # for the page that is being waited on. If the page was a special system bitmap, an acronym like 
								"GAM" or "PFS" would appear instead of the asterisk. The last row gives us the actual object id and index id that the page belongs to. This output is 
								only possible if the AutoWho option "ResolvePageLatches" is set to "Y". This option is off by default due to its extra overhead. 

								Finally, consider an example involving locking. 
									LCK_M_S  = 2401ms											(@waits=0)
									OBJECT{req:S held:IX }  = 2401ms							(@waits=1)
									OBJECT{req:S held:IX }{dbid:10 id:1109578991}  = 2401ms		(@waits=2)
								Both @waits=1 and @waits=2 tell us what type of lock is requested (an "OBJECT" lock, e.g. a table), that it is requested
								in "S" mode (an S-lock on a table object is a table-level lock), and that it is held in "IX" mode, which is incompatible with
								table-level S-locks. When @waits=2, we can even tell which DB and object it is. 
';

	RAISERROR(@helpstr,10,1);

SET @helpstr = N'
QueryMem_MB		The amount of memory, in MB, requested or granted to the active request. If the memory has been granted, the value comes from dm_exec_memory_grants.granted_memory_kb. 
				If the memory has NOT yet been granted, the text "(Req)  " is prepended and the value comes from dm_exec_memory_grants.requested_memory_kb. If the value is 0, an empty 
				field is displayed.

				More granular information about the SPID''s query memory usage is available in the "Resources" field.

TempDB_MB		The amount of TempDB space, in MB, that the SPID is currently using. This is calculated by taking each "pair" in dm_db_session_space_usage and dm_db_task_space_usage and 
				subtracting the dealloc field from the alloc field. If the result is < 0, it is floored at 0. The results from each pair are then added together to find the total
				current usage.

				The "lifetime allocation" is also calculated (i.e. just adding the 4 "alloc" fields), and if the value is >= 50 MB, the text " /a: " is appended and then followed by the 
				total allocation. This can help identify SPIDs that have allocated much tempdb and then quickly deallocated it. Such allocation "spikes" can identify short-but-intensive batches.

				If an empty field is displayed, the current usage is 0 and the "lifetime allocation" is < 50 MB (and may be 0 also). More granular information about the SPID''s TempDB 
				usage is available in the "Resources" field.

CPUTime			For active SPIDs, the value from dm_exec_requests.cpu. For idle spids, the value from dm_exec_sessions.cpu. If the value is 0, an empty field is displayed.

PhysicalReads_MB	For active SPIDs, the value from dm_exec_requests.reads. For idle spids, the value from dm_exec_sessions.reads. If the value is 0, an empty field is displayed.

LogicalReads_MB		For active SPIDs, the value from dm_exec_requests.logical_reads. For idle spids, the value from dm_exec_sessions.logical_reads. If the value is 0, an empty field is displayed.

Writes_MB		For active SPIDs, the value from dm_exec_requests.writes. For idle spids, the value from dm_exec_sessions.writes. If the value is 0, an empty field is displayed.
';

	RAISERROR(@helpstr,10,1);

SET @helpstr = N'
Transactions	Only present if @tran is set to Y and transaction information is available for one or more SPIDs in the result set. For those SPIDs whose transactions were captured, 
				a click-able XML value is available and contains a list of the various transactions that the SPID has opened (explicitly or implicitly). The XML value has a begin tag that
				indicates how many KB or MB have been written to the transaction log(s) by the transaction(s).

Resources		When @resources is set to "Y", this column displays an XML value for each user SPID. This XML contains detailed info on TempDB allocations/deallocations, 
				memory request and grant information, CPU and IO info, and a few other odds and ends.

BatchText		If AutoWho has been configured to capture the full Batch text for running SPIDs (via the "ObtainBatchText" option), and the @batch parameter has been set to "Y", then 
				sp_SessionViewer returns a column named "BatchText" that holds the complete text of the currently-running batch. (The "CurrentCommand" statement only ever holds the current
				T-SQL statement, not the complete batch). 
';

	RAISERROR(@helpstr,10,1);

SET @helpstr = N'
QueryPlan		Only present if @plan has been set to "statement" or "full". If @plan is set to "statement", and statement-level query plans are being captured (the "ObtainQueryPlanForStatement" 
				option, on by default) and the active query has run long enough to surpass the "QueryPlanThreshold" and/or "QueryPlanThresholdBlockRel" options (whichever is relevant), 
				the query plan for the currently-running statement is returned in XML form. 

				If @plan is set to full, and batch-level query plans are being captured (the "ObtainQueryPlanForBatch" option, OFF by default), and the active query has surpassed the 
				appropriate query plan threshold value (see above), then the query plan for the entire batch is returned in XML form.

InputBuffer		Only present if the @ibuf parameter has been set to "Y". For spids that have an active or idle duration longer than the "InputBufferThreshold" AutoWho option, 
				the Input Buffer captured by DBCC INPUTBUFFER is returned in XML form.
';

	RAISERROR(@helpstr,10,1);

SET @helpstr = N'
Login			Returns the value from dm_exec_sessions.login_name. If the login_name and original_login_name values differ, the original_login_name value is appended in parentheses.

Client			Only present if the @attr parameter is set to "Y". Returns the value in dm_exec_sessions.host_name

IP				Only present if the @attr parameter is set to "Y". Returns the value in dm_exec_connections.client_net_address

Program			Only present if the @attr parameter is set to "Y". Returns the value in dm_exec_sessions.program

Attributes		Only present if the @attr parameter is set to "Y". Returns a click-able XML value that contains all attributes from dm_exec_sessions and dm_exec_connections that are not 
				displayed elsewhere. Some dm_exec_request info is also displayed (in apposition to the corresponding columns in dm_exec_sessions).
';

	RAISERROR(@helpstr,10,1);

exitloc:

	RETURN 0;


END

GO
