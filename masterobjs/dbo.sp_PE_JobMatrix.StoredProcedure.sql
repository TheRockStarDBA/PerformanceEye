USE [master]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[sp_PE_JobMatrix] 
/*   
	PROCEDURE:		sp_PE_JobMatrix

	AUTHOR:			Aaron Morelli
					email@TBD.com
					@sqlcrossjoin
					sqlcrossjoin.wordpress.com
					https://github.com/amorelli005/PerformanceEye

	PURPOSE: Uses msdb job history to construct a history matrix showing durations and results, to present
	a quick graphical view of job execution for troubleshooting sessions


	FUTURE ENHANCEMENTS: 
		HERE'S WHERE I LEFT OFF: 
			- About to try to add the JobAttributes header to the right of the "LineHeader" header row (but only the one above the first matrix???)
					(blank text above the job names area, and then column headers for "CrModDate" (either create or last modified date), 
						"Notifies", "Owner", and "StepTypes")
					We only want this info if the user asks for it, except for CrModDate, which if @JobAttrib is 1 it could print Create or Mod dates
					if they were within the last 3 days of the end-time of the matrix

		Predictive matrix? (i.e. answer a question like "what is going to run tonight?")
			maybe start with "next_scheduled_run_date" in msdb.dbo.sysjobactivity?

	
		UPDATE: use queries instead. Rather than returning the below data, let a parameter enable the return of clickable XML where various
			helpful MSDB queries are printed, available for copy-paste
			Add some switch or param to cause the proc to return job duration & completion type info broken down by important dimensions
				Overall
				By day of week
				By hour of day?

		Perhaps add a NVARCHAR(MAX) parameter that receives a comma-delimited list of job names and only returns those in the output

		need to re-think what the @Debug variable actually should do

		search for the "TODO" string and consider each one for inclusion in final code

		Be consistent in using "Matrix time window" and "cell time window"; review @Help doc and Blog post

		create a script version 


    CHANGE LOG:	
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
	exec sp_JobsAtAGlance @Help=N'Y'

	exec sp_JobsAtAGlance	@PointInTime = NULL,			@HoursBack = 20,			@HoursForward = -1,
							@ToConsole = N'N',				@FitOnScreen = N'Y',
							@DisplayConfigOptions = 1,		@DisplayAgentLog = 1, 
							@Queries = 0,					@Help='N',					@Debug=0
*/
(
	@PointInTime					DATETIME		= NULL,
	@HoursBack						TINYINT			= 20,			--0 to 48
	@HoursForward					TINYINT			= 0,			--0 to 20
	@ToConsole						NCHAR(1)		= N'N',
	@FitOnScreen					NCHAR(1)		= N'Y',
	@DisplayConfigOptions			TINYINT			= 1,			-- 0 = No; 1 = Only different from default; 2 = All (important) config opts
	@DisplayAgentLog				INT				= 1,			-- 0 = No; 1 = only when Sev 1 records exist; 2 Always display the first file; 3 Display All Files
	@Queries						NVARCHAR(10)	= N'N',			-- saving room for tags (i.e. if # of queries grows, categorize them and let user focus)
	@Help							NCHAR(1) = N'N',
	@Debug							TINYINT	 = 0
)
AS
BEGIN
	SET NOCOUNT ON;
	SET DATEFIRST 7;		--Needed for the predictive matrix when handling schedules

	/*********************************************************************************************************************************************
	*********************************************************************************************************************************************

														Part 0: Variables, Validation, Temp Table Definitions

	*********************************************************************************************************************************************
	*********************************************************************************************************************************************/
	--***Location 0: Variable declarations
	
	--General variables
	DECLARE 
		@lv__HelpText						VARCHAR(MAX),
		@lv__ErrorText						VARCHAR(MAX),
		@lv__ErrorSeverity					INT,
		@lv__ErrorState						INT,
		@lv__OutputVar						NVARCHAR(MAX),
		@lv__OutputLength					INT,
		@lv__CurrentPrintLocation			INT,
		@lv__beforedt						DATETIME,
		@lv__afterdt						DATETIME,
		@lv__slownessthreshold				SMALLINT
		;

	SET @lv__beforedt = GETDATE();
	SET @lv__slownessthreshold = 250;		--number of milliseconds over which a ***dbg message will be printed; aids
											--immediate review of where in the proc a given execution was slow

	--Matrix variables		
	--UPDATE 2015-04-03: with the beginning of development on the predictive matrix, the @lv__mtx__ variables only refer to the historical matrix
	DECLARE 
		@lv__mtx__SQLAgentStartTime			DATETIME, 
		@lv__mtx__SQLServerStartTime		DATETIME,
		@lv__mtx__SQLStartTimeResult		SMALLINT,		--0 NULL, 1 success, 2 older than @lv__mtx__OverallWindowEndTime
		@lv__mtx__SQLAgentTimeResult		SMALLINT,		--0 NULL, 1 success, 2 @lv__mtx__OverallWindowEndTime is still older than the oldest Agent log looked at
		@lv__mtx__WindowLength_minutes		SMALLINT,
		@lv__mtx__MatrixWidth				SMALLINT,
		@lv__mtx__LineHeaderMod				INT, 
		@lv__mtx__TimeHeaderMod				INT,
		@lv__mtx__EmptyChar					NCHAR(1),
		@lv__mtx__CurrentTime_WindowBegin	DATETIME,
		@lv__mtx__CurrentTime_WindowEnd		DATETIME,
		@lv__mtx__OverallWindowBeginTime	DATETIME, 
		@lv__mtx__OverallWindowEndTime		DATETIME,
		@lv__mtx__MaxJobNameLength			SMALLINT,
		@lv__mtx__HeaderLine				NVARCHAR(4000), 
		@lv__mtx__HeaderHours				NVARCHAR(4000),
		@lv__mtx__Replicate1				SMALLINT, 
		@lv__mtx__Replicate2				SMALLINT,
		@lv__mtx__CountMatrixRows_1			INT, 
		@lv__mtx__CountMatrixRows_3			INT,
		@lv__mtx__CountMatrixRows_5			INT,
		@lv__mtx__Matrix3HasHeader			BIT,
		@lv__mtx__PrintOverallBeginTime		VARCHAR(30), 
		@lv__mtx__PrintOverallEndTime		VARCHAR(30)
		;

	SET @lv__mtx__SQLStartTimeResult = 0;
	SET @lv__mtx__SQLAgentTimeResult = 0;

	--Predictive matrix variables
	DECLARE 
		@lv__pred__WindowLength_minutes		SMALLINT,
		@lv__pred__MatrixWidth				SMALLINT,
		@lv__pred__LineHeaderMod			INT, 
		@lv__pred__TimeHeaderMod			INT,
		@lv__pred__EmptyChar				NCHAR(1),
		@lv__pred__CurrentTime_WindowBegin	DATETIME,
		@lv__pred__CurrentTime_WindowEnd	DATETIME,
		@lv__pred__OverallWindowBeginTime	DATETIME, 
		@lv__pred__OverallWindowEndTime		DATETIME,
		@lv__pred__MaxJobNameLength			SMALLINT,
		@lv__pred__HeaderLine				NVARCHAR(4000), 
		@lv__pred__HeaderHours				NVARCHAR(4000),
		@lv__pred__Replicate1				SMALLINT, 
		@lv__pred__Replicate2				SMALLINT,
		@lv__pred__CountMatrixRows_1		INT, 
		--@lv__mtx__CountMatrixRows_3		INT,		--no need to separate matrix 1 and 3 for predictive, since we always assume jobs succeed
		@lv__pred__CountMatrixRows_5		INT,
		@lv__pred__Matrix5HasHeader			BIT,
		@lv__pred__PrintOverallBeginTime	VARCHAR(30), 
		@lv__pred__PrintOverallEndTime		VARCHAR(30)
		;

	--Config Option variables
	DECLARE 
		@lv__cfg__MaxHistoryRows			INT,
		@lv__cfg__MaxHistoryRowsPerJob		INT,
		@lv__cfg__tmpregstr					NVARCHAR(200),
		@lv__cfg__ServiceStartupSetting		INT,
		@lv__cfg__ShouldAgentRestartSQL		INT,
		@lv__cfg__errorlog_file				NVARCHAR(255),
		@lv__cfg__errorlogging_level		INT,			-- 1 = error, 2 = warning, 4 = information
		@lv__cfg__error_recipient			NVARCHAR(30),	-- Network address of error popup recipient
		@lv__cfg__monitor_autostart			INT,
		@lv__cfg__local_host_server			SYSNAME,		-- Alias of local host server
		@lv__cfg__job_shutdown_timeout		INT,
		@lv__cfg__login_timeout				INT,
		@lv__cfg__idle_cpu_percent			INT,
		@lv__cfg__idle_cpu_duration			INT,
		@lv__cfg__oem_errorlog				INT,
		@lv__cfg__alert_replace_runtime_tokens INT,
		@lv__cfg__cpu_poller_enabled		INT,
		@lv__cfg__use_databasemail			INT,
		@lv__cfg__databasemail_profile		SYSNAME
	;

	--SQL Agent Log variables
	DECLARE 
		@lv__log__maxTabID					INT,
		@lv__log__log1processing			SMALLINT,		--not started; 1 load completed; -1 load failed; 2 processing completed; -2 processing failed; -3 cancelled due to previous failure
		@lv__log__log2processing			SMALLINT,
		@lv__log__log3processing			SMALLINT,
		@lv__log__AgentLogString			VARCHAR(MAX);

	SET @lv__log__log1processing = 0;
	SET @lv__log__log2processing = 0;
	SET @lv__log__log3processing = 0;

	--Final Display control bits
	DECLARE 
		@output__DisplayMatrix				BIT,
		@output__DisplayPredictive			BIT,
		@output__DisplayAgentLog			BIT,
		@output__DisplayConfig				BIT,
		@output__DisplayQueries				BIT,
		@outputType__Matrix					NVARCHAR(10),
		@outputType__Predictive				NVARCHAR(10);

	SET @output__DisplayAgentLog = CASE WHEN @DisplayAgentLog = 0 THEN 0 ELSE 1 END; 
	SET @output__DisplayConfig = CASE WHEN @DisplayConfigOptions = 0 THEN 0 ELSE 1 END;
	SET @output__DisplayQueries = CASE WHEN LOWER(@Queries) = N'n' THEN 0 ELSE 1 END;		--if user types anything but 'N', we assume they want to see the queries

	--The user may want to see either, both, or none of the matrices
	IF @HoursBack > 0 
	BEGIN
		SET @output__DisplayMatrix = 1;
	END
	ELSE
	BEGIN
		SET @output__DisplayMatrix = 0;
	END

	IF @HoursForward > 0 
	BEGIN
		SET @output__DisplayPredictive = 1;
	END
	ELSE
	BEGIN
		SET @output__DisplayPredictive = 0;
	END

	--Note that we sent the output type even if one or both of the matrices are not requested
	IF @ToConsole = N'N'
	BEGIN
		SET @outputType__Matrix = N'XML'
		SET @outputType__Predictive = N'XML'
	END
	ELSE
	BEGIN
		--only 1 thing can be sent to the console. Historical matrix gets priority
		IF @output__DisplayMatrix = 1
		BEGIN
			SET @outputType__Matrix = N'CONSOLE'
			SET @outputType__Predictive = N'XML'
		END
		ELSE
		BEGIN
			SET @outputType__Matrix = N'XML'

			IF @outputType__Predictive = 1
			BEGIN
				SET @outputType__Predictive = N'CONSOLE'
			END
			ELSE
			BEGIN
				SET @outputType__Predictive = N'XML'
			END
		END
	END


	--***Location 1: Parameter NULL handling
	SET @PointInTime = ISNULL(@PointInTime, GETDATE());
	SET @HoursBack = ISNULL(@HoursBack,20);
	SET @HoursForward = ISNULL(@HoursForward,0);
	SET @ToConsole = ISNULL(@ToConsole, GETDATE());
	SET @FitOnScreen = ISNULL(@FitOnScreen, N'Y');
	SET @DisplayConfigOptions = ISNULL(@DisplayConfigOptions,1);
	SET @DisplayAgentLog = ISNULL(@DisplayAgentLog,1);
	SET @Queries = ISNULL(@Queries,N'N');
	SET @Help = ISNULL(@Help, N'N');
	SET @Debug = ISNULL(@Debug,0);

	--***Location 2: Console print of example proc exec
	PRINT @lv__HelpText;

	IF @Help <> N'N'
	BEGIN
		GOTO helploc
	END

	--***Location 3: Parameter value validation
	IF @PointInTime < CONVERT(DATETIME, '2000-01-01') OR @PointInTime > DATEADD(DAY, 3, GETDATE())
	BEGIN
		RAISERROR('The @PointInTime value is restricted to values between 2000-01-01 and 3 days into the future.',15,1);
		RETURN -1;
	END

	IF ISNULL(@HoursBack,-1) < 0 or ISNULL(@HoursBack,-1) > 48
	BEGIN
		RAISERROR('The @HoursBack parameter must be a non-null value between 0 and 48 inclusive.', 15, 1);
		RETURN -1;
	END

	IF ISNULL(@HoursForward,-1) < 0 or ISNULL(@HoursForward,-1) > 20
	BEGIN
		RAISERROR('The @HoursForward parameter must be a non-null value between 0 and 20 inclusive.', 15, 1);
		RETURN -1;
	END


	IF @ToConsole NOT IN (N'N', N'Y')
	BEGIN
		RAISERROR('The @ToConsole parameter must be either Y or N',15,1);
		RETURN -1;
	END

	IF @FitOnScreen NOT IN (N'N', N'Y')
	BEGIN
		RAISERROR('The @FitOnScreen parameter must be either Y or N',15,1);
		RETURN -1;
	END

	IF @DisplayAgentLog NOT IN (0, 1, 2, 3)
	BEGIN
		RAISERROR('The @DisplayAgentLog parameter must be one of the following: 0 = No; 1 = Currently log only, and only when Sev 1 records exist; 2 = Always display current log; 3 = Always display last 3 log files', 15, 1);
		RETURN -1;
	END

	IF @DisplayConfigOptions NOT IN (0, 1, 2)
	BEGIN
		RAISERROR('The @DisplayConfigOptions parameter must be either 0 = No, 1 = only different from default, or 2 = Always display.', 15, 1);
		RETURN -1;
	END

	--If the user attempts to type anything for this variable, set to N'Y'
	IF ISNULL(@Queries, N'Y') <> N'N'
	BEGIN
		SET @Queries = N'Y'
	END

	--***Location 4: Temp table definition
	--Holds the "header" info for the matrix, and also holds the start and end times of each of our time windows.

	CREATE TABLE #TimeWindows_Hist (
		WindowID INT NOT NULL PRIMARY KEY CLUSTERED,
		WindowBegin DATETIME NOT NULL,
		WindowEnd DATETIME NOT NULL,
		TimeHeaderChar NCHAR(1) NOT NULL,
		LineHeaderChar NCHAR(1) NOT NULL
	);

	CREATE TABLE #TimeWindows_Pred (
		WindowID INT NOT NULL PRIMARY KEY CLUSTERED,
		WindowBegin DATETIME NOT NULL,
		WindowEnd DATETIME NOT NULL,
		TimeHeaderChar NCHAR(1) NOT NULL,
		LineHeaderChar NCHAR(1) NOT NULL
	);

	--A list of SQL Agent jobs, the # of runs and failures, which sub-matrix the job falls into, and the
	-- display order within the sub-matrix
	CREATE TABLE #Jobs (
		JobID INT NOT NULL IDENTITY PRIMARY KEY CLUSTERED,
		JobName NVARCHAR(256) NOT NULL,
		IsEnabled TINYINT NOT NULL,
		Notifies TINYINT NOT NULL,
		CreateDate DATETIME NOT NULL, 
		LastModifiedDate DATETIME NOT NULL,
		OwnerPrincipalName VARCHAR(256),
		native_job_id UNIQUEIDENTIFIER NOT NULL,
		JobRuns INT NOT NULL,
		JobFailures INT NOT NULL,
		CompletionsAllTime INT NULL,
		AvgJobDur_seconds BIGINT NULL,		--average duration, including failures.
		AvgSuccessDur_seconds BIGINT NULL,	--average duration, only including successes. Either measure can be a faulty predictor
												-- of future duration. (e.g. a failure occurring almost right away, leading to a very short duration)
		MatrixNumber INT NOT NULL,
		DisplayOrder INT NOT NULL, 
		StepTypes VARCHAR(100)			--comma-delimited list of the different TYPES of steps for the job
	);

	--Populated with a cross join between the #TimeWindows_Hist table and the #Jobs table, which logically gives us
	-- each matrix "line" (a series of cells/time windows for each job)
	CREATE TABLE #JobArrays_Hist (
		JobID INT NOT NULL, 
		WindowID INT NOT NULL, 
		WindowBegin DATETIME NOT NULL, 
		WindowEnd DATETIME NOT NULL, 
		CellText NCHAR(1) NOT NULL
	);

	CREATE TABLE #JobArrays_Pred (
		JobID INT NOT NULL, 
		WindowID INT NOT NULL, 
		WindowBegin DATETIME NOT NULL, 
		WindowEnd DATETIME NOT NULL, 
		CellText NCHAR(1) NOT NULL
	);

	--TODO: some query plans might benefit from an index here. Investigate the best
	-- field for the clustered index key
	CREATE TABLE #JobInstances (
		native_job_id UNIQUEIDENTIFIER NOT NULL, 
		job_run_status INT,				--0=failed; 1=succeeded; 2=retry; 3=cancelled
		JobStartTime DATETIME, 
		JobEndTime DATETIME, 
		JobDisplayEndTime DATETIME,		--helps us do certain display logic for jobs that are "still running" (i.e. don't have an end time)
		JobExpectedEndTime DATETIME
	);

	CREATE TABLE #HypotheticalRuns (
		native_job_id UNIQUEIDENTIFIER NOT NULL, 
		JobStartTime DATETIME,
		JobExpectedEndTime DATETIME
	)

	--Holds the contents of master.dbo.xp_sqlagent_enum_jobs so that we can determine which jobs are 
	-- running (and thus don't have a completion record in the sysjobhistory table)
	CREATE TABLE #CurrentlyRunningJobs1 ( 
		Job_ID UNIQUEIDENTIFIER,
		Last_Run_Date INT,
		Last_Run_Time INT,
		Next_Run_Date INT,
		Next_Run_Time INT,
		Next_Run_Schedule_ID INT,
		Requested_To_Run INT,
		Request_Source INT,
		Request_Source_ID VARCHAR(100),
		Running INT,
		Current_Step INT,
		Current_Retry_Attempt INT, 
		aState INT
	);

	CREATE TABLE #CurrentlyRunningJobs2 (
		native_job_id UNIQUEIDENTIFIER NOT NULL, 
		JobStartTime DATETIME, 
		JobEndTime DATETIME, 
		JobDisplayEndTime DATETIME,		--helps us do certain display logic for jobs that are "still running" (i.e. don't have an end time)
		JobExpectedEndTime DATETIME
	);

	--a table to cache query results for job schedules
	CREATE TABLE #SchedulesForThisJob (
		freq_type smallint not null, 
		freq_interval smallint not null,
		freq_subday_type smallint not null,
		freq_subday_interval smallint not null,
		freq_relative_interval smallint not null,
		freq_recurrence_factor smallint not null,
		active_start_date int,
		active_end_date int,
		active_start_time int, 
		active_end_time int,
		ActiveStartTime datetime, 
		ActiveEndTime datetime,
		NextScheduleTimeFromSystem datetime
	);

	CREATE TABLE #ScheduleNextExecTimes (
		ExecStartTime DATETIME NOT NULL
	);

	--We place various substrings here before assembling them into the XML value
	CREATE TABLE #OutputStringLineItems_Hist (
		RowType TINYINT, 
		JobID INT, 
		MatrixNumber INT, 
		DisplayOrder INT, 
		CellString NVARCHAR(MAX)
	);

	CREATE TABLE #OutputStringLineItems_Pred (
		RowType TINYINT, 
		JobID INT, 
		MatrixNumber INT, 
		DisplayOrder INT, 
		CellString NVARCHAR(MAX)
	);

	--Where we place the SQL Agent log messages before we assemble them into the XML value
	CREATE TABLE #SQLAgentLog (
		idcol INT IDENTITY PRIMARY KEY CLUSTERED, 
		FileNumber INT, 
		isLastRecord INT, 
		LogDate DATETIME, 
		ErrorLevel INT, 
		aText VARCHAR(MAX)
	);

	--Config option list
	CREATE TABLE #OptionsToDisplay (
		idcol INT IDENTITY PRIMARY KEY, 
		OptionTag varchar(100), 
		OptionValue varchar(100), 
		OptionNormalValue varchar(100)
	);

	SET @lv__afterdt = GETDATE();

	--If the proc runs longer than expected, scattering these duration tests throughout the proc, then sending the
	-- results to the console, will help the user understand where the proc is taking its time.
	IF DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt) > @lv__slownessthreshold
	BEGIN
		SET @lv__ErrorText = N'   ***dbg: Proc setup and Temp Table creation took ' + CONVERT(NVARCHAR(40),DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt)) + N' milliseconds';
		RAISERROR(@lv__ErrorText, 10, 1) WITH NOWAIT;
	END
	/*********************************************************************************************************************************************
	*********************************************************************************************************************************************

														Part 1: Job History Matrix

	*********************************************************************************************************************************************
	*********************************************************************************************************************************************/
	SET @lv__beforedt = GETDATE();

	--If we're skipping both matrices: 
	IF @output__DisplayMatrix = 0 AND @output__DisplayPredictive = 0
	BEGIN
		GOTO configstart
	END

	--NOTE: for now, even if the historical matrix is off and the predictive matrix is on, we are going to execute
	-- the following historical-matrix logic, rather than move the job-info-collection logic to an earlier point in
	-- the procedure.

	--***Location 2: Determine matrix dimensions
	-- A cell time window begins at the zero second of a (e.g. 10)-minute time boundary, and ends immediately before
	-- a new (e.g. 10)-minute time boundary starts. Let's find out how wide our window lengths are, and derive the
	-- @lv__mtx__MatrixWidth from that.

	/* The calculation for @lv__mtx__WindowLength_minutes and @lv__mtx__MatrixWidth deserves a full explanation:

		Several design goals heavily influence the characteristics of the job history matrix:
			- The user can control whether everything fits on one screen or not (@FitOnScreen)
			- We want the top of the hour to always align with the start time of a time window of a matrix cell,
				(i.e. you'd never have a time window like '10:55:00 - 11:04:59')
				which means the possible values for @lv__mtx__WindowLength_minutes must always be a root of 60.
				This limits us to 1, 2, 3, 4, 5, 6, 10, 15, 20, 30, and 60

			- The fact that we want to print numeric labels for the top-of-the-hour effectively rules out
				60 (because they would overlap) and 30 (because they would be immediately adjacent)

			- It is also nice for the time windows to align with intuitive sub-divisions of the 60-minutes in
				an hour, especially the half-hour. Thus, a time window of 10 minutes is great because you have 
				6 intuitive markers and can place a special tick mark/label identifying the 30-minutes.
				
				Conversely, a time window length of 4 minutes is not ideal, since while it is a root of 60, it is
				not a root of 30.
				
				(Note that having a time window length of 20 minutes is also somewhat intuitive, and is handy
				when @HoursBack is very large).

		The net effect of all this is that the time windows allowed for our matrix are:
				1, 2, 3, 5, 6, 10, 15, and 20


		The time window length also depends on whether the user wants everything in 1 screen, or is ok with scrolling.
		If @FitsOnScreen='Y', the matrix is kept in between 100 and 145 characters (with a couple exceptions on the lower side)
		and if ='N', then the matrix is kept to 360 maximum.

		Here's the complete chart for the historical matrix:

	If @FitOnScreen = 'Y'
		@HoursBack = 1 --> cell width = 1		MatrixWidth = 60		(60-wide doesn't look as good, so we actually bump it up to 1.5 hours)
		@HoursBack = 2 --> cell width = 1		MatrixWidth = 120
		@HoursBack = 3 --> cell width = 2		MatrixWidth = 90
		@HoursBack = 4 --> cell width = 2		MatrixWidth = 120
		@HoursBack = 5 --> cell width = 3/2		MatrixWidth = 100/150	(100 width for XML, 150 width for console)
		@HoursBack = 6 --> cell width = 3		MatrixWidth = 120
		@HoursBack = 7 --> cell width = 3		MatrixWidth = 140
		@HoursBack = 8 --> cell width = 5		MatrixWidth = 96	
		@HoursBack = 9 --> cell width = 5		MatrixWidth = 108	
		@HoursBack =10 --> cell width = 5		MatrixWidth = 120
		@HoursBack =11 --> cell width = 5		MatrixWidth = 132
		@HoursBack =12 --> cell width = 5		MatrixWidth = 144
		@HoursBack =13 --> cell width = 6		MatrixWidth = 130
		@HoursBack =14 --> cell width = 6		MatrixWidth = 140
		@HoursBack =15 --> cell width = 10		MatrixWidth = 90	
		@HoursBack =16 --> cell width = 10		MatrixWidth = 96	
		@HoursBack =17 --> cell width = 10		MatrixWidth = 102	
		@HoursBack =18 --> cell width = 10		MatrixWidth = 108	
		@HoursBack =19 --> cell width = 10		MatrixWidth = 114	
		@HoursBack =20 --> cell width = 10		MatrixWidth = 120
		@HoursBack =21 --> cell width = 10		MatrixWidth = 126
		@HoursBack =22 --> cell width = 10		MatrixWidth = 132
		@HoursBack =23 --> cell width = 10		MatrixWidth = 138
		@HoursBack =24 --> cell width = 15		MatrixWidth = 96	
		@HoursBack =25 --> cell width = 15		MatrixWidth = 100
		@HoursBack =26 --> cell width = 15		MatrixWidth = 104
		@HoursBack =27 --> cell width = 15		MatrixWidth = 108
		@HoursBack =28 --> cell width = 15		MatrixWidth = 112
		@HoursBack =29 --> cell width = 15		MatrixWidth = 116
		@HoursBack =30 --> cell width = 15		MatrixWidth = 120
		@HoursBack =31 --> cell width = 15		MatrixWidth = 124
		@HoursBack =32 --> cell width = 15		MatrixWidth = 128
		@HoursBack =33 --> cell width = 15		MatrixWidth = 132
		@HoursBack =34 --> cell width = 15		MatrixWidth = 136
		@HoursBack =35 --> cell width = 15		MatrixWidth = 140
		@HoursBack =36 --> cell width = 20		MatrixWidth = 108
		@HoursBack =37 --> cell width = 20		MatrixWidth = 111
		@HoursBack =38 --> cell width = 20		MatrixWidth = 114
		@HoursBack =39 --> cell width = 20		MatrixWidth = 117
		@HoursBack =40 --> cell width = 20		MatrixWidth = 120
		@HoursBack =41 --> cell width = 20		MatrixWidth = 123
		@HoursBack =42 --> cell width = 20		MatrixWidth = 126
		@HoursBack =43 --> cell width = 20		MatrixWidth = 129
		@HoursBack =44 --> cell width = 20		MatrixWidth = 132
		@HoursBack =45 --> cell width = 20		MatrixWidth = 135
		@HoursBack =46 --> cell width = 20		MatrixWidth = 138
		@HoursBack =47 --> cell width = 20		MatrixWidth = 141
		@HoursBack =48 --> cell width = 20		MatrixWidth = 144


	If @FitOnScreen = 'N'

		@HoursBack = 1 --> cell width = 1		MatrixWidth = 60	
		@HoursBack = 2 --> cell width = 1		MatrixWidth = 120
		@HoursBack = 3 --> cell width = 1		MatrixWidth = 180
		@HoursBack = 4 --> cell width = 1		MatrixWidth = 240
		@HoursBack = 5 --> cell width = 1		MatrixWidth = 300
		@HoursBack = 6 --> cell width = 1		MatrixWidth = 360

		@HoursBack = 7 --> cell width = 2		MatrixWidth = 210
		@HoursBack = 8 --> cell width = 2		MatrixWidth = 240
		@HoursBack = 9 --> cell width = 2		MatrixWidth = 270
		@HoursBack =10 --> cell width = 2		MatrixWidth = 300
		@HoursBack =11 --> cell width = 2		MatrixWidth = 330
		@HoursBack =12 --> cell width = 2		MatrixWidth = 360

		--Let's skip cell widths = 3 and 4 because they don't line up with tick marks as well

		@HoursBack =13 --> cell width = 5		MatrixWidth = 156
		@HoursBack =14 --> cell width = 5		MatrixWidth = 168
		@HoursBack =15 --> cell width = 5		MatrixWidth = 180
		@HoursBack =16 --> cell width = 5		MatrixWidth = 192
		@HoursBack =17 --> cell width = 5		MatrixWidth = 204
		@HoursBack =18 --> cell width = 5		MatrixWidth = 216

		@HoursBack =25 --> cell width = 5		MatrixWidth = 300	
		@HoursBack =26 --> cell width = 5		MatrixWidth = 312
		@HoursBack =27 --> cell width = 5		MatrixWidth = 324
		@HoursBack =28 --> cell width = 5		MatrixWidth = 336
		@HoursBack =29 --> cell width = 5		MatrixWidth = 348
		@HoursBack =30 --> cell width = 5		MatrixWidth = 360

		@HoursBack =31 --> cell width = 10		MatrixWidth = 186
		@HoursBack =32 --> cell width = 10		MatrixWidth = 192
		@HoursBack =33 --> cell width = 10		MatrixWidth = 198
		@HoursBack =34 --> cell width = 10		MatrixWidth = 204
		@HoursBack =35 --> cell width = 10		MatrixWidth = 210
		@HoursBack =36 --> cell width = 10		MatrixWidth = 216

		@HoursBack =37 --> cell width = 10		MatrixWidth = 222
		@HoursBack =38 --> cell width = 10		MatrixWidth = 228
		@HoursBack =39 --> cell width = 10		MatrixWidth = 234
		@HoursBack =40 --> cell width = 10		MatrixWidth = 240
		@HoursBack =41 --> cell width = 10		MatrixWidth = 246
		@HoursBack =42 --> cell width = 10		MatrixWidth = 252

		@HoursBack =43 --> cell width = 10		MatrixWidth = 258
		@HoursBack =44 --> cell width = 10		MatrixWidth = 264
		@HoursBack =45 --> cell width = 10		MatrixWidth = 270
		@HoursBack =46 --> cell width = 10		MatrixWidth = 276
		@HoursBack =47 --> cell width = 10		MatrixWidth = 282
		@HoursBack =48 --> cell width = 10		MatrixWidth = 288
	*/

	IF @FitOnScreen = N'Y'
	BEGIN
		SELECT @lv__mtx__WindowLength_minutes = CASE 
				WHEN @HoursBack BETWEEN 1 AND 2 THEN 1
				WHEN @HoursBack BETWEEN 3 AND 4 THEN 2
				WHEN @HoursBack = 5
					THEN (
						CASE WHEN @ToConsole = N'N' THEN 3
							ELSE 2
						END
					)
				WHEN @HoursBack BETWEEN 6 AND 7 THEN 3
				WHEN @HoursBack BETWEEN 8 AND 12 THEN 5
				WHEN @HoursBack BETWEEN 13 AND 14 THEN 6
				WHEN @HoursBack BETWEEN 15 AND 23 THEN 10
				WHEN @HoursBack BETWEEN 24 AND 35 THEN 15
				WHEN @HoursBack BETWEEN 36 AND 48 THEN 20
			ELSE 1	--shouldn't hit this
			END;
	END
	ELSE
	BEGIN
		SELECT @lv__mtx__WindowLength_minutes = CASE 
				WHEN @HoursBack BETWEEN 1 AND 6 THEN 1
				WHEN @HoursBack BETWEEN 7 AND 12 THEN 2
				WHEN @HoursBack BETWEEN 13 AND 30 THEN 5
				WHEN @HoursBack BETWEEN 31 AND 48 THEN 10
			ELSE 1 --shouldn't hit this
			END;
	END

	--For @HoursBack=1, since our minimum cell width is 1 minute, we only end up with a 60-char wide matrix, and only 1
	-- hour-marker in the header. That doesn't look as good, so let's bump up the size of the matrix (and the time window)
	-- by 30
	IF @HoursBack = 1
	BEGIN
		SET @lv__mtx__MatrixWidth = 90;
	END
	ELSE
	BEGIN
		--Matrix width is easy to calculate once we have window length
		SET @lv__mtx__MatrixWidth = @HoursBack*60 / @lv__mtx__WindowLength_minutes;
	END
	
	--For the "Time Header" line, we want to mark inter-hour "landmarks" to make rough time identification easier.
	--We also want to do something similar for the "Line Header" line
	IF @lv__mtx__WindowLength_minutes IN (1,2)
	BEGIN
		SET @lv__mtx__LineHeaderMod = 10;		--print ticks every 10 minutes
		SET @lv__mtx__TimeHeaderMod = 20;		--print '+' chars every 20 min, but not on the hour
	END
	ELSE IF @lv__mtx__WindowLength_minutes IN (3,5)
	BEGIN
		SET @lv__mtx__LineHeaderMod = 15;
		SET @lv__mtx__TimeHeaderMod = 30;
	END
	ELSE IF @lv__mtx__WindowLength_minutes IN (6,10,15)
	BEGIN
		SET @lv__mtx__LineHeaderMod = 30;
		SET @lv__mtx__TimeHeaderMod = 30;
	END
	ELSE 
		--IF @lv__mtx__WindowLength_minutes = 20		the only other option at this time is 20
	BEGIN
		SET @lv__mtx__LineHeaderMod = -1;
		SET @lv__mtx__TimeHeaderMod = -1;
	END

	--Because (n)varchar strings are trimmed, we use an underscore for most of the string manipulation and then 
	--do a REPLACE(<expr>, @lv__mtx__EmptyChar, N' ') at the end.
	SET @lv__mtx__EmptyChar = N'_';

	--***Location 3: Determine last window of matrix
	--The @PointInTime is very likely NOT the exact endpoint for an x-minute time window. Let's find the endpoint for the
	-- time window that we are in currently.
	BEGIN TRY
		SELECT 
				@lv__mtx__CurrentTime_WindowBegin = ss3.CurrentTime_WindowBegin,
				--Aaron: changing the end time to be equal to the start time of the next window, rather than a few milliseconds before, 
				-- to avoid the off-chance that a job completion slips through the cracks
				--@lv__mtx__CurrentTime_WindowEnd = DATEADD(MILLISECOND, -10, DATEADD(MINUTE, @lv__mtx__WindowLength_minutes, CurrentTime_WindowBegin))
				@lv__mtx__CurrentTime_WindowEnd = DATEADD(MINUTE, @lv__mtx__WindowLength_minutes, CurrentTime_WindowBegin)
			FROM (
				SELECT [CurrentTime_WindowBegin] = DATEADD(MINUTE, NthWindowFromTopOfHour*@lv__mtx__WindowLength_minutes, CurrentTime_HourBase), 
					CurrentTime, 
					CurrentTime_HourBase, 
					CurrentMinute, 
					CurrentHour, 
					NthWindowFromTopOfHour
				FROM (
					SELECT [NthWindowFromTopOfHour] = CurrentMinute / @lv__mtx__WindowLength_minutes,		--zero-based, of course
						[CurrentTime_HourBase] = DATEADD(HOUR, CurrentHour, 
																CONVERT(DATETIME,
																		CONVERT(VARCHAR(20), CurrentTime, 101)
																		)
														),
						CurrentTime, 
						CurrentMinute, 
						CurrentHour 
					FROM (
						SELECT [CurrentMinute] = DATEPART(MINUTE, CurrentTime), 
							[CurrentHour] = DATEPART(HOUR, CurrentTime),
							CurrentTime
						FROM 
							(SELECT [CurrentTime] = @PointInTime) ss0
						) ss1
					) ss2
			) ss3
			;
	END TRY
	BEGIN CATCH
		RAISERROR(N'Unable to construct the final time window. The job history matrix will not be displayed.', 11, 1);
		SET @output__DisplayMatrix = 0;
		SELECT @lv__ErrorText = ERROR_MESSAGE(), 
				@lv__ErrorSeverity	= ERROR_SEVERITY(), 
				@lv__ErrorState = ERROR_STATE();
		SET @lv__ErrorText = @lv__ErrorText + ' (Severity: ' + CONVERT(VARCHAR(20),@lv__ErrorSeverity) + ') (State: ' + CONVERT(VARCHAR(20),@lv__ErrorState) + ')'

		RAISERROR( @lv__ErrorText, 11, 1);

		GOTO aftermatrix
	END CATCH

	----***Location 4: Now build our array of time windows for the whole matrix
	BEGIN TRY
		;WITH t0 AS (
			SELECT 0 as col1 UNION ALL
			SELECT 0 UNION ALL
			SELECT 0 UNION ALL
			SELECT 0
		),
		t1 AS (
			SELECT ref1.col1 FROM t0 as ref1
				CROSS JOIN t0 as ref2
				CROSS JOIN t0 as ref3
				CROSS JOIN t0 as ref4
		),
		nums AS (
			SELECT ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) as rn
			FROM t1
		)
		INSERT INTO #TimeWindows_Hist (WindowID, WindowBegin, WindowEnd, TimeHeaderChar, LineHeaderChar)
		SELECT 
			CellReverseOrder, 
			WindowBegin,
			WindowEnd,
			TimeHeaderChar = (
				CASE 
					--When we are on the top of the hour, we usually print hour information (first digit of a 2-digit hour)
					WHEN NthWindowFromTopOfHour = 0
						THEN (CASE 
								--For @HoursBack>24, print even hours. Otherwise, always print the hour
								WHEN @HoursBack > 24 
									THEN (
										CASE WHEN DATEPART(HOUR, WindowBegin) % 2 = 0 THEN SUBSTRING(CONVERT(NVARCHAR(20),DATEPART(HOUR, WindowBegin)),1,1)
											ELSE N'.'
										END 
									)
								ELSE SUBSTRING(CONVERT(NVARCHAR(20),DATEPART(HOUR, WindowBegin)),1,1)
								END
								)
					--When it is the second window of the hour, we check to see if we have a double-digit hour # and print the second digit
					WHEN NthWindowFromTopOfHour = 1
						THEN (
							CASE WHEN DATEPART(HOUR, WindowBegin) < 10 THEN N'.'
								ELSE SUBSTRING(REVERSE(CONVERT(NVARCHAR(20),DATEPART(HOUR, WindowBegin))),1,1)
							END
						)
					--should we print the Time Header intra-hour marker?
					WHEN @lv__mtx__TimeHeaderMod <> -1 AND DATEPART(MINUTE, WindowBegin) % @lv__mtx__TimeHeaderMod = 0
						THEN (CASE 
								WHEN @HoursBack > 24 THEN '.'		--too high-level for intra-hour markers
								ELSE '+'
							END
						)
					ELSE '.'	--should never hit this case
				END 
				),
			LineHeaderChar = (
				CASE 
					WHEN DATEPART(MINUTE, WindowBegin) % @lv__mtx__LineHeaderMod = 0 THEN '|' 
					ELSE '-'
				END 
				)
		FROM (
			SELECT 
				CellReverseOrder, 
				CurrentTime_WindowBegin, 
				CurrentTime_WindowEnd, 
				WindowBegin, 
				WindowEnd,
				[NthWindowFromTopOfHour] = DATEPART(MINUTE, WindowBegin)  / @lv__mtx__WindowLength_minutes
			FROM (
				SELECT TOP (@lv__mtx__MatrixWidth) 
					rn as CellReverseOrder,
					@lv__mtx__CurrentTime_WindowBegin as CurrentTime_WindowBegin, 
					@lv__mtx__CurrentTime_WindowEnd as CurrentTime_WindowEnd,
					DATEADD(MINUTE, 0-@lv__mtx__WindowLength_minutes*(rn-1), @lv__mtx__CurrentTime_WindowBegin) as WindowBegin,
					DATEADD(MINUTE, 0-@lv__mtx__WindowLength_minutes*(rn-1), @lv__mtx__CurrentTime_WindowEnd) as WindowEnd
				FROM nums 
				ORDER BY rn ASC
			) ss0
		) ss1
		OPTION(MAXDOP 1);
	END TRY
	BEGIN CATCH
		RAISERROR(N'Unable to define the complete list of time window boundaries. The job history matrix will not be displayed.', 11, 1);
		SET @output__DisplayMatrix = 0;
		SELECT @lv__ErrorText = ERROR_MESSAGE(), 
				@lv__ErrorSeverity	= ERROR_SEVERITY(), 
				@lv__ErrorState = ERROR_STATE();
		SET @lv__ErrorText = @lv__ErrorText + ' (Severity: ' + CONVERT(VARCHAR(20),@lv__ErrorSeverity) + ') (State: ' + CONVERT(VARCHAR(20),@lv__ErrorState) + ')'

		RAISERROR( @lv__ErrorText, 11, 1);

		GOTO aftermatrix
	END CATCH

	IF @Debug = 1
	BEGIN
		SELECT 'Contents of #TimeWindows_Hist' as DebugLocation, tw.WindowID, tw.WindowBegin, tw.WindowEnd, tw.TimeHeaderChar, tw.LineHeaderChar
		FROM #TimeWindows_Hist tw
		ORDER BY tw.WindowID;
	END

	--Get overall min/max times, as we'll use these later in the proc
	SELECT 
		@lv__mtx__OverallWindowBeginTime = MIN(tw.WindowBegin), 
		@lv__mtx__OverallWindowEndTime = MAX(tw.WindowEnd)
	FROM #TimeWindows_Hist tw;

	/*
	select getdate() as curtm,
		@lv__mtx__CurrentTime_WindowBegin as curBegin, @lv__mtx__CurrentTime_WindowEnd as curEnd,
		@lv__mtx__OverallWindowBeginTime as overallBegin, @lv__mtx__OverallWindowEndTime as overallEnd,
		@lv__mtx__MatrixWidth as MatrixWidth, @lv__mtx__WindowLength_minutes as WinLength_min

	select * from #TimeWindows_Hist tw order by tw.WindowID desc
	*/

	SET @lv__afterdt = GETDATE();

	IF DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt) > @lv__slownessthreshold
	BEGIN
		SET @lv__ErrorText = N'   ***dbg: constructing Historical Matrix time windows took ' + CONVERT(NVARCHAR(40),DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt)) + N' milliseconds';
		RAISERROR(@lv__ErrorText, 10, 1) WITH NOWAIT;
	END

	IF @Debug = 1
	BEGIN
		SELECT @lv__mtx__OverallWindowBeginTime as OverallWindowBeginTime, @lv__mtx__OverallWindowEndTime as OverallWindowEndTime;
	END
	
	
	--***Location 0: Obtain SQL Server and SQL Agent start times
	SET @lv__beforedt = GETDATE();

	SELECT @lv__mtx__SQLServerStartTime = d.create_date FROM sys.databases d WHERE d.database_id = 2;

	SELECT @lv__mtx__SQLAgentStartTime = ss.agent_start_date
	FROM (
		SELECT TOP 1 s.agent_start_date
		FROM msdb.dbo.syssessions s
		WHERE s.agent_start_date < @lv__mtx__OverallWindowEndTime
		ORDER BY s.agent_start_date DESC
	) ss;

	--Our @PointInTime could be older than the most recent SQL Server restart time. Since SQL Server log files can be quite large, we're not going
	-- to go digging in them for the previous restart time.
	IF @lv__mtx__SQLServerStartTime IS NOT NULL
	BEGIN
		IF @lv__mtx__SQLServerStartTime > @lv__mtx__OverallWindowEndTime
		BEGIN
			SET @lv__mtx__SQLStartTimeResult = 2;
		END
		ELSE 
		BEGIN
			SET @lv__mtx__SQLStartTimeResult = 1;
		END
	END

	SET @lv__afterdt = GETDATE();

	IF DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt) > @lv__slownessthreshold
	BEGIN
		SET @lv__ErrorText = N'   ***dbg: obtaining SQL Server and Agent start times took ' + CONVERT(NVARCHAR(40),DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt)) + N' milliseconds';
		RAISERROR(@lv__ErrorText, 10, 1) WITH NOWAIT;
	END

	SET @lv__beforedt = GETDATE();

	--***Location 1: Get da jobs
	BEGIN TRY
		INSERT INTO #Jobs (
			JobName, 
			IsEnabled, 
			Notifies, 
			CreateDate, 
			LastModifiedDate,  
			OwnerPrincipalName,
			native_job_id, 
			JobRuns, 
			JobFailures, 
			MatrixNumber, 
			DisplayOrder)
		SELECT 
			j.name, 
			j.enabled, 
			CASE WHEN j.notify_level_email > 0 OR j.notify_level_netsend > 0 OR j.notify_level_page > 0 THEN 1 ELSE 0 END,
			j.date_created, 
			j.date_modified,
			p.name,
			j.job_id, 
			0, 
			0, 
			5,	--start off assuming that each job lacks a successful completion. We'll change this field after examining job history
			ROW_NUMBER() OVER (ORDER BY j.name ASC)
		FROM msdb.dbo.sysjobs j 
			INNER JOIN sys.server_principals p
				ON j.owner_sid = p.sid 
		WHERE 1=1
		/*
		j.date_created < @lv__mtx__OverallWindowEndTime		--Don't show a job if it didn't exist before the end time of our matrix
															NOTE: changed for now... putting job names in parentheses
		*/
		;
	END TRY
	BEGIN CATCH
		RAISERROR(N'Unable to obtain a list of jobs from msdb.dbo.sysjobs. The job matrices will not be displayed.', 11, 1);
		SET @output__DisplayMatrix = 0;
		SET @output__DisplayPredictive = 0;
		SELECT @lv__ErrorText = ERROR_MESSAGE(), 
				@lv__ErrorSeverity	= ERROR_SEVERITY(), 
				@lv__ErrorState = ERROR_STATE();
		SET @lv__ErrorText = @lv__ErrorText + ' (Severity: ' + CONVERT(VARCHAR(20),@lv__ErrorSeverity) + ') (State: ' + CONVERT(VARCHAR(20),@lv__ErrorState) + ')'

		RAISERROR( @lv__ErrorText, 11, 1);

		GOTO configstart
	END CATCH

	
	IF NOT EXISTS (SELECT 1 FROM #Jobs)
	BEGIN
		PRINT ('No SQL Agent jobs found on this instance. No Job matrices will be printed')
		SET @output__DisplayMatrix = 0;
		SET @output__DisplayPredictive = 0;

		GOTO configstart
	END

	SET @lv__afterdt = GETDATE();

	IF DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt) > @lv__slownessthreshold
	BEGIN
		SET @lv__ErrorText = N'   ***dbg: obtaining list of jobs took ' + CONVERT(NVARCHAR(40),DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt)) + N' milliseconds';
		RAISERROR(@lv__ErrorText, 10, 1) WITH NOWAIT;
	END

	IF @Debug = 1
	BEGIN
		SELECT j.JobID, j.JobName, j.native_job_id, j.IsEnabled, j.JobRuns, j.JobFailures, j.MatrixNumber, j.DisplayOrder
		FROM #Jobs j
		ORDER BY j.JobName ASC
	END


	SET @lv__beforedt = GETDATE();

	--***Location 5: Get job completion information
	--The msdb.dbo.sysjobhistory table stores the relevant info in a format that isn't easy to use naturally. Convert the relevant data first.
	BEGIN TRY
		;WITH Job_Completions AS (
			SELECT job_id, run_status, run_date, run_time, run_duration,
				JobStartTime,  	
				[JobEndTime] = (
					DATEADD(HOUR, 
						CONVERT(INT,REVERSE(SUBSTRING(DurationReversed, 5,6))), 
						DATEADD(MINUTE, 
							CONVERT(INT,REVERSE(SUBSTRING(DurationReversed, 3,2))),
							DATEADD(SECOND,
								CONVERT(INT,REVERSE(SUBSTRING(DurationReversed, 1,2))),
								JobStartTime
								)
							)
						)
					)
			FROM (
				SELECT h.job_id
					,h.run_status
					,h.run_date, h.run_time, h.run_duration
					,[JobStartTime] = (
						CASE WHEN (h.run_date IS NULL OR h.run_Time IS NULL 
								OR h.run_date < 19000101 OR h.run_time < 0
								OR h.run_time > 235959)
								THEN NULL 
							ELSE CAST(STR(h.run_date, 8, 0) AS DATETIME) + 
								CAST(STUFF(STUFF(REPLACE(STR(h.run_time, 6), ' ', '0'), 3, 0, ':'), 6, 0, ':') AS DATETIME)
							END)
					,[DurationReversed] = CASE
						WHEN h.run_duration IS NULL THEN NULL 
						WHEN h.run_duration < 0 THEN NULL
						ELSE REVERSE(REPLACE(STR(h.run_duration, 10),' ', '0'))
						END 
				FROM msdb.dbo.sysjobhistory h WITH (NOLOCK)
				WHERE h.step_id = 0		--only look at completion states
			) ss
		) 
		INSERT INTO #JobInstances (native_job_id, job_run_status, JobStartTime, JobEndTime, JobDisplayEndTime)
		SELECT 
			jc.job_id 
			
			--We need to handle when @PointInTime is in the past; jobs that have finished now may have been running
			-- at the "end time" the user has requested.
			,CASE WHEN @lv__mtx__OverallWindowEndTime BETWEEN jc.JobStartTime AND jc.JobEndTime
				THEN 25  --special code for "Running"
				ELSE jc.run_status 
			 END

			,jc.JobStartTime
			,jc.JobEndTime
			,jc.JobEndTime			--since all of these jobs have already finished, the display time = the actual endtime
		FROM Job_Completions jc
		OPTION(MAXDOP 1);
	END TRY
	BEGIN CATCH
		RAISERROR(N'Unable to obtain job completion information. The job matrices will not be displayed.', 11, 1);
		SET @output__DisplayMatrix = 0;
		SET @output__DisplayPredictive = 0;
		SELECT @lv__ErrorText = ERROR_MESSAGE(), 
				@lv__ErrorSeverity	= ERROR_SEVERITY(), 
				@lv__ErrorState = ERROR_STATE();
		SET @lv__ErrorText = @lv__ErrorText + ' (Severity: ' + CONVERT(VARCHAR(20),@lv__ErrorSeverity) + ') (State: ' + CONVERT(VARCHAR(20),@lv__ErrorState) + ')'

		RAISERROR( @lv__ErrorText, 11, 1);

		GOTO aftermatrix
	END CATCH


	--***Location 6: Get currently-running jobs
	--Now, we want to check for jobs that are currently running. (A job could have been running for the last 2 hours
	-- and it won't be in the #JobInstances table yet because the h.step_id=0 record doesn't exist until the job completes)
	
	--Note that we still need to do this even when @PointInTime is for many hours back, since a job could have been running VERY
	-- long amounts of time. 
	BEGIN TRY
		INSERT INTO #CurrentlyRunningJobs1 
			EXECUTE master.dbo.xp_sqlagent_enum_jobs 1, 'derp de derp';

		--We will need the Currently-running job info for the predictive matrix as well, so persist this info
		-- to an intermediate temp table
		INSERT INTO #CurrentlyRunningJobs2
			(native_job_id, JobStartTime, JobEndTime, JobDisplayEndTime)
		SELECT 
			ss.native_job_id, ss.JobStartTime, NULL, ss.JobDisplayEndTime
		FROM (
			SELECT
				[native_job_id] = ja.job_id, 
				[JobStartTime] = ja.start_execution_date, 
				[JobDisplayEndTime] = DATEADD(MINUTE, 1, @lv__mtx__OverallWindowEndTime),		--since a running job will, by definition, always have a '~' or '!' character in 
																	--the last matrix cell, we just push its end-time out just beyond the last time window
																	--note that this doesn't mess up our "historical average" calculation since we only look at
																	-- completed job instances for that calc
				rn = ROW_NUMBER() OVER (PARTITION BY ja.job_id ORDER BY ja.start_execution_date DESC)
			FROM msdb.dbo.sysjobactivity ja
			WHERE ja.start_execution_date IS NOT NULL
			AND ja.start_execution_date <= @lv__mtx__OverallWindowEndTime	--helps us handle @PointInTime values in the past
			AND ja.stop_execution_date IS NULL
		) ss
		WHERE ss.rn = 1			--if sysjobactivity has 2 or more NULL-stop records for the same job, we want the most recent one.

		--Since the sysjobactivity view can have records that actually refer to job instances that never finished (e.g. when SQL Agent was
		-- stopped suddenly), we need to cross-check the data with the results from xp_sqlagent_enum_jobs. Thus, sysjobactivity gets
		-- us the start time for a running job, and xp_sqlagent_enum_jobs gets us assurance that a job really is currently running.
		AND EXISTS (
			SELECT * 
			FROM #CurrentlyRunningJobs1 t
			WHERE t.Job_ID = ss.native_job_id
			AND t.Running = 1
		)
		;

		INSERT INTO #JobInstances
		(native_job_id, job_run_status, JobStartTime, JobEndTime, JobDisplayEndTime)
		SELECT native_job_id, 25, --special code that means "running"
			JobStartTime, JobEndTime, JobDisplayEndTime
		FROM #CurrentlyRunningJobs2;
	END TRY
	BEGIN CATCH
		RAISERROR(N'Error occurred while obtaining information about currently-running jobs. The job history matrices may be incomplete.', 11, 1);
		SELECT @lv__ErrorText = ERROR_MESSAGE(), 
				@lv__ErrorSeverity	= ERROR_SEVERITY(), 
				@lv__ErrorState = ERROR_STATE();
		SET @lv__ErrorText = @lv__ErrorText + ' (Severity: ' + CONVERT(VARCHAR(20),@lv__ErrorSeverity) + ') (State: ' + CONVERT(VARCHAR(20),@lv__ErrorState) + ')'

		RAISERROR( @lv__ErrorText, 11, 1);
	END CATCH

	SET @lv__afterdt = GETDATE();

	IF DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt) > @lv__slownessthreshold
	BEGIN
		SET @lv__ErrorText = N'   ***dbg: obtaining job completions and currently-running jobs took ' + CONVERT(NVARCHAR(40),DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt)) + N' milliseconds';
		RAISERROR(@lv__ErrorText, 10, 1) WITH NOWAIT;
	END

	SET @lv__beforedt = GETDATE();

	--***Location 7: Average job duration
	--Calculate the average duration for each job success (in seconds), and then apply that to the #JobInstances table so that
	-- down below we can determine whether a job instance has run longer than its average
	BEGIN TRY
		--Note that for @PointInTime values in the past, our average runtime calculation may be affected by jobs that succeeded
		-- AFTER @PointInTime. I've debated whether this is appropriate or not.
		-- TODO: reconsider using a rolling average in the future.
		UPDATE targ
		SET targ.JobRuns = ss1.JobRuns,
			targ.JobFailures = ss1.JobFailures,
			targ.MatrixNumber = CASE WHEN ss1.JobFailures > 0 OR ss1.IsCurrentlyRunning > 0 THEN 1 
									WHEN ss1.JobFailures = 0 AND ss1.JobRuns >= 1 THEN 3
									ELSE 5
								END,
			targ.CompletionsAllTime = ss1.CompletionsAllTime,
			targ.AvgJobDur_seconds = CASE WHEN ss1.CompletionsAllTime = 0 THEN 0 
										ELSE ss1.AllDuration / ss1.CompletionsAllTime END,
			targ.AvgSuccessDur_seconds = CASE WHEN ss1.CompletionsAllTime = 0 THEN 0 
										ELSE ss1.SuccessDuration / ss1.CompletionsAllTime END
		FROM #Jobs targ
			INNER JOIN (
			SELECT native_job_id,
				[CompletionsAllTime] = SUM(CompletionsAllTime),
				[SuccessDuration] = SUM(SuccessDuration),
				[AllDuration] = SUM(AllDuration),
				[JobFailures] = SUM(JobFailures),
				[IsCurrentlyRunning] = SUM(IsCurrentlyRunning), 
				[JobRuns] = SUM(JobRuns)
			FROM (
				SELECT 
					ji1.native_job_id,
					--the Count metrics are only supposed to reflect the time within our historical matrix window, and thus may
					-- include our currently-running jobs, while the average duration metrics are supposed to reflect all of our history, 
					-- but not our currently-running jobs (since they haven't finished yet)
					[CompletionsAllTime] = CASE WHEN ji1.JobEndTime IS NULL THEN 0 ELSE 1 END,

					[SuccessDuration] = CASE WHEN ji1.JobEndTime IS NULL OR ji1.job_run_status <> 1 
											THEN 0 ELSE DATEDIFF(SECOND, JobStartTime, JobEndTime) END,

					[AllDuration] = CASE WHEN ji1.JobEndTime IS NULL THEN 0 ELSE DATEDIFF(SECOND, JobStartTime, JobEndTime) END,

					[JobFailures] = CASE WHEN ji1.job_run_status IN (1,25) 
											AND ISNULL(JobEndTime,@lv__mtx__OverallWindowBeginTime) >= @lv__mtx__OverallWindowBeginTime 
											AND JobStartTime <= @lv__mtx__OverallWindowEndTime
										THEN 0 ELSE 1 END,

					[IsCurrentlyRunning] = CASE WHEN ji1.job_run_status = 25 
											AND JobStartTime <= @lv__mtx__OverallWindowEndTime
										THEN 1 ELSE 0 END,

					[JobRuns] = CASE WHEN ISNULL(JobEndTime,@lv__mtx__OverallWindowBeginTime) >= @lv__mtx__OverallWindowBeginTime 
										AND JobStartTime <= @lv__mtx__OverallWindowEndTime
									THEN 1 ELSE 0 END

				FROM #JobInstances ji1
			) ss0
			GROUP BY native_job_id
		) ss1
			ON targ.native_job_id = ss1.native_job_id
		;


		UPDATE ji1 
		--TODO: might want to make this configurable, where the user can choose whether to include failures in the 
		-- average duration info.
		SET JobExpectedEndTime = DATEADD(SECOND, j.AvgSuccessDur_seconds, ji1.JobStartTime) 
		FROM #JobInstances ji1
			INNER JOIN #Jobs j
				ON ji1.native_job_id = j.native_job_id
		;

		--Now that we have average duration, delete the job instances that we know don't matter anymore
		DELETE FROM #JobInstances 
		WHERE 
			--Any job that completed before our overall Window start time is irrelevant
			--Likewise, any job that started after our overall Window start time is also irrelevant
			--Note, however, that jobs that started before our Window can still be relevant if they ended
			-- AFTER our window started. 
			NOT (
				JobEndTime >= @lv__mtx__OverallWindowBeginTime 
				AND JobStartTime <= @lv__mtx__OverallWindowEndTime
				)
		;
	END TRY
	BEGIN CATCH
		RAISERROR(N'Error occurred while calculating average job runtime information and doing JI cleanup.', 11, 1);
		SET @output__DisplayMatrix = 0;
		SELECT @lv__ErrorText = ERROR_MESSAGE(), 
				@lv__ErrorSeverity	= ERROR_SEVERITY(), 
				@lv__ErrorState = ERROR_STATE();
		SET @lv__ErrorText = @lv__ErrorText + ' (Severity: ' + CONVERT(VARCHAR(20),@lv__ErrorSeverity) + ') (State: ' + CONVERT(VARCHAR(20),@lv__ErrorState) + ')'

		RAISERROR( @lv__ErrorText, 11, 1);

		GOTO aftermatrix
	END CATCH

	SET @lv__afterdt = GETDATE();

	IF DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt) > @lv__slownessthreshold
	BEGIN
		SET @lv__ErrorText = N'   ***dbg: calculating average durations and JI cleanup took ' + CONVERT(NVARCHAR(40),DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt)) + N' milliseconds';
		RAISERROR(@lv__ErrorText, 10, 1) WITH NOWAIT;
	END

	SET @lv__beforedt = GETDATE();

	--***Location 9: Now, create our Job Array.
	BEGIN TRY
	/*  logic up a few statements
		UPDATE targ 
		SET targ.JobRuns = ss.JobRuns, 
			targ.JobFailures = ss.JobFailures,
			targ.MatrixNumber = CASE WHEN ss.JobFailures > 0 OR ss.IsCurrentlyRunning > 0 THEN 1 
									WHEN ss.JobFailures = 0 AND ss.JobRuns >= 1 THEN 3
									ELSE 5
								END
		FROM #Jobs targ 
			INNER JOIN (
				SELECT jc.native_job_id, 
					SUM(CASE WHEN jc.job_run_status IN (1,25) THEN 0 ELSE 1 END) as JobFailures, 
					SUM(CASE WHEN jc.job_run_status = 25 THEN 1 ELSE 0 END) as IsCurrentlyRunning,
					SUM(1) as JobRuns
				FROM #JobInstances jc
				GROUP BY jc.native_job_id
				) ss
				ON targ.native_job_id = ss.native_job_id
		;
		*/
		
		INSERT INTO #JobArrays_Hist (JobID, WindowID, WindowBegin, WindowEnd, CellText)
		SELECT ss.JobID, tw.WindowID, tw.WindowBegin, tw.WindowEnd, @lv__mtx__EmptyChar
		FROM (
			SELECT j.JobID
			FROM #Jobs j
			) ss
			CROSS JOIN #TimeWindows_Hist tw
		;
	END TRY
	BEGIN CATCH
		RAISERROR(N'Error occurred while creating the job array. The job history matrix will not be displayed.', 11, 1);
		SET @output__DisplayMatrix = 0;
		SELECT @lv__ErrorText = ERROR_MESSAGE(), 
				@lv__ErrorSeverity	= ERROR_SEVERITY(), 
				@lv__ErrorState = ERROR_STATE();
		SET @lv__ErrorText = @lv__ErrorText + ' (Severity: ' + CONVERT(VARCHAR(20),@lv__ErrorSeverity) + ') (State: ' + CONVERT(VARCHAR(20),@lv__ErrorState) + ')'

		RAISERROR( @lv__ErrorText, 11, 1);

		GOTO aftermatrix
	END CATCH

	SET @lv__afterdt = GETDATE();

	IF DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt) > @lv__slownessthreshold
	BEGIN
		SET @lv__ErrorText = N'   ***dbg: counting job occurrences and constructing the Job Arrays took ' + CONVERT(NVARCHAR(40),DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt)) + N' milliseconds';
		RAISERROR(@lv__ErrorText, 10, 1) WITH NOWAIT;
	END

	SET @lv__beforedt = GETDATE();

	--***Location 10: Cell population for failures
	--Ok, first, update the arrays with any job failures. If a failure has occurred in a Time Window, then we mark that 
	-- time window with an 'X'
	BEGIN TRY 
		UPDATE targ
		SET targ.CellText = CASE WHEN xapp1.job_run_status = 0 THEN 'F'
								WHEN xapp1.job_run_status = 2 THEN 'R'
								WHEN xapp1.job_run_status = 3 THEN 'C'
							ELSE 'X'
							END
		FROM #JobArrays_Hist targ
			INNER JOIN #Jobs j
				ON targ.JobID = j.JobID
			CROSS APPLY (		--the use of CROSS rather than OUTER apply is important here. 
					SELECT TOP 1 jc.job_run_status 
					FROM #JobInstances jc
					WHERE j.native_job_id = jc.native_job_id
					AND jc.job_run_status <> 1
					AND jc.JobDisplayEndTime >= targ.WindowBegin
					AND jc.JobDisplayEndTime < targ.WindowEnd	--remember, endpoint of our window is NOT inclusive (because it is the same as the start of the subsequent window)
					ORDER BY jc.job_run_status ASC		--0 (failure) will sort first, 2 (retry) will sort second, and 3 (cancelled) will sort third
						--note that even if there are multiple jobs with the same run_status value, we don't really care, since we just pull the status
				) xapp1
		;
	END TRY
	BEGIN CATCH
		RAISERROR(N'Error occurred while populating Matrix with job failures. The job history matrix will not be displayed.', 11, 1);
		SET @output__DisplayMatrix = 0;
		SELECT @lv__ErrorText = ERROR_MESSAGE(), 
				@lv__ErrorSeverity	= ERROR_SEVERITY(), 
				@lv__ErrorState = ERROR_STATE();
		SET @lv__ErrorText = @lv__ErrorText + ' (Severity: ' + CONVERT(VARCHAR(20),@lv__ErrorSeverity) + ') (State: ' + CONVERT(VARCHAR(20),@lv__ErrorState) + ')'

		RAISERROR( @lv__ErrorText, 11, 1);

		GOTO aftermatrix
	END CATCH

	SET @lv__afterdt = GETDATE();

	IF DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt) > @lv__slownessthreshold
	BEGIN
		SET @lv__ErrorText = N'   ***dbg: populating Matrix with job failures took ' + CONVERT(NVARCHAR(40),DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt)) + N' milliseconds';
		RAISERROR(@lv__ErrorText, 10, 1) WITH NOWAIT;
	END

	SET @lv__beforedt = GETDATE();

	--***Location 11: Cell population for successes
	--Now, update the array with info on the # of SUCCESSFUL job completions in a given time block. 
	-- Note that we NEVER overwrite a cell that has already been written to.
	BEGIN TRY
		;WITH JobWindowsWithSuccesses AS (
			SELECT ja.JobID, ja.WindowID, 
				COUNT(*) AS NumSuccessfulCompletions
			FROM #JobArrays_Hist ja
				INNER JOIN #Jobs j
					ON ja.JobID = j.JobID
				INNER JOIN #JobInstances jc
					ON j.native_job_id = jc.native_job_id
			WHERE ja.CellText = @lv__mtx__EmptyChar
			AND jc.job_run_status = 1
			AND jc.JobDisplayEndTime >= ja.WindowBegin
			AND jc.JobDisplayEndTime < ja.WindowEnd	--remember, endpoint of our window is NOT inclusive (because it is the same as the start of the subsequent window)
			GROUP BY ja.JobID, ja.WindowID, ja.WindowBegin, ja.WindowEnd
		)
		UPDATE targ
		SET targ.CellText = CASE WHEN jw.NumSuccessfulCompletions >= 9 THEN '9'
								WHEN jw.NumSuccessfulCompletions = 1 THEN '/'
							 ELSE CONVERT(CHAR(1), jw.NumSuccessfulCompletions)
							 END
		FROM #JobArrays_Hist targ
			INNER JOIN JobWindowsWithSuccesses jw
				ON targ.JobID = jw.JobID
				AND targ.WindowID = jw.WindowID
		WHERE jw.NumSuccessfulCompletions > 0
		;
	END TRY
	BEGIN CATCH
		RAISERROR(N'Error occurred while populating Matrix with job successes. The job history matrix will not be displayed.', 11, 1);
		SET @output__DisplayMatrix = 0;
		SELECT @lv__ErrorText = ERROR_MESSAGE(), 
				@lv__ErrorSeverity	= ERROR_SEVERITY(), 
				@lv__ErrorState = ERROR_STATE();
		SET @lv__ErrorText = @lv__ErrorText + ' (Severity: ' + CONVERT(VARCHAR(20),@lv__ErrorSeverity) + ') (State: ' + CONVERT(VARCHAR(20),@lv__ErrorState) + ')'

		RAISERROR( @lv__ErrorText, 11, 1);

		GOTO aftermatrix
	END CATCH

	SET @lv__afterdt = GETDATE();

	IF DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt) > @lv__slownessthreshold
	BEGIN
		SET @lv__ErrorText = N'   ***dbg: populating Matrix with job successes took ' + CONVERT(NVARCHAR(40),DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt)) + N' milliseconds';
		RAISERROR(@lv__ErrorText, 10, 1) WITH NOWAIT;
	END

	SET @lv__beforedt = GETDATE();

	--***Location 12: Cell population for starts
	BEGIN TRY
		;WITH JobStarts AS (
			SELECT ja.JobID, ja.WindowID
			FROM #JobArrays_Hist ja
				INNER JOIN #Jobs j
					ON ja.JobID = j.JobID
			WHERE ja.CellText = @lv__mtx__EmptyChar
			AND EXISTS (SELECT * FROM #JobInstances jc
					WHERE j.native_job_id = jc.native_job_id
					AND jc.JobStartTime >= ja.WindowBegin
					AND jc.JobStartTime < ja.WindowEnd	--remember, endpoint of our window is NOT inclusive (because it is the same as the start of the subsequent window)
					)
		)
		UPDATE targ 
		SET targ.CellText = '^'
		FROM #JobArrays_Hist targ
			INNER JOIN JobStarts js
				ON targ.JobID = js.JobID
				AND targ.WindowID = js.WindowID
		;
	END TRY
	BEGIN CATCH
		RAISERROR(N'Error occurred while populating the Matrix with job starts. The job history matrix will not be displayed.', 11, 1);
		SET @output__DisplayMatrix = 0;
		SELECT @lv__ErrorText = ERROR_MESSAGE(), 
				@lv__ErrorSeverity	= ERROR_SEVERITY(), 
				@lv__ErrorState = ERROR_STATE();
		SET @lv__ErrorText = @lv__ErrorText + ' (Severity: ' + CONVERT(VARCHAR(20),@lv__ErrorSeverity) + ') (State: ' + CONVERT(VARCHAR(20),@lv__ErrorState) + ')'

		RAISERROR( @lv__ErrorText, 11, 1);

		GOTO aftermatrix
	END CATCH

	SET @lv__afterdt = GETDATE();

	IF DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt) > @lv__slownessthreshold
	BEGIN
		SET @lv__ErrorText = N'   ***dbg: populating Matrix with job starts took ' + CONVERT(NVARCHAR(40),DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt)) + N' milliseconds';
		RAISERROR(@lv__ErrorText, 10, 1) WITH NOWAIT;
	END

	SET @lv__beforedt = GETDATE();

	--***Location 13: Cell population for "running"
	--Ok, our final update to the array is to mark all cells with a '~' or '!' where a job was running during that window, but its start was before
	-- the time window started and its end is after the time window started.
	BEGIN TRY
		UPDATE ja
		SET CellText = CASE WHEN xapp1.JobExpectedEndTime < ja.WindowBegin THEN '!'
						ELSE '~'
						END
		FROM #JobArrays_Hist ja
			INNER JOIN #Jobs j
				ON ja.JobID = j.JobID
			CROSS APPLY (
					SELECT TOP 1	--there should only be 1 row anyway...
						ji.JobExpectedEndTime
					FROM #JobInstances ji
					WHERE ji.native_job_id = j.native_job_id
					AND ji.JobStartTime < ja.WindowBegin
					AND ji.JobDisplayEndTime >= ja.WindowEnd		--remember, WindowEnd is actually NOT inclusive
					) xapp1
		WHERE ja.CellText = @lv__mtx__EmptyChar
		;
	END TRY
	BEGIN CATCH
		RAISERROR(N'Error occurred while populating the Matrix with "job running" info. The job history matrix will not be displayed.', 11, 1);
		SET @output__DisplayMatrix = 0;
		SELECT @lv__ErrorText = ERROR_MESSAGE(), 
				@lv__ErrorSeverity	= ERROR_SEVERITY(), 
				@lv__ErrorState = ERROR_STATE();
		SET @lv__ErrorText = @lv__ErrorText + ' (Severity: ' + CONVERT(VARCHAR(20),@lv__ErrorSeverity) + ') (State: ' + CONVERT(VARCHAR(20),@lv__ErrorState) + ')'

		RAISERROR( @lv__ErrorText, 11, 1);

		GOTO aftermatrix
	END CATCH

	SET @lv__afterdt = GETDATE();

	IF DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt) > @lv__slownessthreshold
	BEGIN
		SET @lv__ErrorText = N'   ***dbg: populating Matrix with "running" tokens took ' + CONVERT(NVARCHAR(40),DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt)) + N' milliseconds';
		RAISERROR(@lv__ErrorText, 10, 1) WITH NOWAIT;
	END

	IF @Debug = 1
	BEGIN
		SELECT ja.JobID, ja.WindowID, ja.WindowBegin, ja.WindowEnd, ja.CellText
		FROM #JobArrays_Hist ja
		ORDER BY ja.JobID, ja.WindowID
	END

	SET @lv__beforedt = GETDATE();

	--***Location 14: Determine how many characters of the job name we'll be printing (potentially all)
	SELECT @lv__mtx__MaxJobNameLength = MAX(LEN(col1))
	FROM (
		SELECT col1 = (
			CASE WHEN j.IsEnabled = 0 THEN N'*' ELSE N'$' END +
				CONVERT(NVARCHAR(20),j.JobRuns) + N'/' + CONVERT(NVARCHAR(20),j.JobFailures) + N'  ' + 
				CASE WHEN j.CreateDate > @lv__mtx__OverallWindowEndTime
							THEN N'(' + j.JobName + N')'
						ELSE j.JobName
					END
			)
		FROM #Jobs j
	) ss0;

	--TODO: consider some way of customizing the max length under different circumstances (@ToConsole values, @FitOnScreen values, a user-specified param, etc)
	SET @lv__mtx__MaxJobNameLength = (
			CASE WHEN @lv__mtx__MaxJobNameLength IS NULL THEN 1		--no SQL Agent jobs exist! we shouldn't reach this point
				WHEN @lv__mtx__MaxJobNameLength <= 55 THEN @lv__mtx__MaxJobNameLength	--50 chars is fine whatever the output
				ELSE 55
			END 
			);


	--***Location 15: Construct the header lines
	SET @lv__mtx__HeaderHours = N'';
	SET @lv__mtx__HeaderLine = N'';

	SELECT @lv__mtx__HeaderHours = @lv__mtx__HeaderHours + tw.TimeHeaderChar
	FROM #TimeWindows_Hist tw
	ORDER BY tw.WindowID DESC;

	SELECT @lv__mtx__HeaderLine = @lv__mtx__HeaderLine + tw.LineHeaderChar
	FROM #TimeWindows_Hist tw
	ORDER BY tw.WindowID DESC

	--***Location 16: Creation of the output strings (before final concatenation, in sub-matrices)

	/* Our matrix is really several sub-matrices. Each matrix holds certain "categories" of jobs, based on those jobs' runs/failures/enabled/disabled status:

		For now, here's how we'll organize them:
		Matrix 1
			Jobs that have had a failure or are currently running  (whether disabled or not)
																				(use MatrixNumber=2 for a spacer line)
		Matrix 3
			Jobs not in Matrix 1 that have had at least 1 run
																				(use MatrixNumber=4 for a spacer line)
		Matrix 5
			All other jobs (jobs that haven't run at all, whether disabled or not)
	*/

	INSERT INTO #OutputStringLineItems_Hist 
		(RowType, JobID, MatrixNumber, DisplayOrder, CellString)
	SELECT 0,		-1,			-1,			1,			CASE WHEN @FitOnScreen = 'Y' THEN '' 
															ELSE REPLICATE('$', @lv__mtx__MaxJobNameLength)
															END + N'$' + @lv__mtx__HeaderHours
	UNION ALL 
	SELECT 1,		-1,			-1,			1,			CASE WHEN @FitOnScreen = 'Y' THEN '' 
															ELSE REPLICATE('$', @lv__mtx__MaxJobNameLength)
															END + N'$' + @lv__mtx__HeaderLine
	;

	SET @lv__afterdt = GETDATE();

	IF DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt) > @lv__slownessthreshold
	BEGIN
		SET @lv__ErrorText = N'   ***dbg: obtaining max job name length and header output lines took ' + CONVERT(NVARCHAR(40),DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt)) + N' milliseconds';
		RAISERROR(@lv__ErrorText, 10, 1) WITH NOWAIT;
	END

	SET @lv__beforedt = GETDATE();
	BEGIN TRY 
		INSERT INTO #OutputStringLineItems_Hist
		(RowType, JobID, CellString)
		SELECT 2,
			ja1.JobID, 
			CellString = (
						SELECT [*] = ja2.CellText
						FROM #JobArrays_Hist as ja2
						WHERE ja2.JobID = ja1.JobID
						ORDER BY ja2.WindowID DESC
						FOR XML PATH(N'')
					)
		FROM #JobArrays_Hist AS ja1
		GROUP BY ja1.JobID
		ORDER BY JobID;

		UPDATE targ 
		SET MatrixNumber = j.MatrixNumber, 
			DisplayOrder = j.DisplayOrder
		FROM #OutputStringLineItems_Hist targ
			INNER JOIN #Jobs j
				ON targ.JobID = j.JobID

		UPDATE targ  
		SET targ.CellString = (
				CASE 
					WHEN @FitOnScreen = 'Y' 
						THEN N'|' + targ.CellString + N'|' + ss.JobName
					ELSE SUBSTRING(ss.JobName,1,@lv__mtx__MaxJobNameLength) + N'|' + targ.CellString + N'|'
					END
			)
		FROM #OutputStringLineItems_Hist targ
			INNER JOIN (
					SELECT j.JobID, 
						[JobName] = CASE WHEN j.IsEnabled = 0 THEN N'*' ELSE N' ' END +
						CONVERT(NVARCHAR(20),j.JobRuns) + N'/' + CONVERT(NVARCHAR(20),j.JobFailures) + N'  ' + 
						CASE WHEN j.CreateDate > @lv__mtx__OverallWindowEndTime
								THEN N'(' + j.JobName + N')'
							ELSE j.JobName
						END + REPLICATE('$', 50)
					FROM #Jobs j
					) ss
				ON targ.JobID = ss.JobID
		;
	END TRY
	BEGIN CATCH
		RAISERROR(N'Error occurred while constructing output strings from job arrays. The job history matrix will not be displayed.', 11, 1);
		SET @output__DisplayMatrix = 0;
		SELECT @lv__ErrorText = ERROR_MESSAGE(), 
				@lv__ErrorSeverity	= ERROR_SEVERITY(), 
				@lv__ErrorState = ERROR_STATE();
		SET @lv__ErrorText = @lv__ErrorText + ' (Severity: ' + CONVERT(VARCHAR(20),@lv__ErrorSeverity) + ') (State: ' + CONVERT(VARCHAR(20),@lv__ErrorState) + ')'

		RAISERROR( @lv__ErrorText, 11, 1);

		GOTO aftermatrix
	END CATCH

	SET @lv__afterdt = GETDATE();

	IF DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt) > @lv__slownessthreshold
	BEGIN
		SET @lv__ErrorText = N'   ***dbg: constructing output strings from job arrays took ' + CONVERT(NVARCHAR(40),DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt)) + N' milliseconds';
		RAISERROR(@lv__ErrorText, 10, 1) WITH NOWAIT;
	END

	SET @lv__beforedt = GETDATE();

	--***Location 17: Sub-matrix labels
	--Matrices 1,3,5 hold actual rows, while 2,4,6 are labels indicating which matrix it is.
	-- The below labels are all intentionally the same length (39 chars)
	BEGIN TRY
		SET @lv__mtx__Replicate1 = (@lv__mtx__MatrixWidth - 39) / 2 + 1
		SET @lv__mtx__Replicate2 = @lv__mtx__MatrixWidth - 39 - @lv__mtx__Replicate1

		INSERT INTO #OutputStringLineItems_Hist
			(RowType, JobID, MatrixNumber, DisplayOrder, CellString)
		SELECT 2,		-1,		2,			1, 
			CASE WHEN @FitOnScreen = 'Y' THEN '' 
			ELSE REPLICATE('$', @lv__mtx__MaxJobNameLength)
			END + N'|' + REPLICATE(N'*', @lv__mtx__Replicate1) + N'Currently-running or at least 1 failure' + REPLICATE(N'*', @lv__mtx__Replicate2) + N'|' + NCHAR(10) + NCHAR(13) + NCHAR(10) + NCHAR(13)+ NCHAR(10) + NCHAR(13)
		UNION ALL
		SELECT 2,		-1,		4,			1, 
			CASE WHEN @FitOnScreen = 'Y' THEN '' 
			ELSE REPLICATE('$', @lv__mtx__MaxJobNameLength)
			END + N'|' + REPLICATE(N'*', @lv__mtx__Replicate1) + N'Executed >= 1 time and always succeeded' + REPLICATE(N'*', @lv__mtx__Replicate2) + N'|' + NCHAR(10) + NCHAR(13) + NCHAR(10) + NCHAR(13)+ NCHAR(10) + NCHAR(13)
		UNION ALL 
		SELECT 2,		-1,		6,			1,
			CASE WHEN @FitOnScreen = 'Y' THEN '' 
			ELSE REPLICATE('$', @lv__mtx__MaxJobNameLength)
			END + N'|' + REPLICATE(N'*', @lv__mtx__Replicate1) + N'Did not execute during this time window' + REPLICATE(N'*', @lv__mtx__Replicate2) + N'|' + NCHAR(10) + NCHAR(13) + NCHAR(10) + NCHAR(13)+ NCHAR(10) + NCHAR(13)
		;

		--***Location 18: Decide whether to re-print the header for second and third sub-matrices
		--If there are a lot of job rows in matrix #1 (Currently-running or at least 1 failure), then the user will have to scroll down to see the jobs in 
		-- matrix #3 (Executed >= 1 time, always succeeded). Similarly, if there are a lot of rows in matrix #3, then when the user scrolls down to see them, 
		-- the header rows will not be visible and the user will have to keep scrolling up and down to match times to matrix info. To avoid this, we 
		-- check the # of lines in matrices #1 and #3 combined, and if the result is > a threshold, we add header rows in to matrix 3 as well.
		SELECT 
			@lv__mtx__CountMatrixRows_1 = SUM(CASE WHEN o.MatrixNumber = 1 THEN 1 ELSE 0 END),
			@lv__mtx__CountMatrixRows_3 = SUM(CASE WHEN o.MatrixNumber = 3 THEN 1 ELSE 0 END),
			@lv__mtx__CountMatrixRows_5 = SUM(CASE WHEN o.MatrixNumber = 5 THEN 1 ELSE 0 END)
		FROM #OutputStringLineItems_Hist o


		SET @lv__mtx__Matrix3HasHeader = 0
		IF (@lv__mtx__CountMatrixRows_1 + @lv__mtx__CountMatrixRows_3) >= 35
		BEGIN
			INSERT INTO #OutputStringLineItems_Hist 
				(RowType, JobID, MatrixNumber, DisplayOrder, CellString)
			SELECT 0,		-1,			3,			1,			CASE WHEN @FitOnScreen = 'Y' THEN '' 
																ELSE REPLICATE('$', @lv__mtx__MaxJobNameLength)
																END + N'$' + @lv__mtx__HeaderHours
			UNION ALL 
			SELECT 1,		-1,			3,			1,			CASE WHEN @FitOnScreen = 'Y' THEN '' 
																ELSE REPLICATE('$', @lv__mtx__MaxJobNameLength)
																END + N'$' + @lv__mtx__HeaderLine
			;

			SET @lv__mtx__Matrix3HasHeader = 1;
		END

		--We need similar logic for Matrix 5
		IF (@lv__mtx__Matrix3HasHeader = 0 AND (@lv__mtx__CountMatrixRows_1 + @lv__mtx__CountMatrixRows_3 + @lv__mtx__CountMatrixRows_5) >= 30)
			OR (@lv__mtx__Matrix3HasHeader = 1 AND (@lv__mtx__CountMatrixRows_3 + @lv__mtx__CountMatrixRows_5) >= 35 )
		BEGIN
			INSERT INTO #OutputStringLineItems_Hist 
				(RowType, JobID, MatrixNumber, DisplayOrder, CellString)
			SELECT 0,		-1,			5,			1,			CASE WHEN @FitOnScreen = 'Y' THEN '' 
																ELSE REPLICATE('$', @lv__mtx__MaxJobNameLength)
																END + N'$' + @lv__mtx__HeaderHours
			UNION ALL 
			SELECT 1,		-1,			5,			1,			CASE WHEN @FitOnScreen = 'Y' THEN '' 
																ELSE REPLICATE('$', @lv__mtx__MaxJobNameLength)
																END + N'$' + @lv__mtx__HeaderLine
		END
	END TRY
	BEGIN CATCH
		RAISERROR(N'Error occurred while constructing sub-matrix headers. The job history matrix will not be displayed.', 11, 1);
		SET @output__DisplayMatrix = 0;
		SELECT @lv__ErrorText = ERROR_MESSAGE(), 
				@lv__ErrorSeverity	= ERROR_SEVERITY(), 
				@lv__ErrorState = ERROR_STATE();
		SET @lv__ErrorText = @lv__ErrorText + ' (Severity: ' + CONVERT(VARCHAR(20),@lv__ErrorSeverity) + ') (State: ' + CONVERT(VARCHAR(20),@lv__ErrorState) + ')'

		RAISERROR( @lv__ErrorText, 11, 1);

		GOTO aftermatrix
	END CATCH

	SET @lv__afterdt = GETDATE();

	IF DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt) > @lv__slownessthreshold
	BEGIN
		SET @lv__ErrorText = N'   ***dbg: Sub-matrix headers took ' + CONVERT(NVARCHAR(40),DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt)) + N' milliseconds';
		RAISERROR(@lv__ErrorText, 10, 1) WITH NOWAIT;
	END

	SET @lv__beforedt = GETDATE();

	--***Location 19: Ok, assemble final output
	SET @lv__OutputVar = N'';

	BEGIN TRY
		SELECT @lv__OutputVar = @lv__OutputVar + 
			REPLACE(REPLACE(CellString,N'_', N' '),N'$', N' ') + NCHAR(10) + 
			CASE WHEN RowType < 2 OR MatrixNumber IN (2,4,6)
				THEN N'' 
				ELSE (	N'' --Printing underscores as spacer lines proved to be uglier than just having each line follow consecutively
					/*
						CASE WHEN @Matrix_PrintUnderscores = N'Y' 
							THEN N'|' + REPLICATE(N'_', 156) + N'|' + NCHAR(10)
							ELSE N'' 
						END 
					*/
					)
			END 
		FROM #OutputStringLineItems_Hist
		ORDER BY MatrixNumber, RowType, DisplayOrder
		;

		SET @lv__mtx__PrintOverallBeginTime = CONVERT(VARCHAR(20),@lv__mtx__OverallWindowBeginTime,107) + N' ' + CONVERT(VARCHAR(20),@lv__mtx__OverallWindowBeginTime,108) + N'.' + CONVERT(VARCHAR(20),DATEPART(MILLISECOND, @lv__mtx__OverallWindowBeginTime))
		SET @lv__mtx__PrintOverallEndTime = CONVERT(VARCHAR(20),@lv__mtx__OverallWindowEndTime,107) + N' ' + CONVERT(VARCHAR(20),@lv__mtx__OverallWindowEndTime,108) + N'.' + CONVERT(VARCHAR(20),DATEPART(MILLISECOND, @lv__mtx__OverallWindowEndTime))

		SET @lv__mtx__Replicate1 = @lv__mtx__MatrixWidth - LEN(@lv__mtx__PrintOverallBeginTime) - LEN(@lv__mtx__PrintOverallEndTime);

		SET @lv__OutputVar = 
				CASE WHEN @ToConsole = N'N' THEN N'<?JobHistory -- ' + NCHAR(10)
					ELSE N'' END + 
				--HoursBack and cell minute width labels
				N'@HoursBack parameter value: ' + CONVERT(NVARCHAR(20),@HoursBack) + N'   Each matrix cell = ' + CONVERT(NVARCHAR(20),@lv__mtx__WindowLength_minutes) + N' minute(s)' + 

				--SQL/Agent Starttime labels
				CASE WHEN @lv__mtx__SQLServerStartTime IS NULL THEN N'***** WARNING: could not determine last SQL Server DB engine start time *****'
					 WHEN @lv__mtx__SQLServerStartTime = 2 THEN N'***** NOTE: SQL Server has started up at least once since the end time of this matrix *****'
					WHEN @lv__mtx__SQLServerStartTime BETWEEN @lv__mtx__OverallWindowBeginTime AND @lv__mtx__OverallWindowEndTime 
						THEN NCHAR(10) + N'***** SQL Server DB engine started at ' + CONVERT(VARCHAR(20),@lv__mtx__SQLServerStartTime,107) + N' ' + CONVERT(VARCHAR(20),@lv__mtx__SQLServerStartTime,108) + N'.' + CONVERT(VARCHAR(20),DATEPART(MILLISECOND, @lv__mtx__SQLServerStartTime)) + N' *****'
					ELSE N''
				END +

				CASE WHEN @lv__mtx__SQLAgentStartTime IS NULL THEN N'***** WARNING: could not determine last SQL Agent start time *****'
					WHEN @lv__mtx__SQLAgentStartTime = 2 THEN N'***** NOTE: Could not find the Agent start time immediately preceding this matrix *****'
					WHEN @lv__mtx__SQLAgentStartTime BETWEEN @lv__mtx__SQLAgentStartTime AND @lv__mtx__SQLAgentStartTime 
						AND ABS(DATEDIFF(MINUTE, @lv__mtx__SQLServerStartTime, @lv__mtx__SQLAgentStartTime)) > 1
						THEN NCHAR(10) + N'***** SQL Agent started at ' + CONVERT(VARCHAR(20),@lv__mtx__SQLAgentStartTime,107) + N' ' + CONVERT(VARCHAR(20),@lv__mtx__SQLAgentStartTime,108) + N'.' + CONVERT(VARCHAR(20),DATEPART(MILLISECOND, @lv__mtx__SQLAgentStartTime)) + N' *****'
					ELSE N''
				END + NCHAR(10) + NCHAR(13) +

				--Begin/End timestamp labels
				N' ' + CASE WHEN @FitOnScreen = 'N' THEN REPLICATE(' ', @lv__mtx__MaxJobNameLength) ELSE '' END + @lv__mtx__PrintOverallBeginTime + REPLICATE(N' ', @lv__mtx__Replicate1) + @lv__mtx__PrintOverallEndTime + N' ' + NCHAR(10) + 
			@lv__OutputVar + 
			CASE WHEN @ToConsole = N'N' THEN NCHAR(10) + NCHAR(13) + N'-- ?>'
				ELSE N'' END
			;
	END TRY
	BEGIN CATCH
		RAISERROR(N'Error occurred while constructing the final Matrix output string. The job history matrix will not be displayed.', 11, 1);
		SET @output__DisplayMatrix = 0;
		SELECT @lv__ErrorText = ERROR_MESSAGE(), 
				@lv__ErrorSeverity	= ERROR_SEVERITY(), 
				@lv__ErrorState = ERROR_STATE();
		SET @lv__ErrorText = @lv__ErrorText + ' (Severity: ' + CONVERT(VARCHAR(20),@lv__ErrorSeverity) + ') (State: ' + CONVERT(VARCHAR(20),@lv__ErrorState) + ')'

		RAISERROR( @lv__ErrorText, 11, 1);

		GOTO aftermatrix
	END CATCH

	SET @lv__afterdt = GETDATE();

	IF DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt) > @lv__slownessthreshold
	BEGIN
		SET @lv__ErrorText = N'   ***dbg: constructing final Matrix output took ' + CONVERT(NVARCHAR(40),DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt)) + N' milliseconds';
		RAISERROR(@lv__ErrorText, 10, 1) WITH NOWAIT;
	END

aftermatrix:


	/*********************************************************************************************************************************************
	*********************************************************************************************************************************************

														Part 2: Predictive Matrix

	*********************************************************************************************************************************************
	*********************************************************************************************************************************************/
	IF @output__DisplayPredictive = 0
	BEGIN
		GOTO configstart
	END

	/* The design goals for the predictive matrix are essentially the same as for the historical matrix. See the corresponding note
		near the top of the historical matrix section
		"
		...		The net effect of all this is that the time windows allowed for our matrix are:
				1, 2, 3, 5, 6, 10, 15, and 20

		The time window length also depends on whether the user wants everything in 1 screen, or is ok with scrolling.
		If @FitsOnScreen='Y', the matrix is kept in between 100 and 145 characters (with a couple exceptions on the lower side)
		and if ='N', then the matrix is kept to 360 maximum.
		"

		Here's the complete chart for the predictive matrix:

	If @FitOnScreen = 'Y'
		@HoursBack = 1 --> cell width = 1		MatrixWidth = 60		(60-wide doesn't look as good, so we actually bump it up to 1.5 hours)
		@HoursBack = 2 --> cell width = 1		MatrixWidth = 120
		@HoursBack = 3 --> cell width = 2		MatrixWidth = 90
		@HoursBack = 4 --> cell width = 2		MatrixWidth = 120
		@HoursBack = 5 --> cell width = 3/2		MatrixWidth = 100/150	(100 width for XML, 150 width for console)
		@HoursBack = 6 --> cell width = 3		MatrixWidth = 120
		@HoursBack = 7 --> cell width = 3		MatrixWidth = 140
		@HoursBack = 8 --> cell width = 5		MatrixWidth = 96	
		@HoursBack = 9 --> cell width = 5		MatrixWidth = 108	
		@HoursBack =10 --> cell width = 5		MatrixWidth = 120
		@HoursBack =11 --> cell width = 5		MatrixWidth = 132
		@HoursBack =12 --> cell width = 5		MatrixWidth = 144
		@HoursBack =13 --> cell width = 6		MatrixWidth = 130
		@HoursBack =14 --> cell width = 6		MatrixWidth = 140
		@HoursBack =15 --> cell width = 10		MatrixWidth = 90	
		@HoursBack =16 --> cell width = 10		MatrixWidth = 96	
		@HoursBack =17 --> cell width = 10		MatrixWidth = 102	
		@HoursBack =18 --> cell width = 10		MatrixWidth = 108	
		@HoursBack =19 --> cell width = 10		MatrixWidth = 114	
		@HoursBack =20 --> cell width = 10		MatrixWidth = 120

	If @FitOnScreen = 'N'

		@HoursBack = 1 --> cell width = 1		MatrixWidth = 60	
		@HoursBack = 2 --> cell width = 1		MatrixWidth = 120
		@HoursBack = 3 --> cell width = 1		MatrixWidth = 180
		@HoursBack = 4 --> cell width = 1		MatrixWidth = 240
		@HoursBack = 5 --> cell width = 1		MatrixWidth = 300
		@HoursBack = 6 --> cell width = 1		MatrixWidth = 360

		@HoursBack = 7 --> cell width = 2		MatrixWidth = 210
		@HoursBack = 8 --> cell width = 2		MatrixWidth = 240
		@HoursBack = 9 --> cell width = 2		MatrixWidth = 270
		@HoursBack =10 --> cell width = 2		MatrixWidth = 300
		@HoursBack =11 --> cell width = 2		MatrixWidth = 330
		@HoursBack =12 --> cell width = 2		MatrixWidth = 360

		--Let's skip cell widths = 3 and 4 because they don't line up with tick marks as well

		@HoursBack =13 --> cell width = 5		MatrixWidth = 156
		@HoursBack =14 --> cell width = 5		MatrixWidth = 168
		@HoursBack =15 --> cell width = 5		MatrixWidth = 180
		@HoursBack =16 --> cell width = 5		MatrixWidth = 192
		@HoursBack =17 --> cell width = 5		MatrixWidth = 204
		@HoursBack =18 --> cell width = 5		MatrixWidth = 216

		@HoursBack =18 --> cell width = 5		MatrixWidth = 216
		@HoursBack =18 --> cell width = 5		MatrixWidth = 216
	*/

	IF @FitOnScreen = N'Y'
	BEGIN
		SELECT @lv__pred__WindowLength_minutes = CASE 
				WHEN @HoursForward BETWEEN 1 AND 2 THEN 1
				WHEN @HoursForward BETWEEN 3 AND 4 THEN 2
				WHEN @HoursForward = 5
					THEN (
						CASE WHEN @ToConsole = N'N' THEN 3
							ELSE 2
						END
					)
				WHEN @HoursForward BETWEEN 6 AND 7 THEN 3
				WHEN @HoursForward BETWEEN 8 AND 12 THEN 5
				WHEN @HoursForward BETWEEN 13 AND 14 THEN 6
				WHEN @HoursForward BETWEEN 15 AND 20 THEN 10
			ELSE 1	--shouldn't hit this
			END;
	END
	ELSE
	BEGIN
		SELECT @lv__pred__WindowLength_minutes = CASE 
				WHEN @HoursForward BETWEEN 1 AND 6 THEN 1
				WHEN @HoursForward BETWEEN 7 AND 12 THEN 2
				WHEN @HoursForward BETWEEN 13 AND 20 THEN 5
			ELSE 1 --shouldn't hit this
			END;
	END

	--For @HoursForward=1, since our minimum cell width is 1 minute, we only end up with a 60-char wide matrix, and only 1
	-- hour-marker in the header. That doesn't look as good, so let's bump up the size of the matrix (and the time window)
	-- by 30
	IF @HoursForward = 1
	BEGIN
		SET @lv__pred__MatrixWidth = 90;
	END
	ELSE
	BEGIN
		--Matrix width is easy to calculate once we have window length
		SET @lv__pred__MatrixWidth = @HoursForward*60 / @lv__pred__WindowLength_minutes;
	END

	--For the "Time Header" line, we want to mark inter-hour "landmarks" to make rough time identification easier.
	--We also want to do something similar for the "Line Header" line
	IF @lv__pred__WindowLength_minutes IN (1,2)
	BEGIN
		SET @lv__pred__LineHeaderMod = 10;		--print ticks every 10 minutes
		SET @lv__pred__TimeHeaderMod = 20;		--print '+' chars every 20 min, but not on the hour
	END
	ELSE IF @lv__pred__WindowLength_minutes IN (3,5)
	BEGIN
		SET @lv__pred__LineHeaderMod = 15;
		SET @lv__pred__TimeHeaderMod = 30;
	END
	ELSE IF @lv__pred__WindowLength_minutes IN (6,10,15)
	BEGIN
		SET @lv__pred__LineHeaderMod = 30;
		SET @lv__pred__TimeHeaderMod = 30;
	END
	ELSE 
		--IF @lv__pred__WindowLength_minutes = 20		the only other option at this time is 20
	BEGIN
		SET @lv__pred__LineHeaderMod = -1;
		SET @lv__pred__TimeHeaderMod = -1;
	END

	--Because (n)varchar strings are trimmed, we use an underscore for most of the string manipulation and then 
	--do a REPLACE(<expr>, @lv__mtx__EmptyChar, N' ') at the end.
	SET @lv__pred__EmptyChar = N'_';

	--***Location 3: Determine first window of matrix
	--The @PointInTime is very likely NOT the exact begin point for an x-minute time window. Let's find the begin point for the
	-- time window that we are in currently.
	BEGIN TRY
		SELECT 
				@lv__pred__CurrentTime_WindowBegin = ss3.CurrentTime_WindowBegin,
				@lv__pred__CurrentTime_WindowEnd = DATEADD(MINUTE, @lv__pred__WindowLength_minutes, CurrentTime_WindowBegin)
			FROM (
				SELECT [CurrentTime_WindowBegin] = DATEADD(MINUTE, NthWindowFromTopOfHour*@lv__pred__WindowLength_minutes, CurrentTime_HourBase), 
					CurrentTime, 
					CurrentTime_HourBase, 
					CurrentMinute, 
					CurrentHour, 
					NthWindowFromTopOfHour
				FROM (
					SELECT [NthWindowFromTopOfHour] = CurrentMinute / @lv__pred__WindowLength_minutes,		--zero-based, of course
						[CurrentTime_HourBase] = DATEADD(HOUR, CurrentHour, 
																CONVERT(DATETIME,
																		CONVERT(VARCHAR(20), CurrentTime, 101)
																		)
														),
						CurrentTime, 
						CurrentMinute, 
						CurrentHour 
					FROM (
						SELECT [CurrentMinute] = DATEPART(MINUTE, CurrentTime), 
							[CurrentHour] = DATEPART(HOUR, CurrentTime),
							CurrentTime
						FROM 
							(SELECT [CurrentTime] = @PointInTime) ss0
						) ss1
					) ss2
			) ss3
			;
	END TRY
	BEGIN CATCH
		RAISERROR(N'Unable to construct the initial time window. The job predictive matrix will not be displayed.', 11, 1);
		SET @output__DisplayPredictive = 0;
		SELECT @lv__ErrorText = ERROR_MESSAGE(), 
				@lv__ErrorSeverity	= ERROR_SEVERITY(), 
				@lv__ErrorState = ERROR_STATE();
		SET @lv__ErrorText = @lv__ErrorText + ' (Severity: ' + CONVERT(VARCHAR(20),@lv__ErrorSeverity) + ') (State: ' + CONVERT(VARCHAR(20),@lv__ErrorState) + ')'

		RAISERROR( @lv__ErrorText, 11, 1);

		GOTO afterpredict
	END CATCH

	----***Location 4: Now build our array of time windows for the whole matrix
	BEGIN TRY
		;WITH t0 AS (
			SELECT 0 as col1 UNION ALL
			SELECT 0 UNION ALL
			SELECT 0 UNION ALL
			SELECT 0
		),
		t1 AS (
			SELECT ref1.col1 FROM t0 as ref1
				CROSS JOIN t0 as ref2
				CROSS JOIN t0 as ref3
				CROSS JOIN t0 as ref4
		),
		nums AS (
			SELECT ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) as rn
			FROM t1
		)
		INSERT INTO #TimeWindows_Pred (WindowID, WindowBegin, WindowEnd, TimeHeaderChar, LineHeaderChar)
		SELECT 
			CellReverseOrder, 
			WindowBegin,
			WindowEnd,
			TimeHeaderChar = (
				CASE 
					--When we are on the top of the hour, we usually print hour information (first digit of a 2-digit hour)
					WHEN NthWindowFromTopOfHour = 0
						THEN (CASE 
								--For @HoursForward>24, print even hours. Otherwise, always print the hour
								WHEN @HoursForward > 24 
									THEN (
										CASE WHEN DATEPART(HOUR, WindowBegin) % 2 = 0 THEN SUBSTRING(CONVERT(NVARCHAR(20),DATEPART(HOUR, WindowBegin)),1,1)
											ELSE N'.'
										END 
									)
								ELSE SUBSTRING(CONVERT(NVARCHAR(20),DATEPART(HOUR, WindowBegin)),1,1)
								END
								)
					--When it is the second window of the hour, we check to see if we have a double-digit hour # and print the second digit
					WHEN NthWindowFromTopOfHour = 1
						THEN (
							CASE WHEN DATEPART(HOUR, WindowBegin) < 10 THEN N'.'
								ELSE SUBSTRING(REVERSE(CONVERT(NVARCHAR(20),DATEPART(HOUR, WindowBegin))),1,1)
							END
						)
					--should we print the Time Header intra-hour marker?
					WHEN @lv__pred__TimeHeaderMod <> -1 AND DATEPART(MINUTE, WindowBegin) % @lv__pred__TimeHeaderMod = 0
						THEN (CASE 
								WHEN @HoursForward > 24 THEN '.'		--too high-level for intra-hour markers
								ELSE '+'
							END
						)
					ELSE '.'	--should never hit this case
				END 
				),
			LineHeaderChar = (
				CASE 
					WHEN DATEPART(MINUTE, WindowBegin) % @lv__pred__LineHeaderMod = 0 THEN '|' 
					ELSE '-'
				END 
				)
		FROM (
			SELECT 
				CellReverseOrder, 
				CurrentTime_WindowBegin, 
				CurrentTime_WindowEnd, 
				WindowBegin, 
				WindowEnd,
				[NthWindowFromTopOfHour] = DATEPART(MINUTE, WindowBegin)  / @lv__pred__WindowLength_minutes
			FROM (
				SELECT TOP (@lv__pred__MatrixWidth) 
					rn as CellReverseOrder,
					@lv__pred__CurrentTime_WindowBegin as CurrentTime_WindowBegin, 
					@lv__pred__CurrentTime_WindowEnd as CurrentTime_WindowEnd,
					DATEADD(MINUTE, @lv__pred__WindowLength_minutes*(rn-1), @lv__pred__CurrentTime_WindowBegin) as WindowBegin,
					DATEADD(MINUTE, @lv__pred__WindowLength_minutes*(rn-1), @lv__pred__CurrentTime_WindowEnd) as WindowEnd
				FROM nums 
				ORDER BY rn ASC
			) ss0
		) ss1
		OPTION(MAXDOP 1);
	END TRY
	BEGIN CATCH
		RAISERROR(N'Unable to define the complete list of time window boundaries. The job predictive matrix will not be displayed.', 11, 1);
		SET @output__DisplayPredictive = 0;
		SELECT @lv__ErrorText = ERROR_MESSAGE(), 
				@lv__ErrorSeverity	= ERROR_SEVERITY(), 
				@lv__ErrorState = ERROR_STATE();
		SET @lv__ErrorText = @lv__ErrorText + ' (Severity: ' + CONVERT(VARCHAR(20),@lv__ErrorSeverity) + ') (State: ' + CONVERT(VARCHAR(20),@lv__ErrorState) + ')'

		RAISERROR( @lv__ErrorText, 11, 1);

		GOTO afterpredict
	END CATCH

	
	IF @Debug = 1
	BEGIN
		SELECT 'Contents of #TimeWindows_Pred' as DebugLocation, tw.WindowID, tw.WindowBegin, tw.WindowEnd, tw.TimeHeaderChar, tw.LineHeaderChar
		FROM #TimeWindows_Pred tw
		ORDER BY tw.WindowID;
	END

	--Get overall min/max times, as we'll use these later in the proc
	SELECT 
		@lv__pred__OverallWindowBeginTime = MIN(tw.WindowBegin), 
		@lv__pred__OverallWindowEndTime = MAX(tw.WindowEnd)
	FROM #TimeWindows_Hist tw;
	
	SET @lv__afterdt = GETDATE();

	IF DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt) > @lv__slownessthreshold
	BEGIN
		SET @lv__ErrorText = N'   ***dbg: constructing Predictive Matrix time windows took ' + CONVERT(NVARCHAR(40),DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt)) + N' milliseconds';
		RAISERROR(@lv__ErrorText, 10, 1) WITH NOWAIT;
	END

	IF @Debug = 1
	BEGIN
		SELECT @lv__pred__OverallWindowBeginTime as OverallWindowBeginTime, @lv__pred__OverallWindowEndTime as OverallWindowEndTime;
	END


	--We are now ready to start hypothetically "running" jobs. This means taking the average job duration and "walking forward" through
	-- the schedule until we hit our matrix end time. Note that we always assume a job will succeed, even if it has only failures in its history.

	--Remember to take into account currently-running jobs.
	DECLARE @lv__pred__cursor__HypotheticalCurrentTime DATETIME,
			@lv__pred__cursor__KeepOuterLooping NCHAR(1),
			@lv__pred__cursor__KeepInnerLooping NCHAR(1),
			@lv__pred__cursor__KeepScheduleLooping NCHAR(1),
			@lv__pred__cursor__NativeJobID UNIQUEIDENTIFIER,
			@lv__pred__cursor__JobName VARCHAR(500),
			@lv__pred__cursor__PreviousJobName VARCHAR(500),
			@lv__pred__cursor__AvgJobDur_seconds BIGINT,
			@lv__pred__cursor__AvgSuccessDur_seconds BIGINT,
			@lv__pred__cursor__LastStartTime DATETIME,
			@lv__pred__cursor__NextScheduledTime DATETIME,
			@lv__pred__cursor__HypoStartTime DATETIME,
			@lv__pred__cursor__HypoEndTime DATETIME,
			@lv__pred__cursor__HypoLastStartTime DATETIME,
			@lv__pred__cursor__tmptime datetime,
			@lv__pred__cursor__tmptime2 datetime,
			@lv__pred__cursor__tmptime3 datetime
			;

	DECLARE @lv__pred__schedule__freq_type smallint, 
		 @lv__pred__schedule__freq_interval smallint,
		 @lv__pred__schedule__freq_subday_type smallint,
		 @lv__pred__schedule__freq_subday_interval smallint,
		 @lv__pred__schedule__freq_relative_interval smallint,
		 @lv__pred__schedule__freq_recurrence_factor smallint,
		 @lv__pred__schedule__active_start_date int,
		 @lv__pred__schedule__active_end_date int,
		 @lv__pred__schedule__active_start_time int, 
		 @lv__pred__schedule__active_end_time int,
		 @lv__pred__schedule__ActiveStartTime datetime, 
		 @lv__pred__schedule__ActiveEndtime datetime,
		 @lv__pred__schedule__DailyStartTime datetime, 
		 @lv__pred__schedule__DailyEndtime datetime,
		 @lv__pred__schedule__MinNextTime DATETIME,
		 @lv__pred__schedule__NextScheduleTimeFromSystem DATETIME,
		 @lv__pred__schedule__lotime datetime, 
		 @lv__pred__schedule__hitime datetime, 
		 @lv__pred__schedule__intvdiff int
		 ;

	DECLARE @debugtime datetime;
	set @debugtime = dateadd(minute, 5, getdate());
	
	SET @lv__pred__cursor__PreviousJobName = '<Loop not yet started>';
	SET @lv__pred__cursor__KeepOuterLooping = N'Y';
	
	DECLARE CreateHypotheticalRuns CURSOR FOR
	SELECT native_job_id, 
		j.JobName,
		AvgJobDur_seconds, 
		AvgSuccessDur_seconds,
		xapp1.start_execution_date,
		xapp1.next_scheduled_run_date
	FROM #Jobs j
		OUTER APPLY (
			SELECT TOP 1 
				ja.next_scheduled_run_date, ja.start_execution_date
			FROM msdb.dbo.sysjobactivity ja
			WHERE ja.job_id = j.native_job_id
			ORDER BY ja.session_id desc, ja.next_scheduled_run_date desc
			) xapp1
	WHERE j.IsEnabled = 1
	ORDER BY JobName ASC;

	OPEN CreateHypotheticalRuns

	WHILE @lv__pred__cursor__KeepOuterLooping = N'Y'
	BEGIN
		FETCH CreateHypotheticalRuns INTO 
			@lv__pred__cursor__NativeJobID,
			@lv__pred__cursor__JobName,
			@lv__pred__cursor__AvgJobDur_seconds,
			@lv__pred__cursor__AvgSuccessDur_seconds,
			@lv__pred__cursor__LastStartTime,
			@lv__pred__cursor__NextScheduledTime;

		IF @@FETCH_STATUS <> 0
		BEGIN
			SET @lv__pred__cursor__KeepOuterLooping = N'N';
			CONTINUE;
		END

		SET @lv__pred__cursor__HypotheticalCurrentTime = GETDATE();
		--even if @PointInTime is in the future, we still need to start
		--from the present because when we arrive at @PointInTime we 
		--need to know which jobs are "currently running".

		/*
		SELECT @lv__pred__cursor__NativeJobID as JobID,
			@lv__pred__cursor__JobName as JobName,
			@lv__pred__cursor__AvgJobDur_seconds as AvgDur_seconds,
			@lv__pred__cursor__AvgSuccessDur_seconds as AvgSuccessDur_seconds,
			@lv__pred__cursor__LastStartTime as LastStartTime,
			@lv__pred__cursor__NextScheduledTime as NextScheduledTime
		*/

		--TODO: insert a hypothetical row for currently-running jobs and set the
		-- hypothetical time to the end of that job's execution
		--also make sure to handle the @lv__pred__cursor__LastStartTime variable correctly
		SET @lv__pred__cursor__HypoStartTime = NULL;
		SET @lv__pred__cursor__HypoEndTime = NULL;

		SELECT 
			@lv__pred__cursor__HypoStartTime = ss.JobStartTime,
			@lv__pred__cursor__HypoEndTime = ss.JobEndTime
		FROM (
			SELECT TOP 1 
				c.JobStartTime, 
				JobEndTime = DATEADD(SECOND, @lv__pred__cursor__AvgSuccessDur_seconds, c.JobStartTime)
			FROM #CurrentlyRunningJobs2 c
			WHERE c.native_job_id = @lv__pred__cursor__NativeJobID
			ORDER BY c.JobStartTime DESC
		) ss;

		IF @lv__pred__cursor__HypoStartTime IS NOT NULL AND @lv__pred__cursor__HypoEndTime IS NOT NULL
		BEGIN
			INSERT INTO #HypotheticalRuns 
			(native_job_id, JobStartTime, JobExpectedEndTime)
			SELECT @lv__pred__cursor__NativeJobID, @lv__pred__cursor__HypoStartTime, @lv__pred__cursor__HypoEndTime;

			--Set our starting point to the time when the currently-running job is expected to end.
			SET @lv__pred__cursor__HypotheticalCurrentTime = DATEADD(SECOND, 1, @lv__pred__cursor__HypoEndTime);

			--TODO: do we need to track "last start time" or "last completion time"? if so, do that here
		END
		ELSE
		BEGIN
			--the job is not currently-running. 
			SET @lv__pred__cursor__HypotheticalCurrentTime = GETDATE();
		END


		TRUNCATE TABLE #SchedulesForThisJob;

		INSERT INTO #SchedulesForThisJob (
			freq_type,
			freq_interval,
			freq_subday_type,
			freq_subday_interval,
			freq_relative_interval,
			freq_recurrence_factor,
			active_start_date,
			active_end_date,
			active_start_time,
			active_end_time,
			ActiveStartTime,
			ActiveEndtime,
			NextScheduleTimeFromSystem
		)
		select 
			--ss.schedule_id, 
			--ss.name as ScheduleName, 
			--ss.enabled, 
			ss.freq_type, 
			ss.freq_interval, 
			ss.freq_subday_type,
			ss.freq_subday_interval,
			ss.freq_relative_interval,
			ss.freq_recurrence_factor,
			ss.active_start_date, 
			ss.active_end_date,
			ss.active_start_time,
			ss.active_end_time,
			ActiveStartTime = (
					CASE WHEN (ss.active_start_date IS NULL OR ss.active_start_time IS NULL )
							THEN GETDATE()
						ELSE CAST(STR(ss.active_start_date, 8, 0) AS DATETIME) + 
							CAST(STUFF(STUFF(REPLACE(STR(ss.active_start_time, 6), ' ', '0'), 3, 0, ':'), 6, 0, ':') AS DATETIME)
						END),
			ActiveEndTime = (
					CASE WHEN (ss.active_end_date IS NULL OR ss.active_end_time IS NULL )
							THEN GETDATE()
						ELSE CAST(STR(ss.active_end_date, 8, 0) AS DATETIME) + 
							CAST(STUFF(STUFF(REPLACE(STR(ss.active_end_time, 6), ' ', '0'), 3, 0, ':'), 6, 0, ':') AS DATETIME)
						END),
			NextScheduleTimeFromSystem = (
						CASE WHEN ( ISNULL(js.next_run_date,0) = 0 OR ISNULL(js.next_run_time,0) = 0 )
							THEN NULL
						ELSE CAST(STR(js.next_run_date, 8, 0) AS DATETIME) + 
							CAST(STUFF(STUFF(REPLACE(STR(js.next_run_time, 6), ' ', '0'), 3, 0, ':'), 6, 0, ':') AS DATETIME)
						END)
		from msdb.dbo.sysjobs j
			inner join msdb.dbo.sysjobschedules js
				on js.job_id = j.job_id
			inner join msdb.dbo.sysschedules ss
				on js.schedule_id = ss.schedule_id
		where 1=1
		and j.job_id = @lv__pred__cursor__NativeJobID
	--	debug: omitting this and ss.enabled = 1
		and ss.freq_type < 64		--64=SQL Agent startup, and 128 is CPU idle; can't predict that so exclude those schedules
		;

		--print @@ROWCOUNT


		SET @lv__pred__cursor__KeepInnerLooping = N'Y'

		--Until our hypothetical time exceeds the predictive-matrix end time, we need
		-- to continue generating "hypothetical" runs for the job and moving the hypothetical time forward. 
		WHILE @lv__pred__cursor__KeepInnerLooping = N'Y'
			AND @lv__pred__cursor__HypotheticalCurrentTime <= @lv__pred__OverallWindowEndTime
			--debug: 
			AND @lv__pred__cursor__HypotheticalCurrentTime <= @debugtime
		BEGIN
			TRUNCATE TABLE #ScheduleNextExecTimes;

			DECLARE PredictIterationSchedules CURSOR FOR
			SELECT freq_type,
				freq_interval,
				freq_subday_type,
				freq_subday_interval,
				freq_relative_interval,
				freq_recurrence_factor,
				active_start_date,
				active_end_date,
				active_start_time,
				active_end_time,
				ActiveStartTime,
				ActiveEndtime,
				NextScheduleTimeFromSystem
			FROM #SchedulesForThisJob;

			SET @lv__pred__cursor__KeepScheduleLooping = N'Y';

			OPEN PredictIterationSchedules;

			WHILE @lv__pred__cursor__KeepScheduleLooping = N'Y'
			BEGIN
				FETCH PredictIterationSchedules INTO 
					@lv__pred__schedule__freq_type,
					 @lv__pred__schedule__freq_interval,
					 @lv__pred__schedule__freq_subday_type,
					 @lv__pred__schedule__freq_subday_interval,
					 @lv__pred__schedule__freq_relative_interval,
					 @lv__pred__schedule__freq_recurrence_factor,
					 @lv__pred__schedule__active_start_date,
					 @lv__pred__schedule__active_end_date,
					 @lv__pred__schedule__active_start_time,
					 @lv__pred__schedule__active_end_time,
					 @lv__pred__schedule__ActiveStartTime,
					 @lv__pred__schedule__ActiveEndtime,
					 @lv__pred__schedule__NextScheduleTimeFromSystem

				IF @@FETCH_STATUS <> 0
				BEGIN
					SET @lv__pred__cursor__KeepScheduleLooping = N'N';
					CONTINUE;
				END

				/*
				
				select @lv__pred__cursor__JobName,@lv__pred__schedule__freq_type,
					 @lv__pred__schedule__freq_interval,
					 @lv__pred__schedule__freq_subday_type,
					 @lv__pred__schedule__freq_subday_interval,
					 @lv__pred__schedule__freq_relative_interval,
					 @lv__pred__schedule__freq_recurrence_factor,
					 @lv__pred__schedule__active_start_date,
					 @lv__pred__schedule__active_end_date,
					 @lv__pred__schedule__active_start_time,
					 @lv__pred__schedule__active_end_time,
					 @lv__pred__schedule__ActiveStartTime,
					 @lv__pred__schedule__ActiveEndtime,
					 @lv__pred__schedule__DailyStartTime,
					 @lv__pred__schedule__DailyEndTime
					 */

				--Our Daily Start and Daily End times always refer to the same day as our Hypothetical time does, but are based on 
				-- the Active Start and Active End times of the schedule, and thus need to be reset each time through the loop
				SET @lv__pred__schedule__DailyStartTime = CONVERT(datetime,CONVERT(varchar(20), @lv__pred__cursor__HypotheticalCurrentTime,101)) + 
															CAST(STUFF(STUFF(REPLACE(STR(@lv__pred__schedule__active_start_time, 6), ' ', '0'), 3, 0, ':'), 6, 0, ':') AS DATETIME);
				SET @lv__pred__schedule__DailyEndTime = CONVERT(datetime,CONVERT(varchar(20), @lv__pred__cursor__HypotheticalCurrentTime,101)) + 
															CAST(STUFF(STUFF(REPLACE(STR(@lv__pred__schedule__active_end_time, 6), ' ', '0'), 3, 0, ':'), 6, 0, ':') AS DATETIME);

				--One-time jobs are easy... just look at the ActiveStartTime value and if it is in the future, add it
				IF @lv__pred__schedule__freq_type = 1
				BEGIN
					IF @lv__pred__schedule__ActiveStartTime >= @lv__pred__cursor__HypotheticalCurrentTime
					BEGIN
						INSERT INTO #ScheduleNextExecTimes (ExecStartTime)
						SELECT @lv__pred__schedule__ActiveStartTime;
					END
				END


				--Daily-frequency
				IF @lv__pred__schedule__freq_type = 4
				BEGIN
					--@lv__pred__schedule__freq_interval means "every X days". The base for this calculation is ActiveStartTime, 
					-- so let's figure out the next day equal-to-or-after Hypo time that this schedule should run.

					SET @lv__pred__schedule__intvdiff = DATEDIFF(DAY, @lv__pred__schedule__ActiveStartTime, 
															@lv__pred__cursor__HypotheticalCurrentTime) / @lv__pred__schedule__freq_interval;

					SET @lv__pred__schedule__lotime = DATEADD(DAY, @lv__pred__schedule__intvdiff*@lv__pred__schedule__freq_interval, 
														@lv__pred__schedule__ActiveStartTime);

					SET @lv__pred__schedule__hitime = DATEADD(DAY, (@lv__pred__schedule__intvdiff+1)*@lv__pred__schedule__freq_interval, 
																@lv__pred__schedule__ActiveStartTime);

					--Note that (ignoring time portions, just focusing on the day portion), @lv__pred__cursor__HypotheticalCurrentTime is 
					-- always the same day or later as @lotime, and always LESS than @hitime (never equal)
					IF DATEDIFF(DAY, @lv__pred__schedule__lotime, @lv__pred__cursor__HypotheticalCurrentTime) > 0
					BEGIN	--hypo is on a later day than @lotime. Thus, our next execution is @hitime's StartTime value,
							--regardless of whether the schedule type is once-a-day or recurring-through-a-time-window-in-a-day
						--Of course, we need to make sure that @hitime isn't beyond the end of the schedule's expiration date
						IF @lv__pred__schedule__hitime <= @lv__pred__schedule__ActiveEndtime
						BEGIN
							INSERT INTO #ScheduleNextExecTimes (ExecStartTime)
							SELECT @lv__pred__schedule__hitime;
						END
					END
					ELSE
					BEGIN
						--Hypo is on the same day as @lotime. For "Every 1 Day" schedules, we will always enter this block

						--For both one-time-a-day and recurring-time-window schedules, DailyStartTime represents the earliest time
						-- a schedule can run
						IF @lv__pred__cursor__HypotheticalCurrentTime <= @lv__pred__schedule__DailyStartTime
						BEGIN
							INSERT INTO #ScheduleNextExecTimes (ExecStartTime)
							SELECT @lv__pred__schedule__DailyStartTime;
						END
						ELSE
						BEGIN
							IF @lv__pred__schedule__freq_subday_type = 1	--specified time of day
							BEGIN
								--since we know Hypo is already > than DailyStartTime, we missed our run and
								-- thus @hitime is our next run
								IF @lv__pred__schedule__hitime <= @lv__pred__schedule__ActiveEndtime
								BEGIN
									INSERT INTO #ScheduleNextExecTimes (ExecStartTime)
									SELECT @lv__pred__schedule__hitime;
								END
							END
							ELSE	--recurring-time-window schedule; remember, we now know that Hypo is > DailyStartTime
							BEGIN
								IF @lv__pred__cursor__HypotheticalCurrentTime > @lv__pred__schedule__DailyEndtime
								BEGIN	--we're after the time window; again, use @hitime
									IF @lv__pred__schedule__hitime <= @lv__pred__schedule__ActiveEndtime
									BEGIN
										INSERT INTO #ScheduleNextExecTimes (ExecStartTime)
										SELECT @lv__pred__schedule__hitime;
									END
								END
								ELSE
								BEGIN
									--we are still in the active time window. Need to use subday intervals
									SET @lv__pred__cursor__tmptime = @lv__pred__schedule__hitime;

									--Calculate the # of freq_subday_interval occurrences between hypo time and start time,
									-- then add that # of intervals and #+1 back to the start time, and see which one most nearly *follows* Hypo
									SET @lv__pred__schedule__intvdiff = (
										CASE 
											WHEN @lv__pred__schedule__freq_subday_type = 8 
											THEN DATEDIFF(HOUR, @lv__pred__schedule__DailyStartTime, 
														@lv__pred__cursor__HypotheticalCurrentTime) / @lv__pred__schedule__freq_subday_interval
											WHEN @lv__pred__schedule__freq_subday_type = 4
											THEN DATEDIFF(MINUTE, @lv__pred__schedule__DailyStartTime, 
														@lv__pred__cursor__HypotheticalCurrentTime) / @lv__pred__schedule__freq_subday_interval
											WHEN @lv__pred__schedule__freq_subday_type = 2
											THEN DATEDIFF(SECOND, @lv__pred__schedule__DailyStartTime, 
														@lv__pred__cursor__HypotheticalCurrentTime) / @lv__pred__schedule__freq_subday_interval
										END
										);

									SET @lv__pred__schedule__lotime = (
										CASE 
											WHEN @lv__pred__schedule__freq_subday_type = 8 
											THEN DATEADD(HOUR, @lv__pred__schedule__intvdiff*@lv__pred__schedule__freq_subday_interval, 
													@lv__pred__schedule__DailyStartTime)
											WHEN @lv__pred__schedule__freq_subday_type = 4
											THEN DATEADD(MINUTE, @lv__pred__schedule__intvdiff*@lv__pred__schedule__freq_subday_interval, 
													@lv__pred__schedule__DailyStartTime)
											WHEN @lv__pred__schedule__freq_subday_type = 2 
											THEN DATEADD(SECOND, @lv__pred__schedule__intvdiff*@lv__pred__schedule__freq_subday_interval, 
													@lv__pred__schedule__DailyStartTime)
										END );
										
									SET @lv__pred__schedule__hitime = (
										CASE 
											WHEN @lv__pred__schedule__freq_subday_type = 8 
											THEN DATEADD(HOUR, (@lv__pred__schedule__intvdiff+1)*@lv__pred__schedule__freq_subday_interval, 
															@lv__pred__schedule__DailyStartTime)
											WHEN @lv__pred__schedule__freq_subday_type = 4
											THEN DATEADD(MINUTE, (@lv__pred__schedule__intvdiff+1)*@lv__pred__schedule__freq_subday_interval, 
															@lv__pred__schedule__DailyStartTime)
											WHEN @lv__pred__schedule__freq_subday_type = 2 
											THEN DATEADD(SECOND, (@lv__pred__schedule__intvdiff+1)*@lv__pred__schedule__freq_subday_interval, 
															@lv__pred__schedule__DailyStartTime)
										END
										);

									IF @lv__pred__cursor__HypotheticalCurrentTime <= @lv__pred__schedule__lotime
									BEGIN
										--Use "lo" as the next execution time, as long as it is before the time-window-end
										IF @lv__pred__schedule__lotime <= @lv__pred__schedule__DailyEndtime
										BEGIN
											INSERT INTO #ScheduleNextExecTimes (ExecStartTime)
											SELECT @lv__pred__schedule__lotime;
										END
										ELSE
										BEGIN
											--past the end-time for this day, so revert to our original @hitime, which we stored in @tmptime,
											--(after checking that the schedule hasn't expired, of course)
											IF @lv__pred__cursor__tmptime <= @lv__pred__schedule__ActiveEndtime
											BEGIN
												INSERT INTO #ScheduleNextExecTimes (ExecStartTime)
												SELECT @lv__pred__cursor__tmptime;
											END
										END
									END
									ELSE
									BEGIN
										--Use "hi" as the next execution time.
										IF @lv__pred__schedule__hitime <= @lv__pred__schedule__DailyEndtime
										BEGIN
											INSERT INTO #ScheduleNextExecTimes (ExecStartTime)
											SELECT @lv__pred__schedule__hitime;
										END
										ELSE
										BEGIN	--past the end-time for this day, so revert to the old @hitime value, which we
												--stored in tmptime
											IF @lv__pred__cursor__tmptime <= @lv__pred__schedule__ActiveEndtime
											BEGIN
												--Is @hitime beyond the schedule expiration date?
												IF @lv__pred__schedule__hitime <= @lv__pred__schedule__ActiveEndtime
												BEGIN
													INSERT INTO #ScheduleNextExecTimes (ExecStartTime)
													SELECT @lv__pred__cursor__tmptime;
												END
											END
										END	-- is @hi interval-based value before today's window end time?
									END	--is Hypothetical time <= subday-interval-based @lotime?
								END	--is Hypo after or within the active time window?
							END  --IF schedule is one-time-of-day or recurring-time-window 
						END	--compare Hypo to DailyStartTime
					END	--test for whether lotime and Hypo are on the same day
				END	--Daily frequency

				--Weekly frequency	TODO: need to handle when freq_recurrence_factor is > 1. Right now, 
				-- my code just assumes that stuff is being executed every 1 week
				IF @lv__pred__schedule__freq_type = 8
				BEGIN
					IF @lv__pred__schedule__freq_subday_type > 1 
					BEGIN
						--We only need to calculate our next "time-interval alignment" time if > 1. (And even if we do, it will only be used
						-- if Hypo is the same as the next DOW that is applicable, and Hypo is NOT > than the Daily Start Time
						SET @lv__pred__schedule__intvdiff = (
							CASE 
								WHEN @lv__pred__schedule__freq_subday_type = 8 
								THEN DATEDIFF(HOUR, @lv__pred__schedule__DailyStartTime, 
											@lv__pred__cursor__HypotheticalCurrentTime) / @lv__pred__schedule__freq_subday_interval
								WHEN @lv__pred__schedule__freq_subday_type = 4
								THEN DATEDIFF(MINUTE, @lv__pred__schedule__DailyStartTime, 
											@lv__pred__cursor__HypotheticalCurrentTime) / @lv__pred__schedule__freq_subday_interval
								WHEN @lv__pred__schedule__freq_subday_type = 2
								THEN DATEDIFF(SECOND, @lv__pred__schedule__DailyStartTime, 
											@lv__pred__cursor__HypotheticalCurrentTime) / @lv__pred__schedule__freq_subday_interval
							END
							);

						SET @lv__pred__schedule__lotime = (
							CASE 
								WHEN @lv__pred__schedule__freq_subday_type = 8 
								THEN DATEADD(HOUR, @lv__pred__schedule__intvdiff*@lv__pred__schedule__freq_subday_interval, 
										@lv__pred__schedule__DailyStartTime)
								WHEN @lv__pred__schedule__freq_subday_type = 4
								THEN DATEADD(MINUTE, @lv__pred__schedule__intvdiff*@lv__pred__schedule__freq_subday_interval, 
										@lv__pred__schedule__DailyStartTime)
								WHEN @lv__pred__schedule__freq_subday_type = 2 
								THEN DATEADD(SECOND, @lv__pred__schedule__intvdiff*@lv__pred__schedule__freq_subday_interval, 
										@lv__pred__schedule__DailyStartTime)
							END );
										
						SET @lv__pred__schedule__hitime = (
							CASE 
								WHEN @lv__pred__schedule__freq_subday_type = 8 
								THEN DATEADD(HOUR, (@lv__pred__schedule__intvdiff+1)*@lv__pred__schedule__freq_subday_interval, 
												@lv__pred__schedule__DailyStartTime)
								WHEN @lv__pred__schedule__freq_subday_type = 4
								THEN DATEADD(MINUTE, (@lv__pred__schedule__intvdiff+1)*@lv__pred__schedule__freq_subday_interval, 
												@lv__pred__schedule__DailyStartTime)
								WHEN @lv__pred__schedule__freq_subday_type = 2 
								THEN DATEADD(SECOND, (@lv__pred__schedule__intvdiff+1)*@lv__pred__schedule__freq_subday_interval, 
												@lv__pred__schedule__DailyStartTime)
							END
							);
					END	--IF @lv__pred__schedule__freq_subday_type > 1 


					SET @lv__pred__cursor__HypoStartTime = NULL;

					--We have set DATEFIRST to 7, meaning Sunday is the first day of the week
					-- Thus, datepart(dw,Monday) = 2, datepart(dw,Saturday) = 7, etc.

					/*
						The following query figures out which days of the week the schedule is enabled for, 
						and also constructs a list of the next 7 days' worth of execution times, and then
						finally compares that 7 days with the DOW-enabled-for list and returns the MIN datetime
						that is in the DOW-enabled-for list
					*/
					; WITH DaysEnabled_Base AS (
						SELECT [DayEnabled] = CASE WHEN @lv__pred__schedule__freq_interval & 1 > 0 THEN 'sunday' ELSE NULL END UNION ALL
						SELECT CASE WHEN @lv__pred__schedule__freq_interval & 2 > 0 THEN 'monday' ELSE NULL END UNION ALL
						SELECT CASE WHEN @lv__pred__schedule__freq_interval & 4 > 0 THEN 'tuesday' ELSE NULL END UNION ALL
						SELECT CASE WHEN @lv__pred__schedule__freq_interval & 8 > 0 THEN 'wednesday' ELSE NULL END UNION ALL
						SELECT CASE WHEN @lv__pred__schedule__freq_interval & 16 > 0 THEN 'thursday' ELSE NULL END UNION ALL
						SELECT CASE WHEN @lv__pred__schedule__freq_interval & 32 > 0 THEN 'friday' ELSE NULL END UNION ALL
						SELECT CASE WHEN @lv__pred__schedule__freq_interval & 64 > 0 THEN 'saturday' ELSE NULL END
					),
					DaysEnabled AS (
						SELECT DayEnabled FROM DaysEnabled_Base WHERE DayEnabled IS NOT NULL 
					),
					HypoTimeWithActiveRangeConsidered AS (
						--We need to take into account the subday time range (or start time, if a once-a-day thing), but only for 
						-- our hypothetical time. Any future day after our hypothetical time can just use DailyStartTime
						SELECT [HypoTime] = 
							CASE WHEN @lv__pred__schedule__freq_subday_type = 1 
								THEN (CASE WHEN @lv__pred__cursor__HypotheticalCurrentTime > @lv__pred__schedule__DailyStartTime
												--too late, add a day
											THEN DATEADD(DAY,1, @lv__pred__schedule__DailyStartTime)
										ELSE @lv__pred__schedule__DailyStartTime
										END)
								ELSE --an "every X sec/min/hour" pattern. Are we outside (either before or after) our time range?
										--If not, use the interval-based time we calculated above.
									(CASE WHEN @lv__pred__cursor__HypotheticalCurrentTime < @lv__pred__schedule__DailyStartTime
											THEN @lv__pred__schedule__DailyStartTime
										WHEN @lv__pred__cursor__HypotheticalCurrentTime > @lv__pred__schedule__DailyEndtime
										THEN DATEADD(DAY,1, @lv__pred__schedule__DailyStartTime)
										ELSE (CASE WHEN @lv__pred__cursor__HypotheticalCurrentTime <= @lv__pred__schedule__lotime
												THEN @lv__pred__schedule__lotime
												ELSE @lv__pred__schedule__hitime
												END
												)
										END)
								END
					),
					NextWeekOfDays AS (
						SELECT [SeriesOfDays] = 
							HypoTime FROM HypoTimeWithActiveRangeConsidered		--this is either today or tomorrow
						UNION ALL
						--The rest of the days can be based on the Daily Start Time
						SELECT DATEADD(DAY, 1, @lv__pred__schedule__DailyStartTime)	
						UNION ALL
						SELECT DATEADD(DAY, 2, @lv__pred__schedule__DailyStartTime)	
						UNION ALL
						SELECT DATEADD(DAY, 3, @lv__pred__schedule__DailyStartTime)	
						UNION ALL
						SELECT DATEADD(DAY, 4, @lv__pred__schedule__DailyStartTime) 
						UNION ALL
						SELECT DATEADD(DAY, 5, @lv__pred__schedule__DailyStartTime)	
						UNION ALL
						SELECT DATEADD(DAY, 6, @lv__pred__schedule__DailyStartTime)
						UNION ALL
						SELECT DATEADD(DAY, 7, @lv__pred__schedule__DailyStartTime)
					),
					NextWeekWithDOW AS (
						SELECT SeriesOfDays, 
							HypoDOW = CASE DATEPART(dw, SeriesOfDays)
									WHEN 1 THEN 'sunday'
									WHEN 2 THEN 'monday' 
									WHEN 3 THEN 'tuesday'
									WHEN 4 THEN 'wednesday'
									WHEN 5 THEN 'thursday' 
									WHEN 6 THEN 'friday'
									WHEN 7 THEN 'saturday'
								END
						FROM NextWeekOfDays
					)
					SELECT @lv__pred__cursor__HypoStartTime = MIN(n.SeriesOfDays)
					FROM NextWeekWithDOW n
					WHERE n.HypoDOW IN (SELECT DayEnabled FROM DaysEnabled);

					IF @lv__pred__cursor__HypoStartTime IS NOT NULL
					BEGIN
						IF @lv__pred__cursor__HypoStartTime <= @lv__pred__schedule__ActiveEndtime
						BEGIN
							INSERT INTO #ScheduleNextExecTimes (ExecStartTime)
							SELECT @lv__pred__cursor__HypoStartTime;
						END
					END
				END	--Weekly frequency


				--Monthly frequency
				IF @lv__pred__schedule__freq_type = 16
				BEGIN
					/* I've had some difficulty in understanding the logic used by SQL Agent when the schedule is set for
						"Every X months" and X is > 1.		(i.e. the freq_recurrence_factor is > 1)
						Because we limit @PointInTime to 3 days in the future and the @HoursForward value is also limited, in practice
						we can probably just use the Next Scheduled Time that we obtained from msdb.dbo.sysjobschedules.

						If that is NULL somehow, and @lv__pred__schedule__freq_recurrence_factor = 1, then we do
						a "college try" effort on obtaining the next execution.
					*/
					IF @lv__pred__schedule__NextScheduleTimeFromSystem IS NOT NULL 
						AND @lv__pred__cursor__HypotheticalCurrentTime <= @lv__pred__schedule__NextScheduleTimeFromSystem
					BEGIN
						INSERT INTO #ScheduleNextExecTimes (ExecStartTime)
						SELECT @lv__pred__schedule__NextScheduleTimeFromSystem
					END
					ELSE	--our Next Scheduled Time was invalid or our hypothetical time is already past it. If the schedule is an
					BEGIN	-- "Every 1 month" sort of thing, then it is pretty straightforward to capture the next time
						IF @lv__pred__schedule__freq_recurrence_factor = 1
						BEGIN
							--get a "first day of the month" value from our Hypo time
							SET @lv__pred__cursor__tmptime = CONVERT(DATETIME, 
								CONVERT(CHAR(4),YEAR(@lv__pred__cursor__HypotheticalCurrentTime)) + '-' +
								CONVERT(VARCHAR(2),MONTH(@lv__pred__cursor__HypotheticalCurrentTime)) +  '-01'
								);

							--Is our hypothetical time already past the day-of-the-month? If so, add a month.
							IF DATEPART(DAY, @lv__pred__cursor__HypotheticalCurrentTime) > @lv__pred__schedule__freq_interval
							BEGIN
								SET @lv__pred__cursor__tmptime = DATEADD(MONTH, 1, @lv__pred__cursor__tmptime);
							END

							--If our month doesn't have as many days as specified in @lv__pred__schedule__freq_interval, 
							-- then add a month. Since we never have consecutive months with < 31 days, we are guaranteed 
							-- to get a month with enough days just by adding 1 month.
							IF DATEPART(DAY, DATEADD(DAY, -1, DATEADD(MONTH, 1, @lv__pred__cursor__tmptime))) < @lv__pred__schedule__freq_interval
							BEGIN
								SET @lv__pred__cursor__tmptime = DATEADD(MONTH, 1, @lv__pred__cursor__tmptime);
							END

							--tmptime is set to the first of the month. Add the day/time portion in based on the Active Start Time
							SET @lv__pred__cursor__tmptime2 = DATEADD(MONTH, 
								DATEDIFF(MONTH, @lv__pred__schedule__ActiveStartTime, @lv__pred__cursor__tmptime),
								@lv__pred__schedule__ActiveStartTime);

							--And for the end-time
							SET @lv__pred__cursor__tmptime3 = DATEADD(MONTH, 
								DATEDIFF(MONTH, @lv__pred__schedule__ActiveEndtime, @lv__pred__cursor__tmptime),
								@lv__pred__schedule__ActiveEndtime);

							--If this is a day in the future, then add it in
							IF DATEDIFF(DAY, @lv__pred__cursor__HypotheticalCurrentTime, @lv__pred__cursor__tmptime2) > 0 
							BEGIN
								IF @lv__pred__cursor__tmptime2 <= @lv__pred__schedule__ActiveEndtime
								BEGIN
									INSERT INTO #ScheduleNextExecTimes (ExecStartTime)
									SELECT @lv__pred__cursor__tmptime2;
								END
							END
							ELSE	--because of the above logic, tmptime2 cannot be an EARLIER day than Hypo, so it must be the same day
							BEGIN
								IF @lv__pred__cursor__HypotheticalCurrentTime <= @lv__pred__cursor__tmptime2
								BEGIN	--before the start time; doesn't matter whether schedule is one-time-a-day or recurring
									INSERT INTO #ScheduleNextExecTimes (ExecStartTime)
									SELECT @lv__pred__cursor__tmptime2;
								END
								ELSE IF @lv__pred__schedule__freq_subday_type = 1
									OR @lv__pred__cursor__HypotheticalCurrentTime > @lv__pred__cursor__tmptime3
								BEGIN
									--if a once-a-day schedule, or if a recurring schedule that we are after the end-time, 
									-- then we are too late. 
									--Because this proc limits "looking forward" to just a few days, the fact that this
									-- schedule won't be applicable for a months means we aren't even going to try to
									-- assemble the correct "next start time".

									--No-op
									SET @lv__pred__cursor__HypotheticalCurrentTime = @lv__pred__cursor__HypotheticalCurrentTime;
								END
								ELSE
								BEGIN
									--we are in the middle of a recurring-time-window schedule's time frame. 
									--Remember, tmptime2 is the beginning of the range, and tmptime3 is the end
									SET @lv__pred__schedule__intvdiff = (
										CASE 
											WHEN @lv__pred__schedule__freq_subday_type = 8 
											THEN DATEDIFF(HOUR, @lv__pred__cursor__tmptime2, 
														@lv__pred__cursor__HypotheticalCurrentTime) / @lv__pred__schedule__freq_subday_interval
											WHEN @lv__pred__schedule__freq_subday_type = 4
											THEN DATEDIFF(MINUTE, @lv__pred__cursor__tmptime2, 
														@lv__pred__cursor__HypotheticalCurrentTime) / @lv__pred__schedule__freq_subday_interval
											WHEN @lv__pred__schedule__freq_subday_type = 2
											THEN DATEDIFF(SECOND, @lv__pred__cursor__tmptime2, 
														@lv__pred__cursor__HypotheticalCurrentTime) / @lv__pred__schedule__freq_subday_interval
										END
										);

									SET @lv__pred__schedule__lotime = (
										CASE 
											WHEN @lv__pred__schedule__freq_subday_type = 8 
											THEN DATEADD(HOUR, @lv__pred__schedule__intvdiff*@lv__pred__schedule__freq_subday_interval, 
													@lv__pred__cursor__tmptime2)
											WHEN @lv__pred__schedule__freq_subday_type = 4
											THEN DATEADD(MINUTE, @lv__pred__schedule__intvdiff*@lv__pred__schedule__freq_subday_interval, 
													@lv__pred__cursor__tmptime2)
											WHEN @lv__pred__schedule__freq_subday_type = 2 
											THEN DATEADD(SECOND, @lv__pred__schedule__intvdiff*@lv__pred__schedule__freq_subday_interval, 
													@lv__pred__cursor__tmptime2)
										END );

									SET @lv__pred__schedule__hitime = (
										CASE 
											WHEN @lv__pred__schedule__freq_subday_type = 8 
											THEN DATEADD(HOUR, (@lv__pred__schedule__intvdiff+1)*@lv__pred__schedule__freq_subday_interval, 
															@lv__pred__cursor__tmptime2)
											WHEN @lv__pred__schedule__freq_subday_type = 4
											THEN DATEADD(MINUTE, (@lv__pred__schedule__intvdiff+1)*@lv__pred__schedule__freq_subday_interval, 
															@lv__pred__cursor__tmptime2)
											WHEN @lv__pred__schedule__freq_subday_type = 2 
											THEN DATEADD(SECOND, (@lv__pred__schedule__intvdiff+1)*@lv__pred__schedule__freq_subday_interval, 
															@lv__pred__cursor__tmptime2)
										END
										);

									IF @lv__pred__cursor__HypotheticalCurrentTime <= @lv__pred__schedule__lotime
									BEGIN
										--Use "lo" as the next execution time, as long as it is before the time-window-end
										IF @lv__pred__schedule__lotime <= @lv__pred__schedule__DailyEndtime
										BEGIN
											INSERT INTO #ScheduleNextExecTimes (ExecStartTime)
											SELECT @lv__pred__schedule__lotime;
										END
										--ELSE  next exec is a month out, so don't worry about it
									END
									ELSE
									BEGIN
										--Use "hi" as the next execution time.
										IF @lv__pred__schedule__hitime <= @lv__pred__schedule__DailyEndtime
										BEGIN
											INSERT INTO #ScheduleNextExecTimes (ExecStartTime)
											SELECT @lv__pred__schedule__hitime;
										END
										--ELSE	next exec is a month out, so don't worry about it
									END
								END
							END
						END
					END
				END	--Monthly frequency


				--TODO: Monthly-relative frequency		--TODO: worry about freq_recurrence_factor > 1 later???
				IF @lv__pred__schedule__freq_type = 32
				BEGIN

					SET @lv__pred__schedule__freq_type = 32
				END	--Monthly-relative frequency

			END	--end of Schedule loop that obtains "Next Exec Time" for each schedule

			CLOSE PredictIterationSchedules;
			DEALLOCATE PredictIterationSchedules;


			SET @lv__pred__schedule__MinNextTime = NULL;

			SELECT @lv__pred__schedule__MinNextTime = MIN(t.ExecStartTime)
			FROM #ScheduleNextExecTimes t
			WHERE t.ExecStartTime IS NOT NULL;

			IF @lv__pred__schedule__MinNextTime IS NULL
			BEGIN
				--SET @lv__pred__cursor__HypotheticalCurrentTime = DATEADD(SECOND, 1,@lv__pred__OverallWindowEndTime);
				--debug:
				SET @lv__pred__cursor__HypotheticalCurrentTime = DATEADD(minute, 6, @lv__pred__cursor__HypotheticalCurrentTime)
			END
			ELSE
			BEGIN
				INSERT INTO #HypotheticalRuns 
				(native_job_id, JobStartTime, JobExpectedEndTime)
				SELECT @lv__pred__cursor__NativeJobID, @lv__pred__cursor__HypotheticalCurrentTime, 
					DATEADD(SECOND, @lv__pred__cursor__AvgSuccessDur_seconds, @lv__pred__cursor__HypotheticalCurrentTime);

				--Set the new hypothetical time to 1 second after the hypothetical run completes
				SET @lv__pred__cursor__HypotheticalCurrentTime = 
					DATEADD(SECOND, @lv__pred__cursor__AvgSuccessDur_seconds+1, @lv__pred__cursor__HypotheticalCurrentTime)
				--debug: 
				SET @lv__pred__cursor__HypotheticalCurrentTime = DATEADD(minute, 6, @lv__pred__cursor__HypotheticalCurrentTime)
			END
			

			--debug (ensures that we only go through the loop once per Job):
			SET @lv__pred__cursor__HypotheticalCurrentTime = DATEADD(minute, 6, @lv__pred__cursor__HypotheticalCurrentTime)
		END 
	END	--end of outer WHILE loop

	CLOSE CreateHypotheticalRuns
	DEALLOCATE CreateHypotheticalRuns










afterpredict:
	/*********************************************************************************************************************************************
	*********************************************************************************************************************************************

														Part 3: SQL Agent Configuration Options

	*********************************************************************************************************************************************
	*********************************************************************************************************************************************/
configstart: 

	IF @DisplayConfigOptions > 0
	BEGIN
		SET @lv__beforedt = GETDATE();
		--***Location 0: Obtain info from the registry
		BEGIN TRY
			EXECUTE master.dbo.xp_instance_regread N'HKEY_LOCAL_MACHINE',N'SOFTWARE\Microsoft\MSSQLServer\SQLServerAgent',N'JobHistoryMaxRows',@lv__cfg__MaxHistoryRows OUTPUT,N'no_output';
			EXECUTE master.dbo.xp_instance_regread N'HKEY_LOCAL_MACHINE',N'SOFTWARE\Microsoft\MSSQLServer\SQLServerAgent',N'JobHistoryMaxRowsPerJob',@lv__cfg__MaxHistoryRowsPerJob OUTPUT;
			EXECUTE master.dbo.xp_instance_regread N'HKEY_LOCAL_MACHINE',N'SOFTWARE\Microsoft\MSSQLServer\SQLServerAgent',N'RestartSQLServer',@lv__cfg__ShouldAgentRestartSQL OUTPUT;
			EXECUTE master.dbo.xp_instance_regread N'HKEY_LOCAL_MACHINE',N'SOFTWARE\Microsoft\MSSQLServer\SQLServerAgent',N'ErrorLogFile',@lv__cfg__errorlog_file OUTPUT;
			EXECUTE master.dbo.xp_instance_regread N'HKEY_LOCAL_MACHINE',N'SOFTWARE\Microsoft\MSSQLServer\SQLServerAgent',N'ErrorLoggingLevel',@lv__cfg__errorlogging_level OUTPUT;
			EXECUTE master.dbo.xp_instance_regread N'HKEY_LOCAL_MACHINE',N'SOFTWARE\Microsoft\MSSQLServer\SQLServerAgent',N'ErrorMonitor',@lv__cfg__error_recipient OUTPUT;
			EXECUTE master.dbo.xp_instance_regread N'HKEY_LOCAL_MACHINE',N'SOFTWARE\Microsoft\MSSQLServer\SQLServerAgent',N'MonitorAutoStart',@lv__cfg__monitor_autostart OUTPUT;
			EXECUTE master.dbo.xp_instance_regread N'HKEY_LOCAL_MACHINE',N'SOFTWARE\Microsoft\MSSQLServer\SQLServerAgent',N'ServerHost',@lv__cfg__local_host_server OUTPUT;
			EXECUTE master.dbo.xp_instance_regread N'HKEY_LOCAL_MACHINE',N'SOFTWARE\Microsoft\MSSQLServer\SQLServerAgent',N'JobShutdownTimeout',@lv__cfg__job_shutdown_timeout OUTPUT;
			EXECUTE master.dbo.xp_instance_regread N'HKEY_LOCAL_MACHINE',N'SOFTWARE\Microsoft\MSSQLServer\SQLServerAgent',N'LoginTimeout',@lv__cfg__login_timeout OUTPUT;
			EXECUTE master.dbo.xp_instance_regread N'HKEY_LOCAL_MACHINE',N'SOFTWARE\Microsoft\MSSQLServer\SQLServerAgent',N'IdleCPUPercent',@lv__cfg__idle_cpu_percent OUTPUT;
			EXECUTE master.dbo.xp_instance_regread N'HKEY_LOCAL_MACHINE',N'SOFTWARE\Microsoft\MSSQLServer\SQLServerAgent',N'IdleCPUDuration',@lv__cfg__idle_cpu_duration OUTPUT;
			EXECUTE master.dbo.xp_instance_regread N'HKEY_LOCAL_MACHINE',N'SOFTWARE\Microsoft\MSSQLServer\SQLServerAgent',N'OemErrorLog',@lv__cfg__oem_errorlog OUTPUT;
			EXECUTE master.dbo.xp_instance_regread N'HKEY_LOCAL_MACHINE',N'SOFTWARE\Microsoft\MSSQLServer\SQLServerAgent',N'AlertReplaceRuntimeTokens',@lv__cfg__alert_replace_runtime_tokens OUTPUT;
			EXECUTE master.dbo.xp_instance_regread N'HKEY_LOCAL_MACHINE',N'SOFTWARE\Microsoft\MSSQLServer\SQLServerAgent',N'CoreEngineMask',@lv__cfg__cpu_poller_enabled OUTPUT
			EXECUTE master.dbo.xp_instance_regread N'HKEY_LOCAL_MACHINE',N'SOFTWARE\Microsoft\MSSQLServer\SQLServerAgent',N'UseDatabaseMail',@lv__cfg__use_databasemail OUTPUT;
			EXECUTE master.dbo.xp_instance_regread N'HKEY_LOCAL_MACHINE',N'SOFTWARE\Microsoft\MSSQLServer\SQLServerAgent',N'DatabaseMailProfile',@lv__cfg__databasemail_profile OUTPUT;

			SELECT @lv__cfg__tmpregstr = (N'SYSTEM\CurrentControlSet\Services\' + 
						CASE WHEN SERVERPROPERTY('INSTANCENAME') IS NOT NULL
							THEN N'SQLAgent$' + CONVERT (sysname, SERVERPROPERTY('INSTANCENAME'))
							ELSE N'SQLServerAgent'
							END);
			EXECUTE master.dbo.xp_regread N'HKEY_LOCAL_MACHINE',@lv__cfg__tmpregstr,N'Start',@lv__cfg__ServiceStartupSetting OUTPUT;
		END TRY
		BEGIN CATCH
			RAISERROR(N'Error occurred while obtaining Agent config values. Comparison of config option values with defaults will not occur.', 11, 1);
			SET @output__DisplayConfig = 0;
			SELECT @lv__ErrorText = ERROR_MESSAGE(), 
					@lv__ErrorSeverity	= ERROR_SEVERITY(), 
					@lv__ErrorState = ERROR_STATE();
			SET @lv__ErrorText = @lv__ErrorText + ' (Severity: ' + CONVERT(VARCHAR(20),@lv__ErrorSeverity) + ') (State: ' + CONVERT(VARCHAR(20),@lv__ErrorState) + ')'

			RAISERROR( @lv__ErrorText, 11, 1);

			GOTO afterconfig
		END CATCH

		SET @lv__afterdt = GETDATE();

		IF DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt) > @lv__slownessthreshold
		BEGIN
			SET @lv__ErrorText = N'   ***dbg: obtaining SQL Agent config from registry took ' + CONVERT(NVARCHAR(40),DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt)) + N' milliseconds';
			RAISERROR(@lv__ErrorText, 10, 1) WITH NOWAIT;
		END

		SET @lv__beforedt = GETDATE();

		BEGIN TRY
			--***Location 1: Determine which config options we need to return to the user
			IF ISNULL(@lv__cfg__MaxHistoryRows,-1) <> 1000
				OR @DisplayConfigOptions = 2
			BEGIN
				INSERT INTO #OptionsToDisplay(OptionTag, OptionValue, OptionNormalValue)
				SELECT 'Max History Rows', ISNULL(CONVERT(VARCHAR(100),@lv__cfg__MaxHistoryRows),'<null>'), '1000';
			END

			IF ISNULL(@lv__cfg__MaxHistoryRowsPerJob,-1) <> 100
				OR @DisplayConfigOptions = 2
			BEGIN
				INSERT INTO #OptionsToDisplay(OptionTag, OptionValue, OptionNormalValue)
				SELECT 'Max History Rows Per Job', ISNULL(CONVERT(VARCHAR(100),@lv__cfg__MaxHistoryRowsPerJob),'<null>'), '100';
			END

			--2 = automatic, 3 = manual, 4 = disabled
			IF ISNULL(@lv__cfg__ServiceStartupSetting,-1) <> 2 OR @DisplayConfigOptions = 2
			BEGIN
				INSERT INTO #OptionsToDisplay(OptionTag, OptionValue, OptionNormalValue)
				SELECT 'Agent Service Startup', 
					CASE WHEN @lv__cfg__ServiceStartupSetting IS NULL THEN '<null>' 
						WHEN @lv__cfg__ServiceStartupSetting = 2 THEN 'Automatic'
						WHEN @lv__cfg__ServiceStartupSetting = 3 THEN 'Manual' 
						WHEN @lv__cfg__ServiceStartupSetting = 4 THEN 'Disabled'
						ELSE CONVERT(VARCHAR(20), @lv__cfg__ServiceStartupSetting) + ' - Unknown'
						END, 
					'Automatic';
			END

			IF ISNULL(@lv__cfg__ShouldAgentRestartSQL,-5) <> 1 OR @DisplayConfigOptions = 2
			BEGIN
				INSERT INTO #OptionsToDisplay(OptionTag, OptionValue, OptionNormalValue)
				SELECT 'Should Agent Restart SQL Engine?', 
						CASE WHEN @lv__cfg__ShouldAgentRestartSQL IS NULL THEN '<null>'
							WHEN @lv__cfg__ShouldAgentRestartSQL = 1 THEN 'Yes'
							WHEN @lv__cfg__ShouldAgentRestartSQL = 0 THEN 'No'
							ELSE CONVERT(VARCHAR(20), @lv__cfg__ShouldAgentRestartSQL) + ' - Unknown'
							END , 
						'Yes';
			END

			IF ISNULL(@lv__cfg__monitor_autostart,-5) <> 1 OR @DisplayConfigOptions = 2
			BEGIN
				INSERT INTO #OptionsToDisplay(OptionTag, OptionValue, OptionNormalValue)
				SELECT 'Should Agent Restart itself?', 
						CASE WHEN @lv__cfg__monitor_autostart IS NULL THEN '<null>'
							WHEN @lv__cfg__monitor_autostart = 1 THEN 'Yes'
							WHEN @lv__cfg__monitor_autostart = 0 THEN 'No'
							ELSE CONVERT(VARCHAR(20), @lv__cfg__monitor_autostart) + ' - Unknown'
							END, 
					'Yes';
			END

			IF ISNULL(@lv__cfg__errorlogging_level,-1) <> 3 OR @DisplayConfigOptions = 2
			BEGIN
				INSERT INTO #OptionsToDisplay
					(OptionTag, OptionValue, OptionNormalValue)
				SELECT 'Error Log Level', ISNULL(CONVERT(VARCHAR(100),@lv__cfg__errorlogging_level),'<null>'), '3';
			END

			--we only display the location of the error log file if the user has asked to look at all config options
			IF @DisplayConfigOptions = 2
			BEGIN
				INSERT INTO #OptionsToDisplay(OptionTag, OptionValue, OptionNormalValue)
				SELECT 'Error Log Location', ISNULL(CONVERT(VARCHAR(100),@lv__cfg__errorlog_file),'<null>'), 'n/a';
			END

			IF ISNULL(@lv__cfg__cpu_poller_enabled,-5) <> 32 OR @DisplayConfigOptions = 2
			BEGIN
				INSERT INTO #OptionsToDisplay(OptionTag, OptionValue, OptionNormalValue)
				SELECT 'Define Idle CPU Threshold', 
						CASE WHEN @lv__cfg__cpu_poller_enabled IS NULL THEN '<null>'
							WHEN @lv__cfg__cpu_poller_enabled = 32 THEN 'No'
							WHEN @lv__cfg__cpu_poller_enabled = 1 THEN 'Yes'
							ELSE CONVERT(VARCHAR(20), @lv__cfg__cpu_poller_enabled) + ' - Unknown'
							END, 
						'No';

				INSERT INTO #OptionsToDisplay(OptionTag, OptionValue, OptionNormalValue)
				SELECT 'Idle CPU % Threshold', ISNULL(CONVERT(VARCHAR(100),@lv__cfg__idle_cpu_percent),'<null>'), '10';

				INSERT INTO #OptionsToDisplay(OptionTag, OptionValue, OptionNormalValue)
				SELECT 'Idle CPU % Duration', ISNULL(CONVERT(VARCHAR(100),@lv__cfg__idle_cpu_duration),'<null>'), '600';
			END


			IF ISNULL(@lv__cfg__login_timeout,-5) <> 30 OR @DisplayConfigOptions = 2
			BEGIN
				INSERT INTO #OptionsToDisplay(OptionTag, OptionValue, OptionNormalValue)
				SELECT 'Login Timeout (sec)', ISNULL(CONVERT(VARCHAR(100),@lv__cfg__login_timeout),'<null>'), '30';
			END

			IF ISNULL(@lv__cfg__job_shutdown_timeout,-5) <> 15 OR @DisplayConfigOptions = 2
			BEGIN
				INSERT INTO #OptionsToDisplay(OptionTag, OptionValue, OptionNormalValue)
				SELECT 'Job Shutdown Timeout (sec)', ISNULL(CONVERT(VARCHAR(100),@lv__cfg__job_shutdown_timeout),'<null>'), '15';
			END

			IF ISNULL(@lv__cfg__use_databasemail,-1) <> 0 OR @DisplayConfigOptions = 2
			BEGIN
				INSERT INTO #OptionsToDisplay(OptionTag, OptionValue, OptionNormalValue)
				SELECT 'Use Database Mail?', 
						CASE WHEN @lv__cfg__use_databasemail IS NULL THEN '<null>'
							WHEN @lv__cfg__use_databasemail = 0 THEN 'No'
							WHEN @lv__cfg__use_databasemail = 1 THEN 'Yes'
							ELSE CONVERT(VARCHAR(20), @lv__cfg__use_databasemail) + ' - Unknown'
						END , 
						'No';

				IF ISNULL(@lv__cfg__use_databasemail,-1) <> 1 OR @DisplayConfigOptions = 2
				BEGIN
					INSERT INTO #OptionsToDisplay(OptionTag, OptionValue, OptionNormalValue)
					SELECT 'Database Mail Profile', ISNULL(CONVERT(VARCHAR(100),@lv__cfg__databasemail_profile),'<null>'), '<null>';			
				END
			END

			IF @lv__cfg__error_recipient IS NOT NULL OR @DisplayConfigOptions = 2
			BEGIN
				INSERT INTO #OptionsToDisplay(OptionTag, OptionValue, OptionNormalValue)
				SELECT 'Net Send Error Recipient', ISNULL(CONVERT(VARCHAR(100),@lv__cfg__error_recipient),'<null>'), '<null>';
			END

			IF @lv__cfg__local_host_server IS NOT NULL OR @DisplayConfigOptions = 2
			BEGIN
				INSERT INTO #OptionsToDisplay(OptionTag, OptionValue, OptionNormalValue)
				SELECT 'Local Host Server', ISNULL(CONVERT(VARCHAR(100),@lv__cfg__local_host_server),'<null>'), '<null>';
			END

			IF ISNULL(@lv__cfg__oem_errorlog,-1) <> 0 OR @DisplayConfigOptions = 2
			BEGIN
				INSERT INTO #OptionsToDisplay(OptionTag, OptionValue, OptionNormalValue)
				SELECT 'OEM Error Log', 
						CASE WHEN @lv__cfg__oem_errorlog IS NULL THEN '<null>'
							WHEN @lv__cfg__oem_errorlog = 0 THEN 'No'
							WHEN @lv__cfg__oem_errorlog = 1 THEN 'Yes'
							ELSE CONVERT(VARCHAR(20), @lv__cfg__oem_errorlog) + ' - Unknown'
						END , 
						'No';
			END

			IF ISNULL(@lv__cfg__alert_replace_runtime_tokens,-5) <> 0 OR @DisplayConfigOptions = 2
			BEGIN
				INSERT INTO #OptionsToDisplay(OptionTag, OptionValue, OptionNormalValue)
				SELECT 'Alert Token Replacement', 
					CASE WHEN @lv__cfg__alert_replace_runtime_tokens IS NULL THEN '<null>'
						WHEN @lv__cfg__alert_replace_runtime_tokens = 0 THEN 'No'
						WHEN @lv__cfg__alert_replace_runtime_tokens = 1 THEN 'Yes'
						ELSE CONVERT(VARCHAR(20), @lv__cfg__alert_replace_runtime_tokens) + ' - Unknown'
					END , 
					'No';
			END

			--If we have no entries, mark the appropriate display variable
			IF NOT EXISTS (SELECT * FROM #OptionsToDisplay)
			BEGIN
				SET @output__DisplayConfig = 0;
			END
		END TRY
		BEGIN CATCH
			RAISERROR(N'Error occurred while comparing Agent config values with default. Agent config option comparison results will not be displayed.', 11, 1);
			SET @output__DisplayConfig = 0;
			SELECT @lv__ErrorText = ERROR_MESSAGE(), 
					@lv__ErrorSeverity	= ERROR_SEVERITY(), 
					@lv__ErrorState = ERROR_STATE();
			SET @lv__ErrorText = @lv__ErrorText + ' (Severity: ' + CONVERT(VARCHAR(20),@lv__ErrorSeverity) + ') (State: ' + CONVERT(VARCHAR(20),@lv__ErrorState) + ')'

			RAISERROR( @lv__ErrorText, 11, 1);

			GOTO afterconfig
		END CATCH

		SET @lv__afterdt = GETDATE();

		IF DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt) > @lv__slownessthreshold
		BEGIN
			SET @lv__ErrorText = N'   ***dbg: checking SQL Agent config values took ' + CONVERT(NVARCHAR(40),DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt)) + N' milliseconds';
			RAISERROR(@lv__ErrorText, 10, 1) WITH NOWAIT;
		END
	END		--IF @DisplayConfigOptions > 0


afterconfig: 

	/*********************************************************************************************************************************************
	*********************************************************************************************************************************************

														Part 3: SQL Agent Log Files

	*********************************************************************************************************************************************
	*********************************************************************************************************************************************/
	IF @DisplayAgentLog > 0
	BEGIN
		--Read the SQL Agent logs. 

		/* If xp_readerrorlog is passed a log file # that doesn't exist, we get this error
			Msg 22004, Level 16, State 1, Line 0
			xp_readerrorlog() returned error 2, 'The system cannot find the file specified.'

			I tried writing a loop with a TRY/CATCH block to make sure I get all log files, but 
			I got "a severe error has occurred on this command" when trying to insert data into a 
			temp table.
		*/ 
		--Note that this statement still works even when SQL Agent is off
		--Also note that xp_readerrorlog can throw errors that are not catch-able, apparently when it is used to insert into a temp table.
		-- So the TRY/CATCH here is not guaranteed to be effective
		BEGIN TRY
			INSERT INTO #SQLAgentLog (LogDate, ErrorLevel, aText)
			EXEC xp_readerrorlog 0,	--log file # (0 is current)
				2,	--SQL Agent log
				null,	--search string 1
				null,	--search string 2
				null,	--search start time
				null,	--search end time
				'Desc'	--order results 
				;

			SET @lv__log__log1processing = 1
		END TRY
		BEGIN CATCH
			SET @output__DisplayAgentLog = 0;
			RAISERROR(N'Error occurred when obtaining the current SQL Agent error log. Some loss in functionality may occur.', 11, 1);
			SELECT @lv__ErrorText = ERROR_MESSAGE(), 
					@lv__ErrorSeverity	= ERROR_SEVERITY(), 
					@lv__ErrorState = ERROR_STATE();
			SET @lv__ErrorText = @lv__ErrorText + ' (Severity: ' + CONVERT(VARCHAR(20),@lv__ErrorSeverity) + ') (State: ' + CONVERT(VARCHAR(20),@lv__ErrorState) + ')'

			RAISERROR( @lv__ErrorText, 11, 1);
		END CATCH

		IF @lv__log__log1processing = 1		--log has been read, but not processed
		BEGIN
			SET @lv__beforedt = GETDATE();

			--***Location 0: misc handling for current log file
			--There are some SQL Agent log messages that are very common, and are not what we'd typically be looking for. 
			-- make sure these are omitted.
			BEGIN TRY
				DELETE 
				FROM #SQLAgentLog 
				WHERE (
					atext LIKE '%Job completion for % is being logged to sysjobhistory%'
					OR atext LIKE '%Job % has been requested to run by Schedule%'
					OR atext LIKE '%Saving % for all updated job schedules...'
					OR atext LIKE '% job schedule(s) saved%'
					OR (@DisplayAgentLog = 1 AND atext LIKE '%The Messenger service has not been started - NetSend notifications will not be sent%')
				);

				SELECT @lv__log__maxTabID = MAX(idcol) FROM #SQLAgentLog;
 
				UPDATE targ 
				SET FileNumber = 0,
					isLastRecord = CASE WHEN idcol <> @lv__log__maxTabID THEN 0 ELSE 1 END
				FROM #SQLAgentLog targ 
				WHERE FileNumber IS NULL 
				;

				SET @lv__log__log1processing = 2
			END TRY
			BEGIN CATCH
				RAISERROR(N'Error occurred while post-processing the current SQL Agent log file. The SQL Agent log will not be displayed.', 11, 1);
				SET @output__DisplayAgentLog = 0;
				SET @lv__log__log1processing = -2;
				SET @lv__log__log2processing = -3;
				SET @lv__log__log3processing = -3;
				SELECT @lv__ErrorText = ERROR_MESSAGE(), 
						@lv__ErrorSeverity	= ERROR_SEVERITY(), 
						@lv__ErrorState = ERROR_STATE();
				SET @lv__ErrorText = @lv__ErrorText + ' (Severity: ' + CONVERT(VARCHAR(20),@lv__ErrorSeverity) + ') (State: ' + CONVERT(VARCHAR(20),@lv__ErrorState) + ')'

				RAISERROR( @lv__ErrorText, 11, 1);

				GOTO afteragentlog
			END CATCH

			SET @lv__afterdt = GETDATE();

			IF DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt) > @lv__slownessthreshold
			BEGIN
				SET @lv__ErrorText = N'   ***dbg: post-processing on current SQL Agent log file took ' + CONVERT(NVARCHAR(40),DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt)) + N' milliseconds';
				RAISERROR(@lv__ErrorText, 10, 1) WITH NOWAIT;
			END
		END			--IF @lv__log__log1processing = 1		--log has been read, but not processed

		--only get additional files if @DisplayAgentLog = 3
		IF @DisplayAgentLog = 3
		BEGIN
			SET @lv__beforedt = GETDATE();

			--***Location 1: Get the most recent non-active log file and process
			BEGIN TRY
				INSERT INTO #SQLAgentLog (LogDate, ErrorLevel, aText)
					EXEC xp_readerrorlog 1,	--log file # (0 is current)
						2,	--SQL Agent log
						null,	--search string 1
						null,	--search string 2
						null,	--search start time
						null,	--search end time
						'Desc'	--order results 
						;
				SET @lv__log__log2processing = 1;
			END TRY
			BEGIN CATCH
				RAISERROR(N'Error occurred while obtaining the most-recent non-active SQL Agent log file. The SQL Agent log will not be displayed.', 11, 1);
				SET @output__DisplayAgentLog = 0;
				SET @lv__log__log2processing = -1;
				SELECT @lv__ErrorText = ERROR_MESSAGE(), 
						@lv__ErrorSeverity	= ERROR_SEVERITY(), 
						@lv__ErrorState = ERROR_STATE();
				SET @lv__ErrorText = @lv__ErrorText + ' (Severity: ' + CONVERT(VARCHAR(20),@lv__ErrorSeverity) + ') (State: ' + CONVERT(VARCHAR(20),@lv__ErrorState) + ')'

				RAISERROR( @lv__ErrorText, 11, 1);

				GOTO afteragentlog
			END CATCH

			SET @lv__afterdt = GETDATE();

			IF DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt) > @lv__slownessthreshold
			BEGIN
				SET @lv__ErrorText = N'   ***dbg: obtaining most-recent non-active SQL Agent log file took ' + CONVERT(NVARCHAR(40),DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt)) + N' milliseconds';
				RAISERROR(@lv__ErrorText, 10, 1) WITH NOWAIT;
			END


			IF @lv__log__log2processing = 1
			BEGIN 
				SET @lv__beforedt = GETDATE();

				--There are some SQL Agent log messages that are very common, and are not what we'd typically be looking for. 
				-- make sure these are omitted
				BEGIN TRY
					DELETE 
					FROM #SQLAgentLog 
					WHERE (
						atext LIKE '%Job completion for % is being logged to sysjobhistory%'
						OR atext LIKE '%Job % has been requested to run by Schedule%'
						OR atext LIKE '%Saving % for all updated job schedules...'
						OR atext LIKE '% job schedule(s) saved%'
						OR (@DisplayAgentLog = 1 AND atext LIKE '%The Messenger service has not been started - NetSend notifications will not be sent%')
					);

					SELECT @lv__log__maxTabID = MAX(idcol) FROM #SQLAgentLog;

					UPDATE targ 
					SET FileNumber = 1,
						isLastRecord = CASE WHEN idcol <> @lv__log__maxTabID THEN 0 ELSE 1 END
					FROM #SQLAgentLog targ 
					WHERE FileNumber IS NULL 
					;

					SET @lv__log__log2processing = 2;
				END TRY
				BEGIN CATCH
					RAISERROR(N'Error occurred while post-processing the most-recent non-active SQL Agent log file. The SQL Agent log will not be displayed.', 11, 1);
					SET @output__DisplayAgentLog = 0;
					SET @lv__log__log2processing = -2;
					SELECT @lv__ErrorText = ERROR_MESSAGE(), 
							@lv__ErrorSeverity	= ERROR_SEVERITY(), 
							@lv__ErrorState = ERROR_STATE();
					SET @lv__ErrorText = @lv__ErrorText + ' (Severity: ' + CONVERT(VARCHAR(20),@lv__ErrorSeverity) + ') (State: ' + CONVERT(VARCHAR(20),@lv__ErrorState) + ')'

					RAISERROR( @lv__ErrorText, 11, 1);

					GOTO afteragentlog
				END CATCH

				SET @lv__afterdt = GETDATE();

				IF DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt) > @lv__slownessthreshold
				BEGIN
					SET @lv__ErrorText = N'   ***dbg: post-processing on most-recent non-active SQL Agent log file took ' + CONVERT(NVARCHAR(40),DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt)) + N' milliseconds';
					RAISERROR(@lv__ErrorText, 10, 1) WITH NOWAIT;
				END
			END 

			SET @lv__beforedt = GETDATE();

			--***Location 2: Get second-most-recent non-active log file and process
			BEGIN TRY
				INSERT INTO #SQLAgentLog (LogDate, ErrorLevel, aText)
					EXEC xp_readerrorlog 2,	--log file # (0 is current)
						2,	--SQL Agent log
						null,	--search string 1
						null,	--search string 2
						null,	--search start time
						null,	--search end time
						'Desc'		--order results 
						;
				SET @lv__log__log3processing = 1;
			END TRY
			BEGIN CATCH
				RAISERROR(N'Error occurred while obtaining the second-most-recent non-active SQL Agent log file. The SQL Agent log will not be displayed.', 11, 1);
				SET @output__DisplayAgentLog = 0;
				SET @lv__log__log3processing = -1;
				SELECT @lv__ErrorText = ERROR_MESSAGE(), 
						@lv__ErrorSeverity	= ERROR_SEVERITY(), 
						@lv__ErrorState = ERROR_STATE();
				SET @lv__ErrorText = @lv__ErrorText + ' (Severity: ' + CONVERT(VARCHAR(20),@lv__ErrorSeverity) + ') (State: ' + CONVERT(VARCHAR(20),@lv__ErrorState) + ')'

				RAISERROR( @lv__ErrorText, 11, 1);

				GOTO afteragentlog
			END CATCH

			SET @lv__afterdt = GETDATE();

			IF DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt) > @lv__slownessthreshold
			BEGIN
				SET @lv__ErrorText = N'   ***dbg: obtaining second-most-recent non-active SQL Agent log file took ' + CONVERT(NVARCHAR(40),DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt)) + N' milliseconds';
				RAISERROR(@lv__ErrorText, 10, 1) WITH NOWAIT;
			END
			
			IF @lv__log__log3processing = 1
			BEGIN
				SET @lv__beforedt = GETDATE();

				--There are some SQL Agent log messages that are very common, and are not what we'd typically be looking for. 
				-- make sure these are omitted
				BEGIN TRY
					DELETE 
					FROM #SQLAgentLog 
					WHERE (
						atext LIKE '%Job completion for % is being logged to sysjobhistory%'
						OR atext LIKE '%is being queued for the PowerShell subsystem%'
						OR atext LIKE '%Job % has been requested to run by Schedule%'
						OR atext LIKE '%Saving % for all updated job schedules...'
						OR atext LIKE '% job schedule(s) saved%'
						OR (@DisplayAgentLog = 1 AND atext LIKE '%The Messenger service has not been started - NetSend notifications will not be sent%')
					);

					SELECT @lv__log__maxTabID = MAX(idcol) FROM #SQLAgentLog;

					UPDATE targ 
					SET FileNumber = 1,
						isLastRecord = CASE WHEN idcol <> @lv__log__maxTabID THEN 0 ELSE 1 END
					FROM #SQLAgentLog targ 
					WHERE FileNumber IS NULL 
					;
					SET @lv__log__log3processing = 2;
				END TRY
				BEGIN CATCH
					RAISERROR(N'Error occurred while post-processing the second-most-recent non-active SQL Agent log file. The SQL Agent log will not be displayed.', 11, 1);
					SET @output__DisplayAgentLog = 0;
					SET @lv__log__log3processing = -2;
					SELECT @lv__ErrorText = ERROR_MESSAGE(), 
							@lv__ErrorSeverity	= ERROR_SEVERITY(), 
							@lv__ErrorState = ERROR_STATE();
					SET @lv__ErrorText = @lv__ErrorText + ' (Severity: ' + CONVERT(VARCHAR(20),@lv__ErrorSeverity) + ') (State: ' + CONVERT(VARCHAR(20),@lv__ErrorState) + ')'

					RAISERROR( @lv__ErrorText, 11, 1);

					GOTO afteragentlog
				END CATCH

				SET @lv__afterdt = GETDATE();

				IF DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt) > @lv__slownessthreshold
				BEGIN
					SET @lv__ErrorText = N'   ***dbg: post-processing on second-most-recent non-active SQL Agent log file took ' + CONVERT(NVARCHAR(40),DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt)) + N' milliseconds';
					RAISERROR(@lv__ErrorText, 10, 1) WITH NOWAIT;
				END
			END 
		END		--IF @DisplayAgentLog = 3

		SET @lv__beforedt = GETDATE();
		--***Location 3: Construct output string
		SET @lv__log__AgentLogString = '';

		BEGIN TRY
			IF @DisplayAgentLog >= 2 OR EXISTS (SELECT * FROM #SQLAgentLog l WHERE l.ErrorLevel = 1)
			BEGIN
				SELECT 
					@lv__log__AgentLogString = @lv__log__AgentLogString + 
						REPLACE(CONVERT(NVARCHAR(40),l.LogDate,102),'.','-') + ' ' +
							CONVERT(NVARCHAR(40),l.LogDate,108) + '.' + 
							CONVERT(NVARCHAR(40),DATEPART(millisecond, l.LogDate)) +
						'   ' + 
						CONVERT(NVARCHAR(40),l.ErrorLevel) + 
						'              ' + 
						REPLACE(REPLACE(l.aText,NCHAR(10),N' '),NCHAR(13), N' ') + NCHAR(10) + 
						CASE WHEN isLastRecord = 1 THEN NCHAR(10) ELSE N'' END
				FROM #SQLAgentLog l 
				WHERE l.LogDate <= @lv__mtx__OverallWindowEndTime		--helps implement a "historical" look at the log file for @PointInTime values other than NULL
				ORDER BY idcol ASC; 

				SET @lv__log__AgentLogString = N'<?SQLAgentLog -- ' + NCHAR(10) + 
					N'LogDate                 ErrorLevel     Text' + NCHAR(10) +
					N'-------------------------------------------------------------------------------------' + NCHAR(10) + 
					@lv__log__AgentLogString + NCHAR(10) + N' -- ?>';
			END 
			ELSE 
			BEGIN
				SET @output__DisplayAgentLog = 0;
			END
		END TRY
		BEGIN CATCH
			RAISERROR(N'Error occurred while constructing the SQL Agent log output string. The SQL Agent log will not be displayed.', 11, 1);
			SET @output__DisplayAgentLog = 0;
			SELECT @lv__ErrorText = ERROR_MESSAGE(), 
					@lv__ErrorSeverity	= ERROR_SEVERITY(), 
					@lv__ErrorState = ERROR_STATE();
			SET @lv__ErrorText = @lv__ErrorText + ' (Severity: ' + CONVERT(VARCHAR(20),@lv__ErrorSeverity) + ') (State: ' + CONVERT(VARCHAR(20),@lv__ErrorState) + ')'

			RAISERROR( @lv__ErrorText, 11, 1);

			GOTO afteragentlog
		END CATCH

		SET @lv__afterdt = GETDATE();

		IF DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt) > @lv__slownessthreshold
		BEGIN
			SET @lv__ErrorText = N'   ***dbg: constructing SQL Agent log output string took ' + CONVERT(NVARCHAR(40),DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt)) + N' milliseconds';
			RAISERROR(@lv__ErrorText, 10, 1) WITH NOWAIT;
		END
	END --IF @DisplayAgentLog > 0

afteragentlog: 

	/*********************************************************************************************************************************************
	*********************************************************************************************************************************************

														Part 4: Display the data!

	*********************************************************************************************************************************************
	*********************************************************************************************************************************************/

	--***Location 0: If returning the matrix to the console, print it and set the logic so it isn't returned by the SELECT
	IF @ToConsole = 'Y' AND @output__DisplayMatrix = 1
	BEGIN
		SET @lv__beforedt = GETDATE();
		SET @lv__OutputLength = LEN(@lv__OutputVar);
		SET @lv__CurrentPrintLocation = 1;

		WHILE @lv__CurrentPrintLocation <= @lv__OutputLength
		BEGIN
			PRINT SUBSTRING(@lv__OutputVar, @lv__CurrentPrintLocation, 8000);
			SET @lv__CurrentPrintLocation = @lv__CurrentPrintLocation + 8000;
		END

		SET @output__DisplayMatrix = 0;	--just printed it, don't return it as XML

		SET @lv__afterdt = GETDATE();

		IF DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt) > @lv__slownessthreshold
		BEGIN
			SET @lv__ErrorText = N'   ***dbg: printing job history matrix to console took ' + CONVERT(NVARCHAR(40),DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt)) + N' milliseconds';
			RAISERROR(@lv__ErrorText, 10, 1) WITH NOWAIT;
		END
	END

	SET @lv__beforedt = GETDATE();

	--***Location 1: IF block that controls which SELECT statement is run (i.e. which columns are returned)
	IF @output__DisplayConfig = 0 AND @output__DisplayAgentLog = 0
	BEGIN
		IF @output__DisplayMatrix = 0
		BEGIN
			SET @Debug = @Debug
		END
		ELSE
		BEGIN
			--Only returning the matrix, our output query is very simple
			SELECT CONVERT(XML, @lv__OutputVar) as [Job History Matrix];
		END
	END
	ELSE IF @output__DisplayConfig = 0 AND @output__DisplayAgentLog <> 0
	BEGIN
		IF @output__DisplayMatrix = 0
		BEGIN
			SELECT CONVERT(XML, @lv__log__AgentLogString) as [SQL Agent Error Log]
		END
		ELSE
		BEGIN
			SELECT CONVERT(XML, @lv__OutputVar) as [Job History Matrix], 
				CONVERT(XML, @lv__log__AgentLogString) as [SQL Agent Error Log]
		END
	END
	ELSE
	BEGIN
		--@output__DisplayConfig must be <> 0. Thus, our formatting query is a bit more complex

		IF @output__DisplayAgentLog <> 0
		BEGIN
			--Include SQL Agent Log

			IF @output__DisplayMatrix = 0
			BEGIN
				SELECT 
					[SQL Agent Option] = ss1.OptionTag, 
					[Current Value] = ss1.OptionValue, 
					[Default/Expected Value] = ss1.OptionNormalValue, 
					[SQL Agent Logs] = ISNULL(ss2.[SQL Agent Logs],N'')
				FROM 
					(
					SELECT top 99999 o.idcol, o.OptionTag, o.OptionValue, o.OptionNormalValue
					FROM #OptionsToDisplay o
					ORDER BY o.idcol
					) ss1
					LEFT OUTER JOIN 
						(SELECT 1 as joincol, CONVERT(XML, @lv__log__AgentLogString) as [SQL Agent Logs]) ss2
						ON ss1.idcol = ss2.joincol
				;
			END
			ELSE
			BEGIN
				SELECT [Job History Matrix] = ISNULL(ss0.[Job History Matrix], N''), 
					[SQL Agent Option] = ss1.OptionTag, 
					[Current Value] = ss1.OptionValue, 
					[Default/Expected Value] = ss1.OptionNormalValue, 
					[SQL Agent Logs] = ISNULL(ss2.[SQL Agent Logs],N'')
				FROM 
					(SELECT 1 as joincol, CONVERT(XML, @lv__OutputVar) as [Job History Matrix]) ss0
					RIGHT OUTER JOIN 
						(
						SELECT top 99999 o.idcol, o.OptionTag, o.OptionValue, o.OptionNormalValue
						FROM #OptionsToDisplay o
						ORDER BY o.idcol
						) ss1
						ON ss1.idcol = ss0.joincol
					LEFT OUTER JOIN 
						(SELECT 1 as joincol, CONVERT(XML, @lv__log__AgentLogString) as [SQL Agent Logs]) ss2
						ON ss1.idcol = ss2.joincol
				;
			END
		END
		ELSE
		BEGIN
			--Omit SQL Agent log; 
			IF @output__DisplayMatrix = 0
			BEGIN
				SELECT 
					[SQL Agent Option] = ss1.OptionTag, 
					[Current Value] = ss1.OptionValue, 
					[Default/Expected Value] = ss1.OptionNormalValue
				FROM 
					(
					SELECT top 99999 o.idcol, o.OptionTag, o.OptionValue, o.OptionNormalValue
					FROM #OptionsToDisplay o
					ORDER BY o.idcol
					) ss1
				;
			END
			ELSE
			BEGIN
				SELECT [Job History Matrix] = ISNULL(ss0.[Job History Matrix], N''), 
					[SQL Agent Option] = ss1.OptionTag, 
					[Current Value] = ss1.OptionValue, 
					[Default/Expected Value] = ss1.OptionNormalValue
				FROM 
					(SELECT 1 as joincol, CONVERT(XML, @lv__OutputVar) as [Job History Matrix]) ss0
					RIGHT OUTER JOIN 
						(
						SELECT top 99999 o.idcol, o.OptionTag, o.OptionValue, o.OptionNormalValue
						FROM #OptionsToDisplay o
						ORDER BY o.idcol
						) ss1
						ON ss1.idcol = ss0.joincol
				;
			END
		END
	END 

	SET @lv__afterdt = GETDATE();

	IF DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt) > @lv__slownessthreshold
	BEGIN
		SET @lv__ErrorText = N'   ***dbg: returning final results took ' + CONVERT(NVARCHAR(40),DATEDIFF(MILLISECOND, @lv__beforedt, @lv__afterdt)) + N' milliseconds';
		RAISERROR(@lv__ErrorText, 10, 1) WITH NOWAIT;
	END

	GOTO finishloc

	/*********************************************************************************************************************************************
	*********************************************************************************************************************************************

														Part 5: Print help

	*********************************************************************************************************************************************
	*********************************************************************************************************************************************/

helploc:

	SET @lv__HelpText = 'sp_JobsAtAGlance aggregates information from the SQL Agent catalog in msdb and presents the most relevant information in an
intuitive and actionable format. sp_JobsAtAGlance produces up to 3 different types of output, depending on the parameters chosen:

	1. A matrix or table ("matrix") of job outcomes, either to the console or via a clickable XML result column. This matrix allows the user 
		to quickly view job start and completion times, outcomes, durations, job enable/disable status, and concurrent execution. The intended
		use case is for the user to evaluate a given window of time in order to quickly find potential problems in job scheduling and execution.

	2. A comparison of SQL Agent config options with their standard defaults, allowing the user to quickly determine whether any SQL Agent 
		configuration options varies from the Microsoft installation defaults.

	3. A clickable XML result column that contains relevant records from the SQL Server error log. This output can alert the user to warning
		or error messages in the SQL Agent subsystem that might otherwise go overlooked.

	For additional help content beyond the below documention, please see: <web site link>
	
	COLUMNS RETURNED:
		The number of columns returned by sp_JobsAtAGlance can vary depending on the parameter values passed in and the state of the SQL Agent 
		configuration values and error log. 

		"Job History Matrix"		Contains an XML value that holds the matrix of job outcomes. This column is not returned when @ToConsole="Y", or
									when there are no SQL Agent jobs present on the system; otherwise, it will always be present.

									For a complete legend to the symbols and other content present in this matrix, see the MATRIX LEGEND section below

		"SQL Agent Option"			Contains a list of SQL Agent configuration options (or rather, descriptive labels that identify SQL Agent 
									config options). This field is returned when @DisplayConfigOptions=2, or when @DisplayConfigOptions=1 and 
									there are config options which have been changed from Microsoft installation defaults'

	PRINT @lv__HelpText;

	SET @lv__HelpText = '
		"Current Value"				Contains the SQL Agent config values that correspond to the config option identified in the "SQL Agent Options" 
									field. This column is returned under the same conditions as the "SQL Agent Option" column.

		"Default/Expected Value"	Contains the Microsoft installation value for the config option identified in the "SQL Agent Option" column.
									This column is returned under the same conditions as the "SQL Agent Option" column.

		"SQL Agent Logs"			Contains an XML value that holds a listing of entries in the SQL Agent log file or files. This column is returned 
									when @DisplayAgentLog = 2 or = 3, or when @DisplayAgentLog=1 and Severity 1 error records exist in the most current 
									log file.


	PARAMETERS
		@PointInTime				Specifies, roughly, the end time (the right-hand side) of the job history matrix. If NULL is passed, the current time is 
									used. Any time between 2000-01-01 and 2099-12-31 is allowed, though dates in the past or future may not find any
									job history data to display.

									NOTE: the matrix is always broken up into evenly-sized time blocks (e.g. 5 minutes long). The time blocks are defined
									such that the boundaries are always on the top of the minute, and the block length must be a root of 60 so that
									the top of the hour is also the start of a time block. This means that the @PointInTime value is usually not exactly
									the end of the matrix, but it is guaranteed to fall into the final time block.

									For example, if NULL is passed and GETDATE() returns 12:53, and @HoursBack=24 (resulting in time blocks of 15 minutes),
									then the final time block will be 12:45 to 1:00.

		@HoursBack					Specifies the length of the time window represented by the job history matrix. Valid values are 1 to 48. The default
									of 20 allows for DBAs checking a server in the mid-morning to see back to the start of the batch window the previous
									night. 
									
									The width of the matrix is also affected by the value passed to @HoursBack. For example, since the minimum 
									minute-length of an individual matrix cell is 1 minute, passing @HoursBack=1 will result in a matrix width of 60 
									characters, while an @HoursBack=2 call will result in a matrix width of 120 characters. 
									
									(See also the @FitOnScreen parameter below)'

	PRINT @lv__HelpText;

	SET @lv__HelpText = '
		@ToConsole					If Y, directs the job history matrix to the console (i.e., typically your SSMS "Messages" tab) instead of to an 
									XML value in the result set. Because the fonts are usually smaller for the Messages tab, this often results in
									more info being visible at a time. The default is N, since it keeps all results together on one SSMS tab.

		@FitOnScreen				If "Y" (the default) is chosen, the job history matrix is structured so that it always takes between 90 and 150 
									characters, with the job names placed on the right-hand side. This increases the likelihood that all of the
									matrix data will fit on one screen if the SSMS Object Explorer and Object Properies windows are collapsed.

									If @FitOnScreen=N is specified, the maximum matrix width is 360 characters. Specifying "N" generally gives a
									more detailed view of the data, as the minute-length of an individual matrix cell is smaller and thus more granular.
									"Y" as the default allows for initial, quick review of job outcomes, and then "N" can be specified along with
									a desired @PointInTime value to closely examine a narrower time window.

		@DisplayConfigOptions		If 1 (the default) is chosen, most of the SQL Agent config options are examined and any variance from Microsoft
									installation defaults are presented to the user. If 2 is chosen, all SQL Agent config options that are examined
									by this procedure are returned, regardless of whether they vary from the defaults. If 0 is chosen, no examination
									is done and no config information is returned. '

	PRINT @lv__HelpText

	SET @lv__HelpText = '
		@DisplayAgentLog			If 1 (the default) is chosen, the most recent SQL Agent error log is examined for Severity 1 errors. If any
									of these errors exist, the full error log (except for some very common and benign messages) is returned as a
									clickable XML value. Other possible values: 
										0		the SQL Agent log is not examined and nothing is returned. 
										2		always displays the current log
										3		always displays the 3 most recent SQL Agent logs

									(The decision to only return the 3 most recent logs is due to the fact that there is currently no way within 
									T-SQL to determine the number of SQL Agent log files retained. And because xp_readerrorlog throws an error
									that is uncatchable by TRY...CATCH, walking log file numbers backwards until error is not an attractive option. 
									It would be very surprising to encounter a SQL instance that has had the number of SQL Agent log files retained
									to less than 3).

		@Help						You know what this is.

		@Debug						Returns a variety of printed information to the console and result sets to the user. At this time the Debug
									functionality is not highly-developed.
	'

	PRINT @lv__HelpText

	SET @lv__HelpText = '
	MATRIX LEGEND
		The job history matrix employs a number of symbols and formatting choices to communicate a large quantity of information in a concise,
		quickly-viewable way. 

			- The job history matrix is really 3 sub-matrices. The top matrix is for jobs that have had at least 1 unsuccessful outcome during
				the matrix time window, or are currently running. This draws the users attention to failing or long-running/hung jobs quickly.
				The middle matrix holds jobs that have run at least 1 time in the time window and have always had successful outcomes.
				The bottom matrix lists jobs that have had no runs during the time window, and is primarily useful to look for jobs that
				should have run but did not (e.g. a DB stats update that may have been disabled, leading to system performance problems).
				This final matrix is redundant, as it will always be empty; a future version of this proc may change the formatting.

				The overall time window covered by all 3 sub-matrices is exactly the same.

			- Job Names are printed on the right-hand side of their sub-matrix when @FitOnScreen="Y", so that the matrix is more likely to 
				remain on the screen. This also allows long Job names to be printed without truncation, since it is assumed that most jobs 
				will be identifiable within the first 20-30 characters of the name.
				
			- The @HoursBack parameter is repeated for the benefit of the user, along with its resulting cell-minute-length. 

			- If the last start time for the SQL Server DB engine falls within the time window, a message appears in the header. If 
				SQL Agent last started within the time window, and its start time was not within 1 minute of the DB engine start time,
				a separate message appears in the header. '
	PRINT @lv__HelpText

	SET @lv__HelpText = '
			- The top sub-matrix always starts with 2 header rows: one that lists the hour markers (in military time) and the other that
				lists tick marks and hyphens. The point of both is to aid the user in quickly identifying the timeframe for a given job
				outcome. Above the 2 header rows are the begin and end timestamps of the time window, with precision down to the millisecond.

			- Each sub-matrix is made up of "cells", and each cell represents a time window for a given job. A time window always begins on
				the "00" second and ends on the "59.990" second. If a job has an outcome (success, failure, cancellation, or retry), that
				outcome will be entered into the cell. If a job was executing for the full duration of that cell, an appropriate symbol
				will be entered. 

			- Matrix Cell Hierarchy
				Because a job may start and stop multiple times within a time window (even a small 1 minute time window), the following
				symbols are presented in precedence order. Since a matrix cell is always only 1 character wide, items higher in the list 
				take precedence over lower items. Thus, a job may have succeeded 10 times and retried once in the time window represented 
				by a given cell, but if it also failed at least once in that time window, an "F" will be printed, giving no indication of 
				the # of successes or retries.

				"F"		The job has failed at least once in the time window represented by this cell.

				"R"		The job has retried at least once in this time window.

				"C"		The job has cancelled at least once in this time window.

				"X"		The job has encountered an unexpected msdb.dbo.sysjobhistory.job_run_status value. Research is needed

				"9"		The job has had 9 or more successful completions during the time window.'
	PRINT @lv__HelpText;
	
	SET @lv__HelpText = ' 
				"<number between 2 and 9>"		The job has had this many successful completions during the time window. Since 
						"9" is the largest single-digit number possible, higher numbers of job completions (e.g. 15) are still
						represented as "9"

				"/"		The job has had 1 successful completion in this time window.

				"^"		The job has started once during this time window, but did not complete in the same time window.

				"~" and "!"		If a job is running for a whole time window (i.e. its start occurred in an earlier time window
						and its completion is in a later time window or it has not completed), then one of 2 "running" symbols
						are used. The "~" symbol is the standard; however, if a job exceeds its average duration (obtained by
						inspecting SUCCESSFUL job outcomes in msdb.dbo.sysjobhistory), then time windows beyond the average 
						duration will receive a "!" symbol.

						To be clear, a given job execution may involve both "~" and "!" symbols. For example, a job 
						may have the following characters on its matrix:
								^~~~!!!!!!!/
						This indicates that the job started in one time window, kept executing for 10 more time windows, and
						finally stopped in its 12th time window. If each time window represents 1 minute, then we know that
						its average duration (for successes) is about 4 minutes (give or take a minute), and this run exceeded
						that average starting in its 5th minute.
	'
	PRINT @lv__HelpText;

	GOTO finishloc


finishloc:
	RETURN 0;
END


GO
