USE [master]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[sp_PE_LongRequests] 
/*   
	PROCEDURE:		sp_PE_Longrequests

	AUTHOR:			Aaron Morelli
					email@TBD.com
					@sqlcrossjoin
					sqlcrossjoin.wordpress.com
					https://github.com/amorelli005/PerformanceEye

	PURPOSE: 

    CHANGE LOG:	2016-05-17  Aaron Morelli		Development Start


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
exec sp_LongRequests @start='2016-05-17 04:00', @end='2016-05-17 06:00', @savespace=N'N'

*/
(
	@start			DATETIME=NULL,			--the start of the time window. If NULL, defaults to 4 hours ago.
	@end			DATETIME=NULL,			-- the end of the time window. If NULL, defaults to 1 second ago.
	@mindur			INT=120,				-- in seconds. Only batch requests with one entry in SAR that is >= this val will be included
	@spids			NVARCHAR(128)=N'',		--comma-separated list of session_ids to include
	@xspids			NVARCHAR(128)=N'',		--comma-separated list of session_ids to exclude
	@dbs			NVARCHAR(512)=N'',		--list of DB names to include
	@xdbs			NVARCHAR(512)=N'',		--list of DB names to exclude
	@attr			NCHAR(1)=N'N',			--Whether to include the session/connection attributes for the request's first entry in sar (in the time range)
	@qplan			NVARCHAR(20)=N'none',		--none / statement		whether to include the query plan for each statement
	@help			NVARCHAR(10)=N'N'		-- "params", "columns", or "all" (anything else <> "N" maps to "all")
)
AS
BEGIN
	SET NOCOUNT ON;
	SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED;
	SET ANSI_PADDING ON;

	DECLARE @scratch__int				INT,
			@helpexec					NVARCHAR(4000),
			@err__msg					NVARCHAR(MAX),
			@DynSQL						NVARCHAR(MAX),
			@helpstr					NVARCHAR(MAX);

	--We always print out the exec syntax (whether help was requested or not) so that the user can switch over to the Messages
	-- tab and see what their options are.
	SET @helpexec = N'
exec sp_PE_LongRequests @start=''<start datetime>'', @end=''<end datetime>'', @mindur=120, 
	@spids=N'''', @xspids=N'''', @dbs=N'''', @xdbs=N'''', 
	@attr=N''N'', @qplan=N''none | statement'',
	@help = N''N | params | columns | all''

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

	SET @helpexec = REPLACE(@helpexec,'<start datetime>', REPLACE(CONVERT(NVARCHAR(20), @start, 102),'.','-') + ' ' + CONVERT(NVARCHAR(20), @start, 108) + '.' + 
													RIGHT(CONVERT(NVARCHAR(20),N'000') + CONVERT(NVARCHAR(20),DATEPART(MILLISECOND, @start)),3)
					);

	IF @end IS NULL
	BEGIN
		SET @end = DATEADD(SECOND,-1, GETDATE());
		RAISERROR('Parameter @end set to 1 second ago because a NULL value was supplied.',10,1) WITH NOWAIT;
	END

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

	IF ISNULL(@attr,N'z') NOT IN (N'N', N'Y')
	BEGIN
		RAISERROR(@helpexec,10,1);
		RAISERROR('Parameter @attr must be either N''N'' or N''Y''.', 16, 1);
		RETURN -1;
	END

	IF ISNULL(@qplan,N'z') NOT IN (N'none', N'statement')
	BEGIN
		RAISERROR(@helpexec,10,1);
		RAISERROR('Parameter @qplan must be either N''none'' or N''statement''.', 16, 1);
		RETURN -1;
	END

	EXEC PerformanceEye.AutoWho.ViewLongBatches @start = @start, 
			@end = @end, 
			@mindur = @mindur,
			@spids = @spids, 
			@xspids = @xspids,
			@dbs = @dbs, 
			@xdbs = @xdbs,
			@attr = @attr,
			@qplan = @qplan
	;

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

	SET @helpstr = @helpexec;
	RAISERROR(@helpstr,10,1) WITH NOWAIT;
	
	IF @Help = N'N'
	BEGIN
		--because the user may want to use sp_SessionViewer and/or sp_QueryProgress next, if they haven't asked for help explicitly, we print out the syntax for 
		--the Session Viewer and Query Progress procedures
		SET @helpstr = '
EXEC sp_PE_SessionViewer @start=''<start datetime>'',@end=''<end datetime>'', --@offset=99999,
	@activity=1, @dur=0,@dbs=N'''',@xdbs=N'''',@spids=N'''',@xspids=N'''',
	@blockonly=N''N'',@attr=N''N'',@resources=N''N'',@batch=N''N'',@plan=N''none'',	--none, statement, full
	@ibuf=N''N'',@bchain=0,@tran=N''N'',@waits=0,		--bchain 0-10, waits 0-3
	@savespace=N''N'',@directives=N''''		--"query(ies)"
	';

		print @helpstr;
		SET @helpstr = '
EXEC dbo.sp_PE_QueryProgress @start=''<start datetime>'',@end=''<end datetime>'', --@offset=99999,	--99999
						@spid=<int>, @request=0, @nodeassociate=N''N'',
						@help=N''N''		--"query(ies)"
		';

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
'

	RAISERROR(@helpstr,10,1);

exitloc:

	RETURN 0;
END

GO
