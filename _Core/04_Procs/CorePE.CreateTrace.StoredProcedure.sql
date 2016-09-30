SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [CorePE].[CreateTrace] 
/*   
	PROCEDURE:		CorePE.CreateTrace

	AUTHOR:			Aaron Morelli
					email@TBD.com
					@sqlcrossjoin
					sqlcrossjoin.wordpress.com
					https://github.com/amorelli005/PerformanceEye

	PURPOSE: Creates an entry in the CorePE.Traces table for the @Utility specified. At this time the
		Traces table is little more than a log table to show when various traces have started or stopped. 

		@Utility is either "AutoWho" or "ServerEye" at this point in time.

		@Type is currently always 'Background" (the value is passed in via the AutoWho or ServerEye Executor).
		The intention with this parameter is to separate background traces (i.e. started by the jobs) with
		traces started by some sort of user-facing procedure. In the future, users may be given an interface
		to start/stop instances of a ServerEye trace that collects some or all of the various system DMVs
		and has a given start/stop time. When the trace is complete, the user would be able to navigate that
		data. The value here would be that the user could set specific start & stop times for the trace
		instead of relying on the intervals present with the standard Background ServerEye trace.

		@IntendedStopTime shows when AutoWho or ServerEye *planned* for the trace to stop. It may have stopped
		a few seconds off from that time, or may be many hours off if a human aborted the trace or something
		went wrong. 

	OUTSTANDING ISSUES: None at this time.

    CHANGE LOG:	
				2016-04-24	Aaron Morelli		Final run-through and commenting

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
EXEC CorePE.CreateTrace @Utility=N'', @Type=N'', @IntendedStopTime='2016-04-24 23:59'
*/
(
	@Utility			NVARCHAR(20),
	@Type				NVARCHAR(20),
	@IntendedStopTime	DATETIME=NULL
)
AS
BEGIN
	SET NOCOUNT ON;

	DECLARE @LastIdentity INT,
		@Reraise NVARCHAR(4000);

	IF @Utility IS NULL
	BEGIN
		RAISERROR('Parameter @Utility cannot be null',16,1);
		RETURN -1;
	END

	IF @Type IS NULL
	BEGIN
		RAISERROR('Parameter @Type cannot be null',16,1);
		RETURN -1;
	END

	IF @IntendedStopTime IS NULL
	BEGIN
		RAISERROR('Parameter @IntendedStopTime cannot be null',16,1);
		RETURN -1;
	END
	
	BEGIN TRY
	
		INSERT INTO CorePE.[Traces]
			([Utility], [Type],  --Take default: CreateTime
				IntendedStopTime, StopTime, AbortCode)
		SELECT @Utility,@Type, 
				@IntendedStopTime, NULL, NULL;

		SET @LastIdentity = SCOPE_IDENTITY();

		IF @LastIdentity > 0
		BEGIN
			RETURN @LastIdentity;
		END
		ELSE
		BEGIN
			IF @LastIdentity IS NULL
			BEGIN
				RAISERROR('The output of SCOPE_IDENTITY() is NULL', 16, 1);
				RETURN -1;
			END
			ELSE
			BEGIN
				RAISERROR('The output of SCOPE_IDENTITY() is <= 0', 16, 1);
				RETURN -2;
			END
		END
	END TRY
	BEGIN CATCH
		SET @Reraise = N'Unexpected error occurred while inserting an new trace record into the trace table: Error # ' + 
			CONVERT(NVARCHAR(20),ERROR_NUMBER()) + N'; Severity: ' + 
			CONVERT(NVARCHAR(20),ERROR_SEVERITY()) + '; State: ' + 
			CONVERT(NVARCHAR(20),ERROR_STATE()) + '; Message: '+ ERROR_MESSAGE();

		RAISERROR(@Reraise, 16, 1);
		RETURN -3;
	END CATCH

	RETURN 0;		--should never hit this
END

GO
