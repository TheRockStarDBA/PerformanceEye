SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [CorePE].[AbortTrace] 
/*   
	PROCEDURE:		CorePE.AbortTrace

	AUTHOR:			Aaron Morelli
					email@TBD.com
					@sqlcrossjoin
					sqlcrossjoin.wordpress.com
					https://github.com/amorelli005/PerformanceEye

	PURPOSE: Provides an interface for humans or programs to stop an AutoWho or ServerEye trace. 

	@Utility is either AutoWho or ServerEye at this time.

	@TraceID is the ID # of the trace, or if NULL, the MAX(ID) from the trace table for the Utility value, which
		is likely to be the current trace.

	@PreventAllDay: if "N", aborts the currently-running trace but does not prevent the trace from being started
		up again (e.g. 15 min later by the "Every 15 Minute" PerformanceEye Master job). If "Y", places a row into
		a signal table (e.g. AutoWho.SignalTable or ServerEye.SignalTable) that indicates that the trace
		should not be started up for the rest of the calendar day. 

	OUTSTANDING ISSUES: 
		@PreventAllDay works for the rest of the day, and thus prevents a trace with day boundaries (the default) 
		from starting up. When the next day arrives at midnight, that signal row becomes irrelevant. However,
		the AutoWho and ServerEye traces can be configured to span a day (e.g. 4pm to 3:59am), and in that case,
		the signal that is entered won't stop such a trace from starting back up at 12:00am. 

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
EXEC CorePE.AbortTrace @Utility=N'AutoWho', @TraceID=NULL, @PreventAllDay=N'N'
*/
(
	@Utility NVARCHAR(20),
	@TraceID INT=NULL, 
	@PreventAllDay NCHAR(1) = N'N'
)
AS
BEGIN
	DECLARE @RowExists INT,
		@StopTime DATETIME;

	IF @Utility IS NULL
	BEGIN
		RAISERROR('Parameter @Utility cannot be NULL',16,1);
		RETURN -1;
	END

	IF @TraceID IS NULL
	BEGIN
		--Just go get the most recent one
		SELECT 
			@RowExists = ss.TraceID,
			@StopTime = ss.StopTime 
		FROM (
			SELECT TOP 1 t.TraceID, t.StopTime
			FROM CorePE.[Traces] t WITH (NOLOCK)
			WHERE Utility = @Utility
			ORDER BY t.TraceID DESC
		) ss
	END
	ELSE
	BEGIN
		SELECT @RowExists = t.TraceID,
			@StopTime = t.StopTime
		FROM CorePE.[Traces] t WITH (NOLOCK)
		WHERE Utility = @Utility
		AND t.TraceID = @TraceID;
	END

	IF @RowExists IS NULL
	BEGIN
		RAISERROR('Parameter value for @TraceID not found or no traces exist in the trace table for the @Utility specified',16,1);
		RETURN -1;
	END

	IF @StopTime IS NOT NULL
	BEGIN
		RAISERROR('Parameter value for @TraceID refers to a trace that has already been stopped.',16,1);
		RETURN -1;
	END

	--If we get this far, we have a trace that has not been stopped.
	
	IF @Utility = N'AutoWho'
	BEGIN 
		IF UPPER(ISNULL(@PreventAllDay,N'Z')) NOT IN (N'N',N'Y')
		BEGIN
			RAISERROR('Parameter @PreventAllDay must be either "N" or "Y"', 16, 1);
			RETURN -1;
		END

		IF UPPER(@PreventAllDay) = N'N'
		BEGIN
			INSERT INTO AutoWho.SignalTable 
			(SignalName, SignalValue, InsertTime)
			VALUES (N'AbortTrace', N'OneTime', GETDATE());		-- N'OneTime' --> the Wrapper proc, when it sees this row for the same day,
																-- will abort the loop early and then delete this row so that
																-- the next time it starts it will continue to run
		END
		ELSE
		BEGIN
			INSERT INTO AutoWho.SignalTable 
				(SignalName, SignalValue, InsertTime)
			VALUES (N'AbortTrace', N'AllDay', GETDATE());		-- N'AllDay' --> the Wrapper proc, when it sees this row for the same day,
																-- will abort the loop early, but wil NOT delete this row. Thus, 
																-- that row will prevent this wrapper proc from running the rest of the day
		END
	END

	IF @Utility = N'ServerEye'
	BEGIN 
		IF UPPER(ISNULL(@PreventAllDay,N'Z')) NOT IN (N'N',N'Y')
		BEGIN
			RAISERROR('Parameter @PreventAllDay must be either "N" or "Y"', 16, 1);
			RETURN -1;
		END

		IF UPPER(@PreventAllDay) = N'N'
		BEGIN
			INSERT INTO ServerEye.SignalTable 
			(SignalName, SignalValue, InsertTime)
			VALUES (N'AbortTrace', N'OneTime', GETDATE());		-- N'OneTime' --> the Wrapper proc, when it sees this row for the same day,
																-- will abort the loop early and then delete this row so that
																-- the next time it starts it will continue to run
		END
		ELSE
		BEGIN
			INSERT INTO ServerEye.SignalTable 
				(SignalName, SignalValue, InsertTime)
			VALUES (N'AbortTrace', N'AllDay', GETDATE());		-- N'AllDay' --> the Wrapper proc, when it sees this row for the same day,
																-- will abort the loop early, but wil NOT delete this row. Thus, 
																-- that row will prevent this wrapper proc from running the rest of the day
		END
	END

	--TODO: as other utilities are added, their "Abort Trace" logic goes here

	RETURN 0;
END


GO
