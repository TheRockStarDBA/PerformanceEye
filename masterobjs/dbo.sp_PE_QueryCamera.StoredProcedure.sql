USE [master]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[sp_PE_QueryCamera]
/*   
	PROCEDURE:		sp_PE_QueryCamera

	AUTHOR:			Aaron Morelli
					email@TBD.com
					@sqlcrossjoin
					sqlcrossjoin.wordpress.com
					https://github.com/amorelli005/PerformanceEye

	PURPOSE: 


	FUTURE ENHANCEMENTS: 
		


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

*/
(
	@spid				INT,
	@request			INT=0,
	@frequency			INT=2,			-- every X seconds the capture runs
	@captures			INT=NULL,		-- if NULL or 0, run until the query ends; if > 0, run that many iterations
	@wait				INT=10,
	@allcaptures		NCHAR(1)=N'N',	-- if code is likely to enter sub-calls and then exit back to the main query
										-- (e.g. scalar functions), then specifying Y here tells the code that once
										-- it sees the query, it will execute @capture number of captures regardless
										-- of what it finds.
	@PKSQLStmtStoreID	BIGINT=NULL
)
AS
BEGIN
	SET NOCOUNT ON;

	IF ISNULL(@spid,-1) <= 0 
	BEGIN
		RAISERROR('Parameter @spid must be > 0.', 16, 1);
		RETURN -1;
	END

	IF ISNULL(@request,-1) <= 0 
	BEGIN
		RAISERROR('Parameter @request must be > 0.', 16, 1);
		RETURN -1;
	END

	IF NOT(ISNULL(@frequency,-1) BETWEEN 1 AND 60)
	BEGIN
		RAISERROR('Parameter @frequency must be between 1 and 60.', 16, 1);
		RETURN -1;
	END

	IF @captures IS NULL
	BEGIN
		SET @captures = 0;
	END
	ELSE
	BEGIN
		IF @captures < 0 OR @captures > 1000
		BEGIN
			RAISERROR('Parameter @captures cannot be < 0. Valid values are NULL, 0, or a positive number <= 1000.', 16, 1);
			RETURN -1;
		END
	END

	IF @wait IS NULL
	BEGIN
		SET @wait = 0;
	END
	ELSE
	BEGIN
		IF @wait < 0
		BEGIN
			RAISERROR('Parameter @wait must be NULL or 0 (no wait) or > 0 (wait # of seconds).', 16, 1);
			RETURN -1;
		END
	END

	IF @allcaptures IS NULL
	BEGIN
		RAISERROR('Parameter @allcaptures cannot be NULL.', 16, 1);
		RETURN -1;
	END

	SET @allcaptures = UPPER(@allcaptures);

	IF @allcaptures NOT IN (N'N', N'Y')
	BEGIN
		RAISERROR('Parameter @allcaptures must be either Y or N.', 16, 1);
		RETURN -1;
	END

	IF @allcaptures = N'Y' AND @captures = 0
	BEGIN
		RAISERROR('A positive number for the @captures parameter must be specified if @allcaptures is set to Y.', 16, 1);
		RETURN -1;
	END

	IF @PKSQLStmtStoreID <= 0
	BEGIN
		RAISERROR('Parameter @PKSQLStmtStoreID must be > 0, and should be a valid entry in the AutoWho statement store.', 16, 1);
		RETURN -1;
	END


	EXEC PerformanceEye.AutoWho.QueryCamera @spid=@spid, @request=@request, @frequency=@frequency,
									@captures=@captures, @wait=@wait, 
									@allcaptures = @allcaptures, @PKSQLStmtStoreID = @PKSQLStmtStoreID; 


	RETURN 0;
END
GO
