USE master 
GO
DECLARE @DBN_input NVARCHAR(256),
		@DatabaseName NVARCHAR(256),
		@DynSQL NVARCHAR(4000);
SET @DBN_input = '$(DBName)';

IF @DBN_input IS NULL
BEGIN
	RAISERROR('Parameter "Database" cannot be null.', 16,1)
END
ELSE
BEGIN
	IF NOT EXISTS (SELECT * FROM sys.databases d
					WHERE d.name = @DBN_input)
	BEGIN
		SET @DynSQL = N'CREATE DATABASE ' + @DBN_input;
		EXEC (@DynSQL);
	END
END
