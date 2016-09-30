DECLARE @HoursToKeep_str NVARCHAR(256);
DECLARE @HoursToKeep INT;

SET @HoursToKeep_str = '$(HoursToKeep)';
SET @HoursToKeep = CONVERT(INT, @HoursToKeep_str); 

EXEC AutoWho.InsertConfigData @HoursToKeep = @HoursToKeep;

