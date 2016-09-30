SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TRIGGER [AutoWho].[trgDEL_AutoWhoOptions] ON [AutoWho].[Options]

FOR DELETE
AS 	BEGIN

--We don't actually allow deletes. So call the reset procedure
RAISERROR('Deletes on the Options table instead cause the options table to be reset.',10,1);
ROLLBACK TRANSACTION;

EXEC AutoWho.ResetOptions;

RETURN;

END
GO
