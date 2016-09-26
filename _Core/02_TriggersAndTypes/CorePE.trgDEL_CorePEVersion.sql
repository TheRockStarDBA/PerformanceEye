USE PerformanceEye
GO
CREATE TRIGGER [CorePE].[trgDEL_CorePEVersion] ON [CorePE].[Version]

FOR DELETE
AS 	BEGIN

INSERT INTO CorePE.Version_History 
([Version], 
EffectiveDate, 
HistoryInsertDate, 
TriggerAction)
SELECT 
Version, 
EffectiveDate, 
getdate(),
'Delete'
FROM deleted
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO