USE PerformanceEye
GO
CREATE TRIGGER [CorePE].[trgUPD_CorePEVersion] ON [CorePE].[Version]

FOR UPDATE
AS 	BEGIN

INSERT INTO CorePE.Version_History 
([Version], 
EffectiveDate, 
HistoryInsertDate,
TriggerAction)
SELECT 
[Version], 
EffectiveDate, 
getdate(),
'Update'
FROM inserted
END
GO
