SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TRIGGER [CorePE].[trgINS_CorePEVersion] ON [CorePE].[Version]

FOR INSERT
AS 	BEGIN

INSERT INTO CorePE.Version_History
([Version], 
EffectiveDate, 
HistoryInsertDate,
TriggerAction)
SELECT 
[Version], 
EffectiveDate, 
GETDATE(),
'Insert'
FROM inserted
END
GO

