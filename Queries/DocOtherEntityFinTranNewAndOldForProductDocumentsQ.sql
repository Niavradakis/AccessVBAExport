PARAMETERS IssuedDocumentIDPar Long;
SELECT IssuedDocumentT.IS_Deleted AS NewIsDeleted, IssuedDocumentBackupT.Is_Deleted AS OldIsDeleted, IssuedDocumentT.Issued_Document_Backup_ID AS FKBackupID, IssuedDocumentBackupT.IssuedDocumentID_Backup_ID AS BackupID, IssuedDocumentT.[Transactor_Financial_ID(Other_Entity)] AS NewOtherEntityFinancialTransactorID, IssuedDocumentBackupT.[Transactor_Financial_ID(Other_Entity)] AS OldOtherEntityFinancialTransactorID, IssuedDocumentT.Issued_Document_ID
FROM (TransactionsT INNER JOIN IssuedDocumentT ON TransactionsT.Transaction_ID = IssuedDocumentT.Transaction_ID) LEFT JOIN IssuedDocumentBackupT ON IssuedDocumentT.Issued_Document_Backup_ID = IssuedDocumentBackupT.IssuedDocumentID_Backup_ID
WHERE (((IssuedDocumentBackupT.Is_Deleted)=False Or (IssuedDocumentBackupT.Is_Deleted) Is Null) AND ((IssuedDocumentT.Issued_Document_ID)=[IssuedDocumentIDPar]));

