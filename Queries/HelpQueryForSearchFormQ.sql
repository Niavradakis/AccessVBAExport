SELECT IssuedDocumentBackupT.Issued_Document_ID, Format(DSum("IssuedDocumentFinancialDetailsT.debit","IssuedDocumentFinancialDetailsT","IssuedDocumentFinancialDetailsT.Issued_Document_ID = " & [Issued_Document_ID] & "AND IssuedDocumentFinancialDetailsT.Transactor_ID " & IIf(IsNull([Forms]![DocumentSearchF].[TransactorIDCbo])," LIKE ""*"" "," = " & [Forms]![DocumentSearchF].[TransactorIDCbo])),"Fixed") AS [вяеысг сум/моу]
FROM IssuedDocumentBackupT;

