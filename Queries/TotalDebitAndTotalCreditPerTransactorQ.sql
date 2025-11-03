SELECT IssuedDocumentFinancialDetailsT.Transactor_ID, Sum(IssuedDocumentFinancialDetailsT.Debit) AS SumOfDebit, Sum(IssuedDocumentFinancialDetailsT.Credit) AS SumOfCredit
FROM TransactionsT INNER JOIN (TransactorsT INNER JOIN (IssuedDocumentT INNER JOIN IssuedDocumentFinancialDetailsT ON IssuedDocumentT.Issued_Document_ID = IssuedDocumentFinancialDetailsT.Issued_Document_ID) ON TransactorsT.Transactor_ID = IssuedDocumentFinancialDetailsT.Transactor_ID) ON TransactionsT.Transaction_ID = IssuedDocumentT.Transaction_ID
WHERE (((TransactionsT.Is_Deleted)=False) AND ((IssuedDocumentFinancialDetailsT.Is_Deleted)=False) AND ((IssuedDocumentT.IS_Deleted)=False))
GROUP BY IssuedDocumentFinancialDetailsT.Transactor_ID;

