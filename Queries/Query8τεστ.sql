SELECT TransactionsT.Transaction_ID, IssuedDocumentT.Issued_Document_ID
FROM TransactionsT INNER JOIN IssuedDocumentT ON TransactionsT.Transaction_ID = IssuedDocumentT.Transaction_ID;

