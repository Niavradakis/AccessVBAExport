SELECT IssuedDocumentFinancialDetailsT.Transactor_ID
FROM IssuedDocumentT INNER JOIN IssuedDocumentFinancialDetailsT ON IssuedDocumentT.Issued_Document_ID = IssuedDocumentFinancialDetailsT.Issued_Document_ID
WHERE (((IssuedDocumentT.New_DATA_OR_Edit_DATA)=True));

