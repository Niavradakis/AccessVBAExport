SELECT IssuedDocumentProductDetailsT.Quantity, IssuedDocumentProductDetailsT.Unit_Price_Before_Discount, IssuedDocumentT.[Transactor_Financial_ID(Other_Entity)], IssuedDocumentProductDetailsT.Product_ID, IssuedDocumentT.Issued_Document_ID, IssuedDocumentT.Issued_Date
FROM IssuedDocumentT INNER JOIN IssuedDocumentProductDetailsT ON (IssuedDocumentT.Issued_Document_ID = IssuedDocumentProductDetailsT.Issued_Document_ID) AND (IssuedDocumentT.Issued_Document_ID = IssuedDocumentProductDetailsT.Issued_Document_ID)
WHERE (((IssuedDocumentT.[Transactor_Financial_ID(Other_Entity)])=167) AND ((IssuedDocumentProductDetailsT.Product_ID)=1840))
ORDER BY IssuedDocumentT.Issued_Date DESC;

