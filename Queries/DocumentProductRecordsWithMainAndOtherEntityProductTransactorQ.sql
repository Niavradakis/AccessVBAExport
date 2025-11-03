SELECT IssuedDocumentProductDetailsT.*, IssuedDocumentT.[Transactor_Product_ID(Main_Entity)], IssuedDocumentT.[Transactor_Product_ID(Other_Entity)], IIf([Main_Entity_Debited(For_Product_Documents)]=-1,[quantity],[quantity]*-1) AS Main_Entity_Quantity, IIf([Main_Entity_Debited(For_Product_Documents)]=0,[quantity],[quantity]*-1) AS Other_Entity_Quantity, IssuedDocumentT.Intention_ID, IssuableDocumentT.Issuable_Document_Description, IssuedDocumentT.Issued_Date
FROM (IssuableDocumentT INNER JOIN IssuedDocumentT ON IssuableDocumentT.Issuable_Document_ID = IssuedDocumentT.Issuable_Document_ID) INNER JOIN IssuedDocumentProductDetailsT ON IssuedDocumentT.Issued_Document_ID = IssuedDocumentProductDetailsT.Issued_Document_ID
WHERE (((IssuableDocumentT.Affects_Inventory)=Yes));

