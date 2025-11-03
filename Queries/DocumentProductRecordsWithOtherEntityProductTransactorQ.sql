SELECT IssuedDocumentProductDetailsT.*, IssuedDocumentT.[Financial_Transaction_Point_ID(Main_Entity)], IssuedDocumentT.[Transactor_Product_ID(Other_Entity)], IIf([Main_Entity_Debited(For_Product_Documents)]=0,-1,0) AS [TICK=Debitor-UNTICK=Creditor], IssuableDocumentT.Affects_Inventory, False AS Is_Financial, True AS iIs_Product
FROM (IssuableDocumentT INNER JOIN IssuedDocumentT ON IssuableDocumentT.Issuable_Document_ID = IssuedDocumentT.Issuable_Document_ID) INNER JOIN IssuedDocumentProductDetailsT ON IssuedDocumentT.Issued_Document_ID = IssuedDocumentProductDetailsT.Issued_Document_ID
WHERE (((IssuableDocumentT.Affects_Inventory)=Yes));

