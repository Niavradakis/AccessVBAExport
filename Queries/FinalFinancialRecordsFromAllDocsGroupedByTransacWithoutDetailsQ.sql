SELECT 
IssuedDocumentFinancialDetailsT.Issued_Document_ID, 
IssuedDocumentFinancialDetailsT.Transactor_ID, 
IssuedDocumentFinancialDetailsT.Amount,
IssuedDocumentFinancialDetailsT.[TICK=Debitor-UNTICK=Creditor]



FROM IssuedDocumentFinancialDetailsT


UNION 

SELECT 
[MapProductRecordsToCompany'sFinancialTransactor&VATTransactorQ].Issued_Document_ID, 
[MapProductRecordsToCompany'sFinancialTransactor&VATTransactorQ].[Company's_Financial_Transactor_ID(Main Entity)], 
Sum([MapProductRecordsToCompany'sFinancialTransactor&VATTransactorQ].Total_Net_Value) AS SumOfTotal_Net_Value, 
[MapProductRecordsToCompany'sFinancialTransactor&VATTransactorQ].[Main_Entity_Debited(For_Product_Documents)]

FROM [MapProductRecordsToCompany'sFinancialTransactor&VATTransactorQ]

GROUP BY 
[MapProductRecordsToCompany'sFinancialTransactor&VATTransactorQ].Issued_Document_ID, 
[MapProductRecordsToCompany'sFinancialTransactor&VATTransactorQ].[Company's_Financial_Transactor_ID(Main Entity)], 
[MapProductRecordsToCompany'sFinancialTransactor&VATTransactorQ].[Main_Entity_Debited(For_Product_Documents)];


UNION

SELECT 
[MapProductRecordsToCompany'sFinancialTransactor&VATTransactorQ].Issued_Document_ID, 
[MapProductRecordsToCompany'sFinancialTransactor&VATTransactorQ].[Transactor_Financial_ID(Other_Entity)], 
Sum([MapProductRecordsToCompany'sFinancialTransactor&VATTransactorQ].Total_Value) AS SumOfTotal_Value, 
[MapProductRecordsToCompany'sFinancialTransactor&VATTransactorQ].[Other_Entity_Debited(For_Product_Documents)]

FROM [MapProductRecordsToCompany'sFinancialTransactor&VATTransactorQ]

GROUP BY 
[MapProductRecordsToCompany'sFinancialTransactor&VATTransactorQ].Issued_Document_ID, 
[MapProductRecordsToCompany'sFinancialTransactor&VATTransactorQ].[Transactor_Financial_ID(Other_Entity)], 
[MapProductRecordsToCompany'sFinancialTransactor&VATTransactorQ].[Other_Entity_Debited(For_Product_Documents)];


UNION SELECT 
[MapProductRecordsToCompany'sFinancialTransactor&VATTransactorQ].Issued_Document_ID, 
[MapProductRecordsToCompany'sFinancialTransactor&VATTransactorQ].VAT_Transactor_ID, 
Sum([MapProductRecordsToCompany'sFinancialTransactor&VATTransactorQ].[VAT#]) AS [SumOfVAT#], 
[MapProductRecordsToCompany'sFinancialTransactor&VATTransactorQ].[Main_Entity_Debited(For_Product_Documents)]

FROM [MapProductRecordsToCompany'sFinancialTransactor&VATTransactorQ]

GROUP BY 
[MapProductRecordsToCompany'sFinancialTransactor&VATTransactorQ].Issued_Document_ID, 
[MapProductRecordsToCompany'sFinancialTransactor&VATTransactorQ].VAT_Transactor_ID, 
[MapProductRecordsToCompany'sFinancialTransactor&VATTransactorQ].[Main_Entity_Debited(For_Product_Documents)];

