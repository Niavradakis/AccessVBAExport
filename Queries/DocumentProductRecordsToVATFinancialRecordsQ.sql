SELECT [MapProductRecordsToCompany'sFinancialTransactor&VATTransactorQ].Issued_Document_ID, [MapProductRecordsToCompany'sFinancialTransactor&VATTransactorQ].VAT_Transactor_ID, Sum([MapProductRecordsToCompany'sFinancialTransactor&VATTransactorQ].[VAT#]) AS [SumOfVAT#], [MapProductRecordsToCompany'sFinancialTransactor&VATTransactorQ].[Main_Entity_Debited(For_Product_Documents)]
FROM [MapProductRecordsToCompany'sFinancialTransactor&VATTransactorQ]
GROUP BY [MapProductRecordsToCompany'sFinancialTransactor&VATTransactorQ].Issued_Document_ID, [MapProductRecordsToCompany'sFinancialTransactor&VATTransactorQ].VAT_Transactor_ID, [MapProductRecordsToCompany'sFinancialTransactor&VATTransactorQ].[Main_Entity_Debited(For_Product_Documents)];

