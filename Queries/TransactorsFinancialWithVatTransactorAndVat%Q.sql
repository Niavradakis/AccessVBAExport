SELECT [Basic_Transactor_Description] & " " & Format([Vat%],'Percent') AS Transactor_Description, TransactorsFinancialQ.Basic_Transactor_Description, TransactorsFinancialQ.Transactor_ID, TransactorsFinancialQ.Basic_Transactor_ID, TransactorsFinancialQ.Transactor_Type_ID, TransactorsFinancialQ.Transactor_Type_Desription, TransactorsFinancialQ.In_Use, TransactorsFinancialQ.Has_VAT_Status, [LinkFinTransactorsToVatTransactorsWithVat%Q].Vat_Transactor_ID, [LinkFinTransactorsToVatTransactorsWithVat%Q].[Vat%]
FROM TransactorsFinancialQ LEFT JOIN [LinkFinTransactorsToVatTransactorsWithVat%Q] ON TransactorsFinancialQ.Transactor_ID = [LinkFinTransactorsToVatTransactorsWithVat%Q].Financial_Transactor_ID
ORDER BY [LinkFinTransactorsToVatTransactorsWithVat%Q].[Vat%] DESC;

