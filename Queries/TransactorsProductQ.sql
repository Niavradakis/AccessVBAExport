SELECT TransactorsBasicT.Basic_Transactor_Description, TransactorsT.Transactor_ID, TransactorsT.Basic_Transactor_ID, TransactorTypesT.Have_Product_Transactions, TransactorTypesT.[Is_Company's_Entity]
FROM TransactorsBasicT INNER JOIN (TransactorTypesT INNER JOIN TransactorsT ON TransactorTypesT.[Transactor_Type_ID] = TransactorsT.Transactor_Type_ID) ON TransactorsBasicT.Basic_Transactor_ID = TransactorsT.Basic_Transactor_ID
WHERE (((TransactorTypesT.Have_Product_Transactions)=Yes));

