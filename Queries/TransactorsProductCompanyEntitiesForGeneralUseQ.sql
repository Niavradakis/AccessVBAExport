SELECT TransactorsBasicT.Basic_Transactor_Description, TransactorsT.Transactor_ID, TransactorsT.Basic_Transactor_ID, TransactorTypesT.Have_Product_Transactions, TransactorTypesT.[Is_Company's_Entity], TransactorTypesT.Transactor_Type_Desription, TransactorsT.Transactor_Type_ID, TransactorsT.In_Use
FROM TransactorTypesT INNER JOIN (TransactorsBasicT INNER JOIN TransactorsT ON TransactorsBasicT.Basic_Transactor_ID = TransactorsT.Basic_Transactor_ID) ON TransactorTypesT.Transactor_Type_ID = TransactorsT.Transactor_Type_ID
WHERE (((TransactorTypesT.Have_Product_Transactions)=Yes) AND ((TransactorTypesT.[Is_Company's_Entity])=Yes) AND ((TransactorsT.Transactor_Type_ID)=1) AND ((TransactorsT.In_Use)=Yes))
ORDER BY TransactorsBasicT.Basic_Transactor_Description;

