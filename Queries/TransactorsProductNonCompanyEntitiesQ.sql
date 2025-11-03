SELECT TransactorsBasicT.Basic_Transactor_Description, TransactorsT.Transactor_ID AS ID, TransactorsT.Basic_Transactor_ID, TransactorTypesT.Transactor_Type_Desription, TransactorsT.Transactor_Type_ID
FROM TransactorTypesT INNER JOIN (TransactorsBasicT INNER JOIN TransactorsT ON TransactorsBasicT.Basic_Transactor_ID = TransactorsT.Basic_Transactor_ID) ON TransactorTypesT.Transactor_Type_ID = TransactorsT.Transactor_Type_ID
WHERE (((TransactorTypesT.Have_Product_Transactions)=Yes) AND ((TransactorTypesT.[Is_Company's_Entity])=No) AND ((TransactorsT.In_Use)=Yes) AND ((TransactorsBasicT.[In Use])=Yes));

