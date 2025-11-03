SELECT TransactorsBasicT.Basic_Transactor_Description, TransactorsT.Transactor_ID, TransactorsT.Basic_Transactor_ID, TransactorTypesT.Transactor_Type_Desription, TransactorTypesT.Transactor_Type_ID, TransactorsT.In_Use, TransactorTypesT.Has_VAT_Status, TransactorTypesT.Is_VAT_Related
FROM TransactorTypesT INNER JOIN (TransactorsBasicT INNER JOIN TransactorsT ON TransactorsBasicT.Basic_Transactor_ID = TransactorsT.Basic_Transactor_ID) ON TransactorTypesT.Transactor_Type_ID = TransactorsT.Transactor_Type_ID
WHERE (((TransactorTypesT.Transactor_Type_ID)=12) AND ((TransactorsT.In_Use)=Yes) AND ((TransactorTypesT.Have_Financial_Transactions)=Yes) AND ((TransactorTypesT.[Is_Company's_Entity])=Yes));

