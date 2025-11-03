SELECT TransactorsBasicT.Basic_Transactor_Description, TransactorsT.Transactor_ID, TransactorsT.Basic_Transactor_ID, TransactorTypesT.Transactor_Type_ID, TransactorTypesT.Transactor_Type_Desription, TransactorsT.In_Use, TransactorTypesT.Has_VAT_Status
FROM TransactorTypesT INNER JOIN (TransactorsBasicT INNER JOIN TransactorsT ON TransactorsBasicT.Basic_Transactor_ID = TransactorsT.Basic_Transactor_ID) ON TransactorTypesT.Transactor_Type_ID = TransactorsT.Transactor_Type_ID
WHERE (((TransactorTypesT.Have_Financial_Transactions)=Yes));

