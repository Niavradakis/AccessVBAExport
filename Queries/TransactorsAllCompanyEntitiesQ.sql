SELECT TransactorsT.Transactor_ID, TransactorsBasicT.Basic_Transactor_Description, TransactorsT.Basic_Transactor_ID, TransactorTypesT.[Is_Company's_Entity], TransactorsT.Transactor_Type_ID
FROM TransactorTypesT INNER JOIN (TransactorsBasicT INNER JOIN TransactorsT ON TransactorsBasicT.Basic_Transactor_ID = TransactorsT.Basic_Transactor_ID) ON TransactorTypesT.Transactor_Type_ID = TransactorsT.Transactor_Type_ID
WHERE (((TransactorTypesT.[Is_Company's_Entity])=Yes));

