SELECT DISTINCT TransactorsBasicT.Basic_Transactor_Description, TransactorsT.Transactor_ID AS ID, TransactorTypesT.Transactor_Type_Desription, TransactorsT.Transactor_Type_ID, TransactorsT.In_Use, TransactorsBasicT.[In Use], TransactorTypesT.[Is_Company's_Entity], TransactorTypesT.Have_Financial_Transactions
FROM TransactorTypesT RIGHT JOIN (TransactorsBasicT RIGHT JOIN TransactorsT ON TransactorsBasicT.Basic_Transactor_ID = TransactorsT.Basic_Transactor_ID) ON TransactorTypesT.Transactor_Type_ID = TransactorsT.Transactor_Type_ID
WHERE (((TransactorTypesT.[Is_Company's_Entity])=No) AND ((TransactorTypesT.Have_Financial_Transactions)=Yes))
ORDER BY TransactorsT.Transactor_ID, TransactorTypesT.Transactor_Type_Desription;

