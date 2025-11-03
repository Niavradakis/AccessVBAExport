SELECT TransactorsT.*, TransactorsBasicT.Basic_Transactor_Description, TransactorTypesT.Transactor_Type_Desription
FROM TransactorTypesT INNER JOIN (TransactorsBasicT RIGHT JOIN TransactorsT ON TransactorsBasicT.Basic_Transactor_ID = TransactorsT.Basic_Transactor_ID) ON TransactorTypesT.Transactor_Type_ID = TransactorsT.Transactor_Type_ID;

