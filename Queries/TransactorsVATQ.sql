SELECT TransactorsBasicT.Basic_Transactor_Description, TransactorsT.Transactor_ID, TransactorsT.Transactor_Type_ID
FROM TransactorsBasicT INNER JOIN TransactorsT ON TransactorsBasicT.Basic_Transactor_ID = TransactorsT.Basic_Transactor_ID
WHERE (((TransactorsT.Transactor_Type_ID)=10));

