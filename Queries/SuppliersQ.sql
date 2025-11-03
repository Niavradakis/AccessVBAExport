SELECT TransactorsT.Transactor_ID, TransactorsBasicT.Basic_Transactor_Description, TransactorsT.Basic_Transactor_ID, TransactorsT.Transactor_Type_ID, TransactorsT.Account_ID, TransactorsT.Notes
FROM TransactorsBasicT RIGHT JOIN TransactorsT ON TransactorsBasicT.Basic_Transactor_ID = TransactorsT.Basic_Transactor_ID
WHERE (((TransactorsT.Transactor_Type_ID)=5 Or (TransactorsT.Transactor_Type_ID)=6));

