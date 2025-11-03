PARAMETERS currenttimestamp DateTime, currentuserID Short, CurrentTransactionID Short;
INSERT INTO TransactionsBackupT ( Transaction_ID, Transaction_Type_ID, Is_Deleted, Transaction_Edit_Timestamp, Transaction_Edit_User_ID )
SELECT TransactionsT.Transaction_ID, TransactionsT.Transaction_Type_ID, TransactionsT.Is_Deleted, [currenttimestamp] AS Transaction_Edit_Timestamp, [currentuserid] AS Transaction_Edit_User_ID
FROM TransactionsT
WHERE (((TransactionsT.Transaction_ID)=[currenttransactionid]));

