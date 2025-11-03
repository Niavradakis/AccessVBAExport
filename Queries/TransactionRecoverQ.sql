PARAMETERS TransactionBackupID Long, TransactionTypeID_ValueToRecover Long, IsDeleted_ValueToRecover Bit;
UPDATE TransactionsT SET TransactionsT.Transaction_Type_ID = [TransactionTypeID_ValueToRecover], TransactionsT.Is_Deleted = [IsDeleted_ValueToRecover], TransactionsT.Transaction_Backup_ID = Null
WHERE (((TransactionsT.Transaction_Backup_ID)=[TransactionBackupID]));

