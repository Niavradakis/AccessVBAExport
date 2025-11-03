SELECT TransactionsT.Transaction_ID, TransactionsT.Transaction_Type_ID
FROM IntentionsTypeT INNER JOIN TransactionsT ON IntentionsTypeT.Intention_Type_ID = TransactionsT.Transaction_Type_ID;

