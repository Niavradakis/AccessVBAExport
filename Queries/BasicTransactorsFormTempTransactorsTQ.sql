SELECT BasicTransactorsQ.Basic_Transactor_ID
FROM (SELECT DISTINCT Basic_Transactor_ID FROM TempTransactorsT)  AS BasicTransactorsQ;

