SELECT
Product_ID,
Transactor_ID,
Debit,
Credit
FROM InventoryBalanceDebitFromMainEntitiesQ

UNION

SELECT
Product_ID,
Transactor_ID,
Debit,
Credit
FROM InventoryBalanceCreditFromMainEntitiesQ

UNION

SELECT
Product_ID,
Transactor_ID,
Debit,
Credit
FROM InventoryBalanceDebitFromOtherEntitiesQ

UNION SELECT
Product_ID,
Transactor_ID,
Debit,
Credit
FROM InventoryBalanceCreditFromOtherEntitiesQ;

