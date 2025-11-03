PARAMETERS ParProductID Long;
SELECT InventoryPerProductPerStorageQ.Product_ID, Sum(InventoryPerProductPerStorageQ.Total_Debit) AS SumOfTotal_Debit, Sum(InventoryPerProductPerStorageQ.Total_Credit) AS SumOfTotal_Credit
FROM InventoryPerProductPerStorageQ
GROUP BY InventoryPerProductPerStorageQ.Product_ID
HAVING (((InventoryPerProductPerStorageQ.Product_ID)=[parproductid]));

