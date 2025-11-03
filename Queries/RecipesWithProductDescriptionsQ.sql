SELECT Recipes‘.*, Products‘.Product_Description, Products‘_1.Product_Description
FROM Products‘ INNER JOIN (Products‘ AS Products‘_1 INNER JOIN Recipes‘ ON Products‘_1.Product_ID = Recipes‘.Consumable_Product_ID) ON Products‘.Product_ID = Recipes‘.Sale_Product_ID;

