SELECT [ProductsT].Product_ID, [ProductsT].Product_Description, [ProductsT].Consumable_In_Recipe
FROM ProductsT
WHERE ((([ProductsT].Consumable_In_Recipe)=Yes));

