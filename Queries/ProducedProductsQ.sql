SELECT [ProductsT].Product_ID, [ProductsT].Product_Description, [ProductsT].Produced_From_Recipe
FROM ProductsT
WHERE ((([ProductsT].Produced_From_Recipe)=Yes))
ORDER BY [ProductsT].Product_Description;

