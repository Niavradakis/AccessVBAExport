SELECT RecipesUniqueProducedProductsQ.Sale_Product_ID, Products‘.Product_Description, Products‘.Produced_From_Recipe
FROM RecipesUniqueProducedProductsQ LEFT JOIN Products‘ ON RecipesUniqueProducedProductsQ.Sale_Product_ID = Products‘.Product_ID
WHERE (((Products‘.Produced_From_Recipe)=False));

