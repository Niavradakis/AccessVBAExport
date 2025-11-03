SELECT ProducedProductsQ.Product_ID, ProducedProductsQ.Product_Description
FROM ProducedProductsQ LEFT JOIN RecipesUniqueProducedProductsQ ON ProducedProductsQ.Product_ID = RecipesUniqueProducedProductsQ.Sale_Product_ID
WHERE (((RecipesUniqueProducedProductsQ.Sale_Product_ID) Is Null));

