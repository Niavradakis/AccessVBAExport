SELECT ProductsT.*, BaseProductsQ.Product_Description AS Base_Product_Description
FROM BaseProductsQ INNER JOIN ProductsT ON BaseProductsQ.Product_ID = ProductsT.Base_Product_ID;

