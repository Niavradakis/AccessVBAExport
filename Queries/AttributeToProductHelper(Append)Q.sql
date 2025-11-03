SELECT SalableProductsQ.Product_ID, SalableProductsQ.Product_Description
FROM AttributeValueToProductT RIGHT JOIN SalableProductsQ ON AttributeValueToProductT.Product_ID = SalableProductsQ.Product_ID
WHERE (((SalableProductsQ.Product_Description) Like "*φίλτρ*") AND ((AttributeValueToProductT.Attribute_Value_ID)=91));

