SELECT AttributeValueToProductT.Product_ID
FROM Products‘ INNER JOIN AttributeValueToProductT ON Products‘.Product_ID = AttributeValueToProductT.Product_ID
WHERE (((Products‘.Product_Description) Like "*bio*") AND ((AttributeValueToProductT.Attribute_Value_ID)=91));

