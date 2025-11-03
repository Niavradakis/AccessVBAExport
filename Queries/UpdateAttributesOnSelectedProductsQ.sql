SELECT SelectProductsToUpdateAttributesQ.Product_ID, AttributeValueToProductT.Attribute_Value_ID, 90 AS [new attribute id]
FROM AttributeValueToProductT INNER JOIN SelectProductsToUpdateAttributesQ ON AttributeValueToProductT.Product_ID = SelectProductsToUpdateAttributesQ.Product_ID
WHERE (((AttributeValueToProductT.Attribute_Value_ID)=91));

