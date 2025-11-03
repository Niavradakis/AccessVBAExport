SELECT LinkAttributeValueToEntitiesT.*, AttributesT.Attribute_Description, ProductsT.*
FROM AttributesT INNER JOIN (LinkAttributeValueToEntitiesT LEFT JOIN ProductsT ON LinkAttributeValueToEntitiesT.Entity_ID = ProductsT.Product_ID) ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID
WHERE (((LinkAttributeValueToEntitiesT.Entity_Type_ID)=1))
ORDER BY AttributesT.Attribute_Description;

