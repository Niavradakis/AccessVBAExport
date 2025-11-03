SELECT DISTINCT AttributesT.Attribute_Description
FROM AttributesT INNER JOIN (ProductsT INNER JOIN LinkAttributeValueToEntitiesT ON ProductsT.Product_ID = LinkAttributeValueToEntitiesT.Entity_ID) ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID
WHERE (((LinkAttributeValueToEntitiesT.Entity_Type_ID)=1));

