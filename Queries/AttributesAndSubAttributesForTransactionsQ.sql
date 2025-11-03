SELECT LinkAttributeValueToEntitiesT.*
FROM AttributesT INNER JOIN LinkAttributeValueToEntitiesT ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID
WHERE (((LinkAttributeValueToEntitiesT.Entity_Type_ID)=6))
ORDER BY AttributesT.Attribute_Description DESC;

