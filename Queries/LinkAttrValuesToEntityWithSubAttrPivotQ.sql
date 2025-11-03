SELECT LinkAttributeValueToEntitiesT.*
FROM AttributesT INNER JOIN LinkAttributeValueToEntitiesT ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID
ORDER BY AttributesT.Attribute_Description;

