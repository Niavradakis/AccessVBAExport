SELECT LinkAttributeValueToEntitiesT.*, AttributesT.Attribute_Description, LinkAttributeValueToEntitiesT.Entity_ID AS LATPQ_Entity_ID
FROM AttributesT INNER JOIN LinkAttributeValueToEntitiesT ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID
WHERE (((LinkAttributeValueToEntitiesT.Entity_Type_ID)=1))
ORDER BY AttributesT.Attribute_Description;

