SELECT LinkAttributeValueToEntitiesT.*, AttributesT.Attribute_Description
FROM AttributesT INNER JOIN LinkAttributeValueToEntitiesT ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID
WHERE (((AttributesT.Attribute_Description) Is Not Null) AND ((LinkAttributeValueToEntitiesT.Entity_Type_ID)=2));

