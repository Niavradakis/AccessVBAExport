SELECT DISTINCT LinkAttributeValueToEntitiesT.Entity_Type_ID, LinkAttributeValueToEntitiesT.Attribute_ID, LinkAttributeValueToEntitiesT.Entity_ID
FROM AttributesT INNER JOIN LinkAttributeValueToEntitiesT ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID
WHERE (((LinkAttributeValueToEntitiesT.Entity_Type_ID)=1) And ((AttributesT.Entities_May_Have_Multiple_Values)=No) And ((LinkAttributeValueToEntitiesT.Entity_ID)=tempvars!TempVarEntityID));

