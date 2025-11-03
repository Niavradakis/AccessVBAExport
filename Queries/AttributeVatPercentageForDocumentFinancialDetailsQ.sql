SELECT LinkAttributeValueToEntitiesT.*, LinkAttributeValueToEntitiesT.Attribute_Value_Number AS VatPercentage
FROM LinkAttributeValueToEntitiesT
WHERE (((LinkAttributeValueToEntitiesT.Entity_Type_ID)=4) AND ((LinkAttributeValueToEntitiesT.Attribute_ID)=16));

