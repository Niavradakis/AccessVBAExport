SELECT LinkSubAttributeValueToAttributeValuesQ.Entity_ID
FROM (SELECT LinkAttributeValueToEntitiesT.*, LinkAttributeValueToEntitiesT.Entity_Type_ID
FROM LinkAttributeValueToEntitiesT
WHERE (((LinkAttributeValueToEntitiesT.Entity_Type_ID)=7)))  AS LinkSubAttributeValueToAttributeValuesQ LEFT JOIN LinkAttributeValueToEntitiesT ON LinkSubAttributeValueToAttributeValuesQ.Entity_ID = LinkAttributeValueToEntitiesT.Link_Attribute_Value_To_Entity_ID
WHERE (((LinkAttributeValueToEntitiesT.Entity_ID) Is Null));

