SELECT DISTINCT LinkAttributeValueToEntitiesT.Attribute_ID, LinkAttributeValueToEntitiesT.Entity_Type_ID, TransactorsT.Transactor_Type_ID, TransactorsT.Transactor_ID
FROM AttributesT INNER JOIN (LinkAttributeValueToEntitiesT LEFT JOIN TransactorsT ON LinkAttributeValueToEntitiesT.Entity_ID = TransactorsT.Transactor_ID) ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID
WHERE (((LinkAttributeValueToEntitiesT.Entity_Type_ID)=2) AND ((TransactorsT.Transactor_Type_ID)=7) AND ((AttributesT.Entities_May_Have_Multiple_Values)=No));

