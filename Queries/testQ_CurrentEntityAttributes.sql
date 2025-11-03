SELECT DISTINCT LinkAttributeValueToEntitiesT.Attribute_ID, IssuedDocumentT.Issuable_Document_ID, AttributesT.Attribute_Description
FROM AttributesT INNER JOIN (LinkAttributeValueToEntitiesT INNER JOIN IssuedDocumentT ON LinkAttributeValueToEntitiesT.Issued_Document_ID_For_Temp_Info = IssuedDocumentT.Issued_Document_ID) ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID
WHERE (((LinkAttributeValueToEntitiesT.Entity_Type_ID)=3) AND ((AttributesT.Entities_May_Have_Multiple_Values)=No));

