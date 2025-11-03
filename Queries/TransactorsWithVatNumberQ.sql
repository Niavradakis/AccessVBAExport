SELECT TransactorsT.Transactor_ID, TransactorsT.Basic_Transactor_ID, TransactorsT.Transactor_Type_ID, LinkAttributeValueToEntitiesT.Attribute_Value_String AS Vat_Number
FROM TransactorsT INNER JOIN LinkAttributeValueToEntitiesT ON TransactorsT.Transactor_ID = LinkAttributeValueToEntitiesT.Entity_ID
WHERE (((LinkAttributeValueToEntitiesT.Attribute_ID)=189));

