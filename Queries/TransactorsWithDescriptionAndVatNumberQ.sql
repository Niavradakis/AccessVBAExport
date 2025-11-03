SELECT TransactorsT.*, TransactorsBasicT.Basic_Transactor_Description, TransactorTypesT.Transactor_Type_Desription, LinkAttributeValueToEntitiesT.Attribute_Value_String AS Vat_Number
FROM (TransactorTypesT INNER JOIN (TransactorsBasicT RIGHT JOIN TransactorsT ON TransactorsBasicT.Basic_Transactor_ID = TransactorsT.Basic_Transactor_ID) ON TransactorTypesT.Transactor_Type_ID = TransactorsT.Transactor_Type_ID) LEFT JOIN LinkAttributeValueToEntitiesT ON TransactorsT.Transactor_ID = LinkAttributeValueToEntitiesT.Entity_ID
WHERE (((LinkAttributeValueToEntitiesT.Attribute_ID)=189));

