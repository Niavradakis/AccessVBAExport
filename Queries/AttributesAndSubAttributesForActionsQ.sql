SELECT LinkAttributeValueToEntitiesT.*, CStr(Nz([Attribute_Value_String],"")) & CStr(Nz(IIf([EntityTypeID_For_RelevantTablePKField] Is Null,[Attribute_Value_Number],SingleRelatedEntityDescription([EntityTypeID_For_RelevantTablePKField],[Attribute_Value_Number])),"")) & CStr(IIf([Attribute_Value_Boolean]=0,"No",IIf([Attribute_Value_Boolean]=-1,"Yes",""))) & CStr(Nz([Attribute_Value_Date],"")) & CStr(Nz([Attribute_Value_Time],"")) & CStr(Nz([Attribute_Value_TImestamp],"")) AS [Attribute Value], AttributesT.EntityTypeID_For_RelevantTablePKField
FROM AttributesT INNER JOIN LinkAttributeValueToEntitiesT ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID
WHERE (((LinkAttributeValueToEntitiesT.Entity_Type_ID)=8))
ORDER BY AttributesT.Attribute_Description;

