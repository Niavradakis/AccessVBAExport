SELECT DISTINCT AttributesT.Attribute_Description, LinkAttributeValueToEntitiesT.Attribute_ID
FROM (SELECT
		DISTINCT LinkAttributeValueToEntitiesT.Attribute_ID,
		IssuedDocumentFinancialDetailsFinalT.Transactor_ID,
		AttributesT.Attribute_Description
	FROM
		AttributesT
	INNER JOIN (LinkAttributeValueToEntitiesT
	LEFT JOIN IssuedDocumentFinancialDetailsFinalT ON
		LinkAttributeValueToEntitiesT.Entity_ID = IssuedDocumentFinancialDetailsFinalT.Issued_Document_Financial_Details_ID) ON
		AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID
	WHERE
		LinkAttributeValueToEntitiesT.Entity_Type_ID = 4
		AND IssuedDocumentFinancialDetailsFinalT.Issued_Document_Financial_Details_ID = 1557
		AND AttributesT.Entities_May_Have_Multiple_Values = No)  AS CurrentEntityAttributesQ RIGHT JOIN (((IssuedDocumentFinancialDetailsFinalT INNER JOIN TransactorsT ON IssuedDocumentFinancialDetailsFinalT.Transactor_ID = TransactorsT.Transactor_ID) LEFT JOIN LinkAttributeValueToEntitiesT ON IssuedDocumentFinancialDetailsFinalT.Issued_Document_Financial_Details_ID = LinkAttributeValueToEntitiesT.Entity_ID) LEFT JOIN AttributesT ON LinkAttributeValueToEntitiesT.Attribute_ID = AttributesT.Attribute_ID) ON CurrentEntityAttributesQ.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID
WHERE (((CurrentEntityAttributesQ.Attribute_ID) Is Null)
		AND ((LinkAttributeValueToEntitiesT.Entity_Type_ID)= 4)
			AND TransactorsT.Transactor_Type_ID =7);

