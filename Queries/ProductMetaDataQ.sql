SELECT ProductMetaDataT.Product_ID, Products‘.[Marketing Label], ProductMetaDataT.[Description of product ingredients for efood]
FROM Products‘ INNER JOIN ProductMetaDataT ON Products‘.Product_ID = ProductMetaDataT.Product_ID;

