SELECT Products‘.Product_ID, Products‘.Product_Description, Products‘.Produced_From_Recipe, ProductToSupplierLink‘.Supplier_ID
FROM Products‘ LEFT JOIN ProductToSupplierLink‘ ON Products‘.Product_ID = ProductToSupplierLink‘.Product_ID
WHERE (((Products‘.Produced_From_Recipe)=False) AND ((ProductToSupplierLink‘.Supplier_ID) Is Null));

