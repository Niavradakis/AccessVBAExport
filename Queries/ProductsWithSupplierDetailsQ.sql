SELECT Products‘.Product_Description, ProductToSupplierLink‘.Product_ID, TransactorsBasicT.Basic_Transactor_Description, SuppliersQ.Transactor_ID, TransactorsBasicT.[In Use]
FROM Products‘ INNER JOIN ((SuppliersQ INNER JOIN ProductToSupplierLink‘ ON SuppliersQ.Transactor_ID = ProductToSupplierLink‘.Supplier_ID) INNER JOIN TransactorsBasicT ON SuppliersQ.Basic_Transactor_ID = TransactorsBasicT.Basic_Transactor_ID) ON Products‘.Product_ID = ProductToSupplierLink‘.Product_ID;

