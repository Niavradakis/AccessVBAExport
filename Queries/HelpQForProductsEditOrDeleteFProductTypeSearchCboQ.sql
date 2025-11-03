SELECT ProductsT.*, ProductsT.Product_Type_ID
FROM ProductsT
WHERE (((ProductsT.Product_Type_ID)=tempvars!ProductTypeToSearch)) Or (((tempvars!ProductTypeToSearch) Is Null));

