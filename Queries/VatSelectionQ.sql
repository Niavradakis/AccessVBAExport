SELECT ProductsT.Product_ID, ProductsT.Product_Description, ProductVATCategoriesT.Product_VAT_Category_ID, VAT_T.VAT_ID, VAT_T.[VAT%], TransactorAddressDetailsT.Transactor_Address_Details_ID, VATStatusT.VAT_Status_ID, TransactorAddressDetailsT.Default_Address_Of_Transactor
FROM (TransactorAddressDetailsT INNER JOIN (VATStatusT INNER JOIN (ProductVATCategoriesT INNER JOIN VAT_T ON ProductVATCategoriesT.Product_VAT_Category_ID = VAT_T.Product_VAT_Category_ID) ON VATStatusT.VAT_Status_ID = VAT_T.Transactor_VAT_Status_ID) ON TransactorAddressDetailsT.VAT_Status_ID = VATStatusT.VAT_Status_ID) INNER JOIN ProductsT ON ProductVATCategoriesT.Product_VAT_Category_ID = ProductsT.VAT_Category_ID
WHERE (((TransactorAddressDetailsT.Default_Address_Of_Transactor)=Yes));

