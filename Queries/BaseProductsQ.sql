SELECT ProductsT.Product_ID, ProductsT.Product_Description, ProductsT.[Marketing Label], ProductsT.Base_Product_ID, ProductsT.VAT_Category_ID, ProductsT.Quantity_Of_Product_Content_In_Package, ProductsT.[Measure_Unit_For_Purchases/Sales_ID], ProductsT.Measure_Unit_Of_Product_Content_In_Package_ID, ProductsT.Product_Type_ID, ProductsT.Product_Category_ID, ProductsT.Product_Subcategory_ID, ProductsT.Barcode_for_internal_use, ProductsT.Barcode_official, ProductsT.SKU, ProductsT.Is_Material, ProductsT.Active, ProductsT.Inventoryable, ProductsT.Notes
FROM ProductsT
WHERE (((ProductsT.Base_Product_ID) Is Null));

