Option Compare Database

Public Sub keycodeNumberFinder()
KeyCode = vbKeyInsert
Debug.Print KeyCode
End Sub

Public Sub EchoChange()
Application.Echo True
End Sub

Public Sub DebugPrintTests()
Debug.Print "SELECT ProductsT.*, ProductsT_1.Product_Description AS BaseProductDescription, ProductVATCategoriesT.Product_VAT_Category_Description, " & _
"MeasureUnitsT.Measure_Unit_Abbreviation AS MUForPurchase, MeasureUnitsT_1.Measure_Unit_Abbreviation AS MUAbbreviationOfPackageContent, ProductTypeT.Type_Description, " & _
"ProductCategoryT.Category_Description, ProductSubcategoryT.Subcategory_Description " & _
"FROM ProductSubcategoryT RIGHT JOIN (ProductCategoryT RIGHT JOIN (ProductTypeT RIGHT JOIN " & _
"((MeasureUnitsT RIGHT JOIN (ProductVATCategoriesT RIGHT JOIN (ProductsT AS ProductsT_1 LEFT JOIN " & _
"ProductsT ON ProductsT_1.Product_ID = ProductsT.Base_Product_ID) ON ProductVATCategoriesT.Product_VAT_Category_ID = ProductsT.VAT_Category_ID) " & _
"ON MeasureUnitsT.Measure_Unit_ID = ProductsT.[Measure_Unit_For_Purchases/Sales_ID]) LEFT JOIN " & _
"MeasureUnitsT AS MeasureUnitsT_1 ON ProductsT.Measure_Unit_Of_Product_Content_In_Package_ID = MeasureUnitsT_1.Measure_Unit_ID) " & _
"ON ProductTypeT.Type_ID = ProductsT.Product_Type_ID) ON ProductCategoryT.Category_ID = ProductsT.Product_Category_ID) " & _
"ON ProductSubcategoryT.Subcategory_ID = ProductsT.Product_Subcategory_ID WHERE Product_ID is not null"

End Sub