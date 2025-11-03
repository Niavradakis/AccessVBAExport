SELECT [ProductsT].Product_ID, [ProductsT].Product_Description, [ProductsT].[Measure_Unit_For_Purchases/Sales_ID], [ProductsT].Sales_Category_ID, [ProductsT].Produced_From_Recipe, IIf(IsNull([Recipe_Record_ID]) And [Produced_From_Recipe]=True,"Λείπει Συνταγή") AS Recipes_Check, Count(RecipesΤ.Recipe_Record_ID) AS CountOfRecipe_Record_ID, [ProductsT].Active
FROM ProductsT LEFT JOIN RecipesΤ ON [ProductsT].Product_ID=RecipesΤ.Sale_Product_ID
GROUP BY [ProductsT].Product_ID, [ProductsT].Product_Description, [ProductsT].[Measure_Unit_For_Purchases/Sales_ID], [ProductsT].Sales_Category_ID, [ProductsT].Produced_From_Recipe, IIf(IsNull([Recipe_Record_ID]) And [Produced_From_Recipe]=True,"Λείπει Συνταγή"), [ProductsT].Active
HAVING ((Not ([ProductsT].Sales_Category_ID) Is Null));

