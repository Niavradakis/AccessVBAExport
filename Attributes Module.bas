Option Compare Database

Public Sub RowsourceWithAttributeID(Cbox As Control, TextToSearchArg As String, EntityTypeIDArg As Integer, EntityIDArg As Long, Optional AttributeFilterOptionsArg As Integer, Optional ReferenceEntityTypeIDArg As Long)
Debug.Print "Exec Priority - " & "Attributes Module - " & "RowsourceWithAttributeID " & Time()
'On Error GoTo ErrorHandler

Dim CboxVar As ComboBox

Set CboxVar = Cbox

Select Case EntityTypeIDArg

 Case 1 'products
      If AttributeFilterOptionsArg = 1 Then

                
         CboxVar.RowSource = "SELECT DISTINCT AttributesT.Attribute_Description, LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "FROM " & _
         "((SELECT DISTINCT LinkAttributeValueToEntitiesT.Attribute_ID, ProductsT.Product_Type_ID, AttributesT.Attribute_Description " & _
         "FROM " & _
         "AttributesT INNER JOIN (LinkAttributeValueToEntitiesT LEFT JOIN ProductsT ON LinkAttributeValueToEntitiesT.Entity_ID = ProductsT.Product_ID) " & _
         "ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "WHERE LinkAttributeValueToEntitiesT.Entity_Type_ID = 1 " & _
         "AND ProductsT.Product_ID =  " & EntityIDArg & _
         "AND AttributesT.Entities_May_Have_Multiple_Values = No)  AS CurrentEntityAttributesQ " & _
         "RIGHT JOIN ((ProductsT LEFT JOIN LinkAttributeValueToEntitiesT ON ProductsT.Product_ID = LinkAttributeValueToEntitiesT.Entity_ID) " & _
         "left join AttributesT on LinkAttributeValueToEntitiesT.Attribute_ID = AttributesT.Attribute_ID) " & _
         "ON CurrentEntityAttributesQ.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID) " & _
         "WHERE CurrentEntityAttributesQ.Attribute_ID Is Null AND LinkAttributeValueToEntitiesT.Entity_Type_ID = 1 " & _
         "AND ProductsT.Product_Type_ID = " & DLookup("ProductsT.Product_Type_ID", "ProductsT", "ProductsT.Product_ID = " & EntityIDArg) & _
         " AND ProductsT.Product_Category_ID = " & DLookup("ProductsT.Product_Category_ID", "ProductsT", "ProductsT.Product_ID = " & EntityIDArg) & _
         " AND ProductsT.Product_Subcategory_ID = " & DLookup("ProductsT.Product_Subcategory_ID", "ProductsT", "ProductsT.Product_ID = " & EntityIDArg) & _
         " AND AttributesT.Attribute_Description Like ""*" & TextToSearchArg & "*"" ORDER BY AttributesT.Attribute_Description;"
          
            CboxVar.ListWidth = CboxVar.ListWidth
          Exit Sub
        Else

        If AttributeFilterOptionsArg = 2 Then
        
        CboxVar.RowSource = "SELECT DISTINCT AttributesT.Attribute_Description, LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "FROM " & _
         "((SELECT DISTINCT LinkAttributeValueToEntitiesT.Attribute_ID, ProductsT.Product_Type_ID, AttributesT.Attribute_Description " & _
         "FROM " & _
         "AttributesT INNER JOIN (LinkAttributeValueToEntitiesT LEFT JOIN ProductsT ON LinkAttributeValueToEntitiesT.Entity_ID = ProductsT.Product_ID) " & _
         "ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "WHERE LinkAttributeValueToEntitiesT.Entity_Type_ID = 1 " & _
         "AND ProductsT.Product_ID =  " & EntityIDArg & _
         "AND AttributesT.Entities_May_Have_Multiple_Values = No)  AS CurrentEntityAttributesQ " & _
         "RIGHT JOIN ((ProductsT LEFT JOIN LinkAttributeValueToEntitiesT ON ProductsT.Product_ID = LinkAttributeValueToEntitiesT.Entity_ID) " & _
         "left join AttributesT on LinkAttributeValueToEntitiesT.Attribute_ID = AttributesT.Attribute_ID) " & _
         "ON CurrentEntityAttributesQ.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID) " & _
         "WHERE CurrentEntityAttributesQ.Attribute_ID Is Null AND LinkAttributeValueToEntitiesT.Entity_Type_ID = 1 " & _
         "AND ProductsT.Product_Type_ID = " & DLookup("ProductsT.Product_Type_ID", "ProductsT", "ProductsT.Product_ID = " & EntityIDArg) & _
         " AND ProductsT.Product_Category_ID = " & DLookup("ProductsT.Product_Category_ID", "ProductsT", "ProductsT.Product_ID = " & EntityIDArg) & _
         " AND AttributesT.Attribute_Description Like ""*" & TextToSearchArg & "*"" ORDER BY AttributesT.Attribute_Description;"
             
             
             CboxVar.ListWidth = CboxVar.ListWidth
           Exit Sub
         Else
         
        If AttributeFilterOptionsArg = 3 Then
       
       CboxVar.RowSource = "SELECT DISTINCT AttributesT.Attribute_Description, LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "FROM " & _
         "((SELECT DISTINCT LinkAttributeValueToEntitiesT.Attribute_ID, ProductsT.Product_Type_ID, AttributesT.Attribute_Description " & _
         "FROM " & _
         "AttributesT INNER JOIN (LinkAttributeValueToEntitiesT LEFT JOIN ProductsT ON LinkAttributeValueToEntitiesT.Entity_ID = ProductsT.Product_ID) " & _
         "ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "WHERE LinkAttributeValueToEntitiesT.Entity_Type_ID = 1 " & _
         "AND ProductsT.Product_ID =  " & EntityIDArg & _
         "AND AttributesT.Entities_May_Have_Multiple_Values = No)  AS CurrentEntityAttributesQ " & _
         "RIGHT JOIN ((ProductsT LEFT JOIN LinkAttributeValueToEntitiesT ON ProductsT.Product_ID = LinkAttributeValueToEntitiesT.Entity_ID) " & _
         "left join AttributesT on LinkAttributeValueToEntitiesT.Attribute_ID = AttributesT.Attribute_ID) " & _
         "ON CurrentEntityAttributesQ.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID) " & _
         "WHERE CurrentEntityAttributesQ.Attribute_ID Is Null AND LinkAttributeValueToEntitiesT.Entity_Type_ID = 1 " & _
         "AND ProductsT.Product_Type_ID = " & DLookup("ProductsT.Product_Type_ID", "ProductsT", "ProductsT.Product_ID = " & EntityIDArg) & _
         " AND AttributesT.Attribute_Description Like ""*" & TextToSearchArg & "*"" ORDER BY AttributesT.Attribute_Description;"
       
              CboxVar.ListWidth = CboxVar.ListWidth
            Exit Sub
           Else
           
        If AttributeFilterOptionsArg = 4 Then
       
        CboxVar.RowSource = "SELECT DISTINCT AttributesT.Attribute_Description, LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "FROM " & _
         "((SELECT DISTINCT LinkAttributeValueToEntitiesT.Attribute_ID, ProductsT.Product_Type_ID, AttributesT.Attribute_Description " & _
         "FROM " & _
         "AttributesT INNER JOIN (LinkAttributeValueToEntitiesT LEFT JOIN ProductsT ON LinkAttributeValueToEntitiesT.Entity_ID = ProductsT.Product_ID) " & _
         "ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "WHERE LinkAttributeValueToEntitiesT.Entity_Type_ID = 1 " & _
         "AND ProductsT.Product_ID =  " & EntityIDArg & _
         "AND AttributesT.Entities_May_Have_Multiple_Values = No)  AS CurrentEntityAttributesQ " & _
         "RIGHT JOIN ((ProductsT LEFT JOIN LinkAttributeValueToEntitiesT ON ProductsT.Product_ID = LinkAttributeValueToEntitiesT.Entity_ID) " & _
         "left join AttributesT on LinkAttributeValueToEntitiesT.Attribute_ID = AttributesT.Attribute_ID) " & _
         "ON CurrentEntityAttributesQ.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID) " & _
         "WHERE CurrentEntityAttributesQ.Attribute_ID Is Null AND LinkAttributeValueToEntitiesT.Entity_Type_ID = 1 " & _
         "and AttributesT.Attribute_Description Like ""*" & TextToSearchArg & "*"" ORDER BY AttributesT.Attribute_Description;"
       
              CboxVar.ListWidth = CboxVar.ListWidth
            Exit Sub
           Else
        If AttributeFilterOptionsArg = 5 Then
             
       CboxVar.RowSource = "SELECT AttributesT.Attribute_Description, AttributesT.Attribute_ID " & _
         "FROM (SELECT DISTINCT LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "FROM " & _
         "AttributesT INNER JOIN (LinkAttributeValueToEntitiesT LEFT JOIN ProductsT ON LinkAttributeValueToEntitiesT.Entity_ID = ProductsT.Product_ID) " & _
         "ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "WHERE LinkAttributeValueToEntitiesT.Entity_Type_ID = 1 " & _
         "AND ProductsT.Product_ID =  " & EntityIDArg & _
         "AND AttributesT.Entities_May_Have_Multiple_Values = No)  AS CurrentEntityAttributesQ " & _
         "RIGHT JOIN AttributesT ON CurrentEntityAttributesQ.Attribute_ID = AttributesT.Attribute_ID " & _
         "WHERE CurrentEntityAttributesQ.Attribute_ID Is Null AND AttributesT.Attribute_Description Like ""*" & TextToSearchArg & "*"" ORDER BY AttributesT.Attribute_Description;"
                     
                          
              CboxVar.ListWidth = CboxVar.ListWidth
            Exit Sub
           Else
           
           End If
           End If
           End If
          End If
         End If

 Case 2 'transactors
      If AttributeFilterOptionsArg = 1 Then

       CboxVar.RowSource = "SELECT DISTINCT AttributesT.Attribute_Description, LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "FROM " & _
         "((SELECT DISTINCT LinkAttributeValueToEntitiesT.Attribute_ID, TransactorsT.Transactor_Type_ID, AttributesT.Attribute_Description " & _
         "FROM " & _
         "AttributesT INNER JOIN (LinkAttributeValueToEntitiesT LEFT JOIN TransactorsT ON LinkAttributeValueToEntitiesT.Entity_ID = TransactorsT.Transactor_ID) " & _
         "ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "WHERE LinkAttributeValueToEntitiesT.Entity_Type_ID = 2 " & _
         "AND TransactorsT.Transactor_ID =  " & EntityIDArg & _
         "AND AttributesT.Entities_May_Have_Multiple_Values = No)  AS CurrentEntityAttributesQ " & _
         "RIGHT JOIN ((TransactorsT LEFT JOIN LinkAttributeValueToEntitiesT ON TransactorsT.Transactor_ID = LinkAttributeValueToEntitiesT.Entity_ID) " & _
         "left join AttributesT on LinkAttributeValueToEntitiesT.Attribute_ID = AttributesT.Attribute_ID) " & _
         "ON CurrentEntityAttributesQ.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID) " & _
         "WHERE CurrentEntityAttributesQ.Attribute_ID Is Null AND LinkAttributeValueToEntitiesT.Entity_Type_ID = 2 " & _
         "AND TransactorsT.Transactor_Type_ID = " & DLookup("TransactorsT.Transactor_Type_ID", "TransactorsT", "TransactorsT.Transactor_ID = " & EntityIDArg) & _
         "and AttributesT.Attribute_Description Like ""*" & TextToSearchArg & "*"" ORDER BY AttributesT.Attribute_Description;"
            
            CboxVar.ListWidth = CboxVar.ListWidth
          Exit Sub
        Else

        If AttributeFilterOptionsArg = 2 Then
               
         CboxVar.RowSource = "SELECT DISTINCT AttributesT.Attribute_Description, LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "FROM " & _
         "((SELECT DISTINCT LinkAttributeValueToEntitiesT.Attribute_ID, TransactorsT.Transactor_Type_ID, AttributesT.Attribute_Description " & _
         "FROM " & _
         "AttributesT INNER JOIN (LinkAttributeValueToEntitiesT LEFT JOIN TransactorsT ON LinkAttributeValueToEntitiesT.Entity_ID = TransactorsT.Transactor_ID) " & _
         "ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "WHERE LinkAttributeValueToEntitiesT.Entity_Type_ID = 2 " & _
         "AND TransactorsT.Transactor_ID =  " & EntityIDArg & _
         "AND AttributesT.Entities_May_Have_Multiple_Values = No)  AS CurrentEntityAttributesQ " & _
         "RIGHT JOIN ((TransactorsT LEFT JOIN LinkAttributeValueToEntitiesT ON TransactorsT.Transactor_ID = LinkAttributeValueToEntitiesT.Entity_ID) " & _
         "left join AttributesT on LinkAttributeValueToEntitiesT.Attribute_ID = AttributesT.Attribute_ID) " & _
         "ON CurrentEntityAttributesQ.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID) " & _
         "WHERE CurrentEntityAttributesQ.Attribute_ID Is Null AND LinkAttributeValueToEntitiesT.Entity_Type_ID = 2 " & _
         "AND AttributesT.Attribute_Description Like ""*" & TextToSearchArg & "*"" ORDER BY AttributesT.Attribute_Description;"
       
             CboxVar.ListWidth = CboxVar.ListWidth
           Exit Sub
         Else
         
        If AttributeFilterOptionsArg = 3 Then
       
             CboxVar.RowSource = "SELECT AttributesT.Attribute_Description, AttributesT.Attribute_ID " & _
             "FROM (SELECT DISTINCT LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "FROM " & _
         "AttributesT INNER JOIN (LinkAttributeValueToEntitiesT LEFT JOIN TransactorsT ON LinkAttributeValueToEntitiesT.Entity_ID = TransactorsT.Transactor_ID) " & _
         "ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "WHERE LinkAttributeValueToEntitiesT.Entity_Type_ID=2 " & _
         "AND TransactorsT.Transactor_ID = " & EntityIDArg & _
         "AND AttributesT.Entities_May_Have_Multiple_Values=No) AS CurrentEntityAttributesQ " & _
         "RIGHT JOIN AttributesT ON CurrentEntityAttributesQ.Attribute_ID = AttributesT.Attribute_ID " & _
             "WHERE CurrentEntityAttributesQ.Attribute_ID Is Null and AttributesT.Attribute_Description Like ""*" & TextToSearchArg & "*"" order by AttributesT.Attribute_Description ;"
       
             CboxVar.ListWidth = CboxVar.ListWidth
            Exit Sub
           End If
          End If
         End If
         
 Case 3 'Issued Document
      If AttributeFilterOptionsArg = 1 Then

       CboxVar.RowSource = "SELECT DISTINCT AttributesT.Attribute_Description, LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "FROM " & _
         "((SELECT DISTINCT LinkAttributeValueToEntitiesT.Attribute_ID, IssuedDocumentT.Issuable_Document_ID, AttributesT.Attribute_Description " & _
         "FROM " & _
         "AttributesT INNER JOIN (LinkAttributeValueToEntitiesT INNER JOIN IssuedDocumentT ON LinkAttributeValueToEntitiesT.Entity_ID = IssuedDocumentT.Issued_Document_ID ) " & _
         "ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "WHERE LinkAttributeValueToEntitiesT.Entity_Type_ID = 3 AND LinkAttributeValueToEntitiesT.Entity_ID = " & EntityIDArg & _
         " AND AttributesT.Entities_May_Have_Multiple_Values = No)  AS CurrentEntityAttributesQ " & _
         " RIGHT JOIN ((IssuedDocumentT LEFT JOIN LinkAttributeValueToEntitiesT ON IssuedDocumentT.Issued_Document_ID = LinkAttributeValueToEntitiesT.Entity_ID) " & _
         "left join AttributesT on LinkAttributeValueToEntitiesT.Attribute_ID = AttributesT.Attribute_ID) " & _
         "ON CurrentEntityAttributesQ.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID) " & _
         " WHERE CurrentEntityAttributesQ.Attribute_ID Is Null AND LinkAttributeValueToEntitiesT.Entity_Type_ID = 3 " & _
         "AND IssuedDocumentT.Issuable_Document_ID = " & DLookup("IssuedDocumentT.Issuable_Document_ID", "IssuedDocumentT", "IssuedDocumentT.Issued_Document_ID = " & EntityIDArg) & _
         " AND AttributesT.Attribute_Description Like ""*" & TextToSearchArg & "*"" ORDER BY AttributesT.Attribute_Description;"
         
  
            CboxVar.ListWidth = CboxVar.ListWidth
          Exit Sub
        Else

        If AttributeFilterOptionsArg = 2 Then
        
        CboxVar.RowSource = "SELECT DISTINCT AttributesT.Attribute_Description, LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "FROM " & _
         "((SELECT DISTINCT LinkAttributeValueToEntitiesT.Attribute_ID, IssuedDocumentT.Issuable_Document_ID, AttributesT.Attribute_Description " & _
         "FROM " & _
         "AttributesT INNER JOIN (LinkAttributeValueToEntitiesT INNER JOIN IssuedDocumentT ON LinkAttributeValueToEntitiesT.Entity_ID = IssuedDocumentT.Issued_Document_ID ) " & _
         "ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "WHERE LinkAttributeValueToEntitiesT.Entity_Type_ID = 3 AND LinkAttributeValueToEntitiesT.Entity_ID = " & EntityIDArg & _
         " AND AttributesT.Entities_May_Have_Multiple_Values = No)  AS CurrentEntityAttributesQ " & _
         " RIGHT JOIN ((IssuedDocumentT LEFT JOIN LinkAttributeValueToEntitiesT ON IssuedDocumentT.Issued_Document_ID = LinkAttributeValueToEntitiesT.Entity_ID) " & _
         "left join AttributesT on LinkAttributeValueToEntitiesT.Attribute_ID = AttributesT.Attribute_ID) " & _
         "ON CurrentEntityAttributesQ.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID) " & _
         "WHERE CurrentEntityAttributesQ.Attribute_ID Is Null AND LinkAttributeValueToEntitiesT.Entity_Type_ID=3 " & _
         " and AttributesT.Attribute_Description Like ""*" & TextToSearchArg & "*"" ORDER BY AttributesT.Attribute_Description;"
             
             CboxVar.ListWidth = CboxVar.ListWidth
           Exit Sub
         Else
         
        If AttributeFilterOptionsArg = 3 Then
       
             CboxVar.RowSource = "SELECT AttributesT.Attribute_Description, AttributesT.Attribute_ID " & _
             "FROM (SELECT DISTINCT LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "FROM " & _
         "AttributesT INNER JOIN (LinkAttributeValueToEntitiesT INNER JOIN IssuedDocumentT ON LinkAttributeValueToEntitiesT.Entity_ID = IssuedDocumentT.Issued_Document_ID) " & _
         "ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "WHERE LinkAttributeValueToEntitiesT.Entity_Type_ID = 3 " & _
         "AND IssuedDocumentT.Issued_Document_ID = " & EntityIDArg & _
         "AND AttributesT.Entities_May_Have_Multiple_Values=No) AS CurrentEntityAttributesQ " & _
         "RIGHT JOIN AttributesT ON CurrentEntityAttributesQ.Attribute_ID = AttributesT.Attribute_ID " & _
             "WHERE CurrentEntityAttributesQ.Attribute_ID Is Null and AttributesT.Attribute_Description Like ""*" & TextToSearchArg & "*"" order by AttributesT.Attribute_Description ;"
       
              CboxVar.ListWidth = CboxVar.ListWidth
            Exit Sub
           End If
          End If
         End If
         
 Case 4 'Document Financial Details
      
  
  Dim TransactorIDvar As Integer
  Dim TransactorTypeIDvar As Integer
  TransactorIDvar = DLookup("IssuedDocumentFinancialDetailsT.Transactor_ID", "IssuedDocumentFinancialDetailsT", "IssuedDocumentFinancialDetailsT.Issued_Document_Financial_Details_ID = " & EntityIDArg)
  TransactorTypeIDvar = DLookup("TransactorsT.Transactor_Type_ID", "TransactorsT", "TransactorsT.Transactor_ID = " & TransactorIDvar)

     If AttributeFilterOptionsArg = 1 Then
       CboxVar.RowSource = "SELECT DISTINCT AttributesT.Attribute_Description, LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "FROM " & _
         "((SELECT DISTINCT LinkAttributeValueToEntitiesT.Attribute_ID, IssuedDocumentFinancialDetailsT.Transactor_ID, AttributesT.Attribute_Description " & _
         "FROM " & _
         "AttributesT INNER JOIN (LinkAttributeValueToEntitiesT INNER JOIN IssuedDocumentFinancialDetailsT ON LinkAttributeValueToEntitiesT.Entity_ID = IssuedDocumentFinancialDetailsT.Issued_Document_Financial_Details_ID ) " & _
         "ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "WHERE LinkAttributeValueToEntitiesT.Entity_Type_ID = 4 " & _
         "AND IssuedDocumentFinancialDetailsT.Issued_Document_Financial_Details_ID = " & EntityIDArg & _
         "AND AttributesT.Entities_May_Have_Multiple_Values = No)  AS CurrentEntityAttributesQ " & _
         "RIGHT JOIN (((IssuedDocumentFinancialDetailsT LEFT JOIN TransactorsT ON IssuedDocumentFinancialDetailsT.Transactor_ID = TransactorsT.Transactor_ID) LEFT JOIN LinkAttributeValueToEntitiesT ON IssuedDocumentFinancialDetailsT.Issued_Document_Financial_Details_ID = LinkAttributeValueToEntitiesT.Entity_ID) " & _
         "left join AttributesT on LinkAttributeValueToEntitiesT.Attribute_ID = AttributesT.Attribute_ID) " & _
         "ON CurrentEntityAttributesQ.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID) " & _
         "WHERE CurrentEntityAttributesQ.Attribute_ID Is Null AND LinkAttributeValueToEntitiesT.Entity_Type_ID=4 " & _
         "AND IssuedDocumentFinancialDetailsT.Transactor_ID = " & DLookup("IssuedDocumentFinancialDetailsT.Transactor_ID", "IssuedDocumentFinancialDetailsT", "IssuedDocumentFinancialDetailsT.Issued_Document_Financial_Details_ID = " & EntityIDArg) & _
         "and AttributesT.Attribute_Description Like ""*" & TextToSearchArg & "*"" ORDER BY AttributesT.Attribute_Description;"
          
        
            CboxVar.ListWidth = CboxVar.ListWidth
          Exit Sub
        Else
        
      If AttributeFilterOptionsArg = 2 Then

       CboxVar.RowSource = "SELECT DISTINCT AttributesT.Attribute_Description, LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "FROM " & _
         "((SELECT DISTINCT LinkAttributeValueToEntitiesT.Attribute_ID, IssuedDocumentFinancialDetailsT.Transactor_ID, AttributesT.Attribute_Description " & _
         "FROM " & _
         "AttributesT INNER JOIN (LinkAttributeValueToEntitiesT INNER JOIN IssuedDocumentFinancialDetailsT ON LinkAttributeValueToEntitiesT.Entity_ID = IssuedDocumentFinancialDetailsT.Issued_Document_Financial_Details_ID ) " & _
         "ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "WHERE LinkAttributeValueToEntitiesT.Entity_Type_ID = 4 " & _
         "AND IssuedDocumentFinancialDetailsT.Issued_Document_Financial_Details_ID = " & EntityIDArg & _
         "AND AttributesT.Entities_May_Have_Multiple_Values = No)  AS CurrentEntityAttributesQ " & _
         "RIGHT JOIN (((IssuedDocumentFinancialDetailsT LEFT JOIN TransactorsT ON IssuedDocumentFinancialDetailsT.Transactor_ID = TransactorsT.Transactor_ID) LEFT JOIN LinkAttributeValueToEntitiesT ON IssuedDocumentFinancialDetailsT.Issued_Document_Financial_Details_ID = LinkAttributeValueToEntitiesT.Entity_ID) " & _
         "left join AttributesT on LinkAttributeValueToEntitiesT.Attribute_ID = AttributesT.Attribute_ID) " & _
         "ON CurrentEntityAttributesQ.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID) " & _
         "WHERE CurrentEntityAttributesQ.Attribute_ID Is Null AND LinkAttributeValueToEntitiesT.Entity_Type_ID=4 " & _
         "AND TransactorsT.Transactor_Type_ID = " & TransactorTypeIDvar & _
         "and AttributesT.Attribute_Description Like ""*" & TextToSearchArg & "*"" ORDER BY AttributesT.Attribute_Description;"
             
             
             CboxVar.ListWidth = CboxVar.ListWidth
          Exit Sub
        Else
        
       If AttributeFilterOptionsArg = 3 Then
       
       CboxVar.RowSource = "SELECT DISTINCT AttributesT.Attribute_Description, LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "FROM " & _
         "((SELECT DISTINCT LinkAttributeValueToEntitiesT.Attribute_ID, IssuedDocumentFinancialDetailsT.Transactor_ID, AttributesT.Attribute_Description " & _
         "FROM " & _
         "AttributesT INNER JOIN (LinkAttributeValueToEntitiesT INNER JOIN IssuedDocumentFinancialDetailsT ON LinkAttributeValueToEntitiesT.Entity_ID = IssuedDocumentFinancialDetailsT.Issued_Document_Financial_Details_ID ) " & _
         "ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "WHERE LinkAttributeValueToEntitiesT.Entity_Type_ID = 4 " & _
         "AND IssuedDocumentFinancialDetailsT.Issued_Document_Financial_Details_ID = " & EntityIDArg & _
         "AND AttributesT.Entities_May_Have_Multiple_Values = No)  AS CurrentEntityAttributesQ " & _
         "RIGHT JOIN (((IssuedDocumentFinancialDetailsT LEFT JOIN TransactorsT ON IssuedDocumentFinancialDetailsT.Transactor_ID = TransactorsT.Transactor_ID) LEFT JOIN LinkAttributeValueToEntitiesT ON IssuedDocumentFinancialDetailsT.Issued_Document_Financial_Details_ID = LinkAttributeValueToEntitiesT.Entity_ID) " & _
         "left join AttributesT on LinkAttributeValueToEntitiesT.Attribute_ID = AttributesT.Attribute_ID) " & _
         "ON CurrentEntityAttributesQ.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID) " & _
         "WHERE CurrentEntityAttributesQ.Attribute_ID Is Null AND LinkAttributeValueToEntitiesT.Entity_Type_ID=4 " & _
         "and AttributesT.Attribute_Description Like ""*" & TextToSearchArg & "*"" ORDER BY AttributesT.Attribute_Description;"
        
             CboxVar.ListWidth = CboxVar.ListWidth
           Exit Sub
         Else
         
        If AttributeFilterOptionsArg = 4 Then
       
             CboxVar.RowSource = "SELECT AttributesT.Attribute_Description, AttributesT.Attribute_ID " & _
             "FROM (SELECT DISTINCT LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "FROM " & _
         "AttributesT INNER JOIN (LinkAttributeValueToEntitiesT LEFT JOIN IssuedDocumentFinancialDetailsT ON LinkAttributeValueToEntitiesT.Entity_ID = IssuedDocumentFinancialDetailsT.Issued_Document_Financial_Details_ID) " & _
         "ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "WHERE LinkAttributeValueToEntitiesT.Entity_Type_ID = 4 " & _
         "AND IssuedDocumentFinancialDetailsT.Issued_Document_Financial_Details_ID = " & EntityIDArg & _
         " AND AttributesT.Entities_May_Have_Multiple_Values=No) AS CurrentEntityAttributesQ " & _
         "RIGHT JOIN AttributesT ON CurrentEntityAttributesQ.Attribute_ID = AttributesT.Attribute_ID " & _
             "WHERE ((CurrentEntityAttributesQ.Attribute_ID) Is Null)and AttributesT.Attribute_Description Like ""*" & TextToSearchArg & "*"" order by AttributesT.Attribute_Description ;"
       
              CboxVar.ListWidth = CboxVar.ListWidth
            Exit Sub
           End If
          End If
         End If
        End If
Case 5 'Document Product Details
    Dim ProductIDvar As Integer
    Dim ProductTypeIDvar As Integer
    Dim ProductCategoryIDvar As Integer
    Dim ProductSubcategoryIDvar As Integer
    
    ProductIDvar = DLookup("IssuedDocumentProductDetailsT.Product_ID", "IssuedDocumentProductDetailsT", "IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID = " & EntityIDArg)
    ProductTypeIDvar = DLookup("ProductsT.Product_Type_ID", "ProductsT", "ProductsT.Product_ID = " & ProductIDvar)
    ProductCategoryIDvar = DLookup("ProductsT.Product_Category_ID", "ProductsT", "ProductsT.Product_ID = " & ProductIDvar)
    ProductSubcategoryIDvar = DLookup("ProductsT.Product_Subcategory_ID", "ProductsT", "ProductsT.Product_ID = " & ProductIDvar)
    
      If AttributeFilterOptionsArg = 1 Then
        
       CboxVar.RowSource = "SELECT DISTINCT AttributesT.Attribute_Description, LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "FROM " & _
         "((SELECT DISTINCT LinkAttributeValueToEntitiesT.Attribute_ID, IssuedDocumentProductDetailsT.Product_ID, AttributesT.Attribute_Description " & _
         "FROM " & _
         "AttributesT INNER JOIN (LinkAttributeValueToEntitiesT INNER JOIN IssuedDocumentProductDetailsT ON LinkAttributeValueToEntitiesT.Entity_ID = IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID) " & _
         "ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "WHERE LinkAttributeValueToEntitiesT.Entity_Type_ID = 5 " & _
         "AND IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID =  " & EntityIDArg & _
         " AND AttributesT.Entities_May_Have_Multiple_Values = No)  AS CurrentEntityAttributesQ " & _
         "RIGHT JOIN (((IssuedDocumentProductDetailsT LEFT JOIN ProductsT ON IssuedDocumentProductDetailsT.Product_ID = ProductsT.Product_ID) LEFT JOIN LinkAttributeValueToEntitiesT ON IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID = LinkAttributeValueToEntitiesT.Entity_ID) " & _
         "left join AttributesT on LinkAttributeValueToEntitiesT.Attribute_ID = AttributesT.Attribute_ID) " & _
         "ON CurrentEntityAttributesQ.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID) " & _
         "WHERE CurrentEntityAttributesQ.Attribute_ID Is Null AND LinkAttributeValueToEntitiesT.Entity_Type_ID = 5 " & _
         "AND ProductsT.Product_Type_ID = " & ProductTypeIDvar & " AND ProductsT.Product_Category_ID = " & ProductCategoryIDvar & " AND ProductsT.Product_Subcategory_ID = " & ProductSubcategoryIDvar & _
         " AND AttributesT.Attribute_Description Like ""*" & TextToSearchArg & "*"" ORDER BY AttributesT.Attribute_Description;"
                   
            CboxVar.ListWidth = CboxVar.ListWidth
          Exit Sub
        Else

        If AttributeFilterOptionsArg = 2 Then
        
         CboxVar.RowSource = "SELECT DISTINCT AttributesT.Attribute_Description, LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "FROM " & _
         "((SELECT DISTINCT LinkAttributeValueToEntitiesT.Attribute_ID, IssuedDocumentProductDetailsT.Product_ID, AttributesT.Attribute_Description " & _
         "FROM " & _
         "AttributesT INNER JOIN (LinkAttributeValueToEntitiesT INNER JOIN IssuedDocumentProductDetailsT ON LinkAttributeValueToEntitiesT.Entity_ID = IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID) " & _
         "ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "WHERE LinkAttributeValueToEntitiesT.Entity_Type_ID = 5 " & _
         "AND IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID =  " & EntityIDArg & _
         " AND AttributesT.Entities_May_Have_Multiple_Values = No)  AS CurrentEntityAttributesQ " & _
         "RIGHT JOIN (((IssuedDocumentProductDetailsT INNER JOIN ProductsT ON IssuedDocumentProductDetailsT.Product_ID = ProductsT.Product_ID) LEFT JOIN LinkAttributeValueToEntitiesT ON IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID = LinkAttributeValueToEntitiesT.Entity_ID) " & _
         "left join AttributesT on LinkAttributeValueToEntitiesT.Attribute_ID = AttributesT.Attribute_ID) " & _
         "ON CurrentEntityAttributesQ.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID) " & _
         "WHERE CurrentEntityAttributesQ.Attribute_ID Is Null AND LinkAttributeValueToEntitiesT.Entity_Type_ID = 5 " & _
         "AND ProductsT.Product_Type_ID = " & ProductTypeIDvar & " AND ProductsT.Product_Category_ID = " & ProductCategoryIDvar & _
         " AND AttributesT.Attribute_Description Like ""*" & TextToSearchArg & "*"" ORDER BY AttributesT.Attribute_Description;"
             
             CboxVar.ListWidth = CboxVar.ListWidth
           Exit Sub
         Else
         
         If AttributeFilterOptionsArg = 3 Then
        
         CboxVar.RowSource = "SELECT DISTINCT AttributesT.Attribute_Description, LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "FROM " & _
         "((SELECT DISTINCT LinkAttributeValueToEntitiesT.Attribute_ID, IssuedDocumentProductDetailsT.Product_ID, AttributesT.Attribute_Description " & _
         "FROM " & _
         "AttributesT INNER JOIN (LinkAttributeValueToEntitiesT INNER JOIN IssuedDocumentProductDetailsT ON LinkAttributeValueToEntitiesT.Entity_ID = IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID) " & _
         "ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "WHERE LinkAttributeValueToEntitiesT.Entity_Type_ID = 5 " & _
         "AND IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID =  " & EntityIDArg & _
         " AND AttributesT.Entities_May_Have_Multiple_Values = No)  AS CurrentEntityAttributesQ " & _
         "RIGHT JOIN (((IssuedDocumentProductDetailsT INNER JOIN ProductsT ON IssuedDocumentProductDetailsT.Product_ID = ProductsT.Product_ID) LEFT JOIN LinkAttributeValueToEntitiesT ON IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID = LinkAttributeValueToEntitiesT.Entity_ID) " & _
         "left join AttributesT on LinkAttributeValueToEntitiesT.Attribute_ID = AttributesT.Attribute_ID) " & _
         "ON CurrentEntityAttributesQ.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID) " & _
         "WHERE CurrentEntityAttributesQ.Attribute_ID Is Null AND LinkAttributeValueToEntitiesT.Entity_Type_ID = 5 " & _
         "AND ProductsT.Product_Type_ID = " & ProductTypeIDvar & _
         " AND AttributesT.Attribute_Description Like ""*" & TextToSearchArg & "*"" ORDER BY AttributesT.Attribute_Description;"
             
             CboxVar.ListWidth = CboxVar.ListWidth
           Exit Sub
         Else
         
        If AttributeFilterOptionsArg = 4 Then
        
         CboxVar.RowSource = "SELECT DISTINCT AttributesT.Attribute_Description, LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "FROM " & _
         "((SELECT DISTINCT LinkAttributeValueToEntitiesT.Attribute_ID, IssuedDocumentProductDetailsT.Product_ID, AttributesT.Attribute_Description " & _
         "FROM " & _
         "AttributesT INNER JOIN (LinkAttributeValueToEntitiesT INNER JOIN IssuedDocumentProductDetailsT ON LinkAttributeValueToEntitiesT.Entity_ID = IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID) " & _
         "ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "WHERE LinkAttributeValueToEntitiesT.Entity_Type_ID = 5 " & _
         "AND IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID =  " & EntityIDArg & _
         " AND AttributesT.Entities_May_Have_Multiple_Values = No)  AS CurrentEntityAttributesQ " & _
         "RIGHT JOIN (((IssuedDocumentProductDetailsT LEFT JOIN ProductsT ON IssuedDocumentProductDetailsT.Product_ID = ProductsT.Product_ID)LEFT JOIN LinkAttributeValueToEntitiesT ON IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID = LinkAttributeValueToEntitiesT.Entity_ID) " & _
         "left join AttributesT on LinkAttributeValueToEntitiesT.Attribute_ID = AttributesT.Attribute_ID) " & _
         "ON CurrentEntityAttributesQ.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID) " & _
         "WHERE CurrentEntityAttributesQ.Attribute_ID Is Null AND LinkAttributeValueToEntitiesT.Entity_Type_ID = 5 " & _
         "and AttributesT.Attribute_Description Like ""*" & TextToSearchArg & "*"" ORDER BY AttributesT.Attribute_Description;"
               
             CboxVar.ListWidth = CboxVar.ListWidth
           Exit Sub
         Else
         
        If AttributeFilterOptionsArg = 5 Then
       
             CboxVar.RowSource = "SELECT AttributesT.Attribute_Description, AttributesT.Attribute_ID " & _
             "FROM (SELECT DISTINCT LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "FROM " & _
         "AttributesT INNER JOIN (LinkAttributeValueToEntitiesT INNER JOIN IssuedDocumentProductDetailsT ON LinkAttributeValueToEntitiesT.Entity_ID = IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID) " & _
         "ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "WHERE LinkAttributeValueToEntitiesT.Entity_Type_ID = 5 " & _
         "AND IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID = " & EntityIDArg & _
         "AND AttributesT.Entities_May_Have_Multiple_Values=No) AS CurrentEntityAttributesQ " & _
         "RIGHT JOIN AttributesT ON CurrentEntityAttributesQ.Attribute_ID = AttributesT.Attribute_ID " & _
             "WHERE CurrentEntityAttributesQ.Attribute_ID Is Null and AttributesT.Attribute_Description Like ""*" & TextToSearchArg & "*"" order by AttributesT.Attribute_Description ;"
       
              CboxVar.ListWidth = CboxVar.ListWidth
            Exit Sub
          Else
         
        If AttributeFilterOptionsArg = 5 Then
       
             CboxVar.RowSource = "SELECT AttributesT.Attribute_Description, AttributesT.Attribute_ID " & _
             "FROM (SELECT DISTINCT LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "FROM " & _
         "AttributesT INNER JOIN (LinkAttributeValueToEntitiesT INNER JOIN IssuedDocumentProductDetailsT ON LinkAttributeValueToEntitiesT.Entity_ID = IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID) " & _
         "ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "WHERE LinkAttributeValueToEntitiesT.Entity_Type_ID = 5 " & _
         "AND IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID = " & EntityIDArg & _
         "AND AttributesT.Entities_May_Have_Multiple_Values=No) AS CurrentEntityAttributesQ " & _
         "RIGHT JOIN AttributesT ON CurrentEntityAttributesQ.Attribute_ID = AttributesT.Attribute_ID " & _
             "WHERE CurrentEntityAttributesQ.Attribute_ID Is Null and AttributesT.Attribute_Description Like ""*" & TextToSearchArg & "*"" order by AttributesT.Attribute_Description ;"
       
              CboxVar.ListWidth = CboxVar.ListWidth
            Exit Sub
            
           End If
          End If
         End If
       End If
      End If
      End If
 Case 6 'Transactions
      If AttributeFilterOptionsArg = 1 Then

       CboxVar.RowSource = "SELECT DISTINCT AttributesT.Attribute_Description, LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "FROM " & _
         "((SELECT DISTINCT LinkAttributeValueToEntitiesT.Attribute_ID, TransactionsT.Transaction_Type_ID, AttributesT.Attribute_Description " & _
         "FROM " & _
         "AttributesT INNER JOIN (LinkAttributeValueToEntitiesT INNER JOIN TransactionsT ON LinkAttributeValueToEntitiesT.Entity_ID = TransactionsT.Transaction_ID) " & _
         "ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "WHERE LinkAttributeValueToEntitiesT.Entity_Type_ID = 6 " & _
         "AND TransactionsT.Transaction_ID =  " & EntityIDArg & _
         " AND AttributesT.Entities_May_Have_Multiple_Values = No)  AS CurrentEntityAttributesQ " & _
         "RIGHT JOIN ((TransactionsT LEFT JOIN LinkAttributeValueToEntitiesT ON TransactionsT.Transaction_ID = LinkAttributeValueToEntitiesT.Entity_ID) " & _
         "left join AttributesT on LinkAttributeValueToEntitiesT.Attribute_ID = AttributesT.Attribute_ID) " & _
         "ON CurrentEntityAttributesQ.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID) " & _
         "WHERE CurrentEntityAttributesQ.Attribute_ID Is Null AND LinkAttributeValueToEntitiesT.Entity_Type_ID = 6 " & _
         "AND TransactionsT.Transaction_Type_ID = " & DLookup("TransactionsT.Transaction_Type_ID", "TransactionsT", "TransactionsT.Transaction_ID = " & EntityIDArg) & _
         " AND AttributesT.Attribute_Description Like ""*" & TextToSearchArg & "*"" ORDER BY AttributesT.Attribute_Description;"
          
            CboxVar.ListWidth = CboxVar.ListWidth
          Exit Sub
        Else

        If AttributeFilterOptionsArg = 2 Then
        
         CboxVar.RowSource = "SELECT DISTINCT AttributesT.Attribute_Description, LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "FROM " & _
         "((SELECT DISTINCT LinkAttributeValueToEntitiesT.Attribute_ID, TransactionsT.Transaction_Type_ID, AttributesT.Attribute_Description " & _
         "FROM " & _
         "AttributesT INNER JOIN (LinkAttributeValueToEntitiesT INNER JOIN TransactionsT ON LinkAttributeValueToEntitiesT.Entity_ID = TransactionsT.Transaction_ID) " & _
         "ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "WHERE LinkAttributeValueToEntitiesT.Entity_Type_ID = 6 " & _
         "AND TransactionsT.Transaction_ID =  " & EntityIDArg & _
         " AND AttributesT.Entities_May_Have_Multiple_Values = No)  AS CurrentEntityAttributesQ " & _
         "RIGHT JOIN ((TransactionsT LEFT JOIN LinkAttributeValueToEntitiesT ON TransactionsT.Transaction_ID = LinkAttributeValueToEntitiesT.Entity_ID) " & _
         "left join AttributesT on LinkAttributeValueToEntitiesT.Attribute_ID = AttributesT.Attribute_ID) " & _
         "ON CurrentEntityAttributesQ.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID) " & _
         "WHERE CurrentEntityAttributesQ.Attribute_ID Is Null AND LinkAttributeValueToEntitiesT.Entity_Type_ID = 6 " & _
         "and AttributesT.Attribute_Description Like ""*" & TextToSearchArg & "*"" ORDER BY AttributesT.Attribute_Description;"
             
             CboxVar.ListWidth = CboxVar.ListWidth
           Exit Sub
         Else
         
        If AttributeFilterOptionsArg = 3 Then
       
             CboxVar.RowSource = "SELECT AttributesT.Attribute_Description, AttributesT.Attribute_ID " & _
             "FROM (SELECT DISTINCT LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "FROM " & _
         "AttributesT INNER JOIN (LinkAttributeValueToEntitiesT INNER JOIN TransactionsT ON LinkAttributeValueToEntitiesT.Entity_ID = TransactionsT.Transaction_ID) " & _
         "ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "WHERE LinkAttributeValueToEntitiesT.Entity_Type_ID = 6 " & _
         "AND TransactionsT.Transaction_ID = " & EntityIDArg & _
         " AND AttributesT.Entities_May_Have_Multiple_Values = No) AS CurrentEntityAttributesQ " & _
         "RIGHT JOIN AttributesT ON CurrentEntityAttributesQ.Attribute_ID = AttributesT.Attribute_ID " & _
         "WHERE CurrentEntityAttributesQ.Attribute_ID Is Null and AttributesT.Attribute_Description Like ""*" & TextToSearchArg & "*"" order by AttributesT.Attribute_Description ;"
       
            
             CboxVar.ListWidth = CboxVar.ListWidth
            Exit Sub
           End If
          End If
         End If
         
 Case 8 'actions
      If AttributeFilterOptionsArg = 0 Then
      
       'CboxVar.RowSource = "SELECT  AttributesT.Attribute_Description, [LInkReferenceID(EntityTypeID)ToAttributeIDsT].Attribute_ID " & _
       "FROM AttributesT RIGHT JOIN ([LInkReferenceID(EntityTypeID)ToAttributeIDsT] LEFT JOIN (SELECT distinct LinkAttributeValueToEntitiesT.Link_Attribute_Value_To_Entity_ID, LinkAttributeValueToEntitiesT.Entity_Type_ID, " & _
       "LinkAttributeValueToEntitiesT.Attribute_ID " & _
       "FROM ActionsT INNER JOIN LinkAttributeValueToEntitiesT ON ActionsT.Action_ID = LinkAttributeValueToEntitiesT.Entity_ID " & _
       "WHERE (((LinkAttributeValueToEntitiesT.Entity_Type_ID)=8 AND (LinkAttributeValueToEntitiesT.Entity_ID) = " & EntityIDArg & ")))  AS Q1 " & _
       "ON [LInkReferenceID(EntityTypeID)ToAttributeIDsT].Attribute_ID = Q1.Attribute_ID) ON AttributesT.Attribute_ID = [LInkReferenceID(EntityTypeID)ToAttributeIDsT].Attribute_ID " & _
       "WHERE (((Q1.Attribute_ID) Is Null) And (([LInkReferenceID(EntityTypeID)ToAttributeIDsT].Entity_Type_ID) = 8) And (([LInkReferenceID(EntityTypeID)ToAttributeIDsT].Reference_ID) = " & ReferenceEntityTypeIDArg & ")) " & _
       "Or ((([LInkReferenceID(EntityTypeID)ToAttributeIDsT].Entity_Type_ID) = 8) And (([LInkReferenceID(EntityTypeID)ToAttributeIDsT].Reference_ID) = " & ReferenceEntityTypeIDArg & ") " & _
       "And (([LInkReferenceID(EntityTypeID)ToAttributeIDsT].Can_Have_Multiple_Values) = Yes)) " & _
       "and AttributesT.Attribute_Description Like ""*" & TextToSearchArg & "*"" ORDER BY [LInkReferenceID(EntityTypeID)ToAttributeIDsT].View_ID;"
   CboxVar.RowSource = "SELECT  AttributesT.Attribute_Description, [LInkReferenceID(EntityTypeID)ToAttributeIDsT].Attribute_ID, Q1.Entity_ID " & _
       "FROM (AttributesT RIGHT JOIN [LInkReferenceID(EntityTypeID)ToAttributeIDsT] ON AttributesT.Attribute_ID =  [LInkReferenceID(EntityTypeID)ToAttributeIDsT].Attribute_ID )left JOIN " & _
       "(SELECT DISTINCT LinkAttributeValueToEntitiesT.Attribute_ID, LinkAttributeValueToEntitiesT.Entity_ID " & _
         "FROM " & _
         "LinkAttributeValueToEntitiesT WHERE LinkAttributeValueToEntitiesT.Entity_Type_ID = 8 AND LinkAttributeValueToEntitiesT.Entity_ID = " & EntityIDArg & ") AS Q1 " & _
         "ON [LInkReferenceID(EntityTypeID)ToAttributeIDsT].Attribute_ID = Q1.Attribute_ID " & _
         "WHERE (Q1.Attribute_ID is null And [LInkReferenceID(EntityTypeID)ToAttributeIDsT].Reference_ID = 1 AND [LInkReferenceID(EntityTypeID)ToAttributeIDsT].Entity_Type_ID = 8) " & _
         "OR (Q1.Attribute_ID is not null AND [LInkReferenceID(EntityTypeID)ToAttributeIDsT].Can_Have_Multiple_Values = true And [LInkReferenceID(EntityTypeID)ToAttributeIDsT].Reference_ID = 1 " & _
         "AND [LInkReferenceID(EntityTypeID)ToAttributeIDsT].Entity_Type_ID = 8) ORDER BY [LInkReferenceID(EntityTypeID)ToAttributeIDsT].View_ID "
            CboxVar.ListWidth = CboxVar.ListWidth
          Exit Sub
        Else
        
      If AttributeFilterOptionsArg = 1 Then
      
       CboxVar.RowSource = "SELECT DISTINCT AttributesT.Attribute_Description, LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "FROM " & _
         "((SELECT DISTINCT LinkAttributeValueToEntitiesT.Attribute_ID, ActionsT.Action_Type_ID, AttributesT.Attribute_Description " & _
         "FROM " & _
         "AttributesT INNER JOIN (LinkAttributeValueToEntitiesT LEFT JOIN ActionsT ON LinkAttributeValueToEntitiesT.Entity_ID = ActionsT.Action_ID) " & _
         "ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "WHERE LinkAttributeValueToEntitiesT.Entity_Type_ID = 8 " & _
         "AND ActionsT.Action_ID =  " & EntityIDArg & _
         "AND AttributesT.Entities_May_Have_Multiple_Values = No)  AS CurrentEntityAttributesQ " & _
         "RIGHT JOIN ((ActionsT LEFT JOIN LinkAttributeValueToEntitiesT ON ActionsT.Action_ID = LinkAttributeValueToEntitiesT.Entity_ID) " & _
         "left join AttributesT on LinkAttributeValueToEntitiesT.Attribute_ID = AttributesT.Attribute_ID) " & _
         "ON CurrentEntityAttributesQ.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID) " & _
         "WHERE CurrentEntityAttributesQ.Attribute_ID Is Null AND LinkAttributeValueToEntitiesT.Entity_Type_ID = 8 " & _
         "AND ActionsT.Action_Type_ID = " & DLookup("ActionsT.Action_Type_ID", "ActionsT", "ActionsT.Action_ID = " & EntityIDArg) & _
         "and AttributesT.Attribute_Description Like ""*" & TextToSearchArg & "*"" ORDER BY AttributesT.Attribute_Description;"
            
            CboxVar.ListWidth = CboxVar.ListWidth
          Exit Sub
        Else

        If AttributeFilterOptionsArg = 2 Then
               
         CboxVar.RowSource = "SELECT DISTINCT AttributesT.Attribute_Description, LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "FROM " & _
         "((SELECT DISTINCT LinkAttributeValueToEntitiesT.Attribute_ID, ActionsT.Action_Type_ID, AttributesT.Attribute_Description " & _
         "FROM " & _
         "AttributesT INNER JOIN (LinkAttributeValueToEntitiesT LEFT JOIN ActionsT ON LinkAttributeValueToEntitiesT.Entity_ID = ActionsT.Action_ID) " & _
         "ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "WHERE LinkAttributeValueToEntitiesT.Entity_Type_ID = 8 " & _
         "AND ActionsT.Action_ID =  " & EntityIDArg & _
         "AND AttributesT.Entities_May_Have_Multiple_Values = No)  AS CurrentEntityAttributesQ " & _
         "RIGHT JOIN ((ActionsT LEFT JOIN LinkAttributeValueToEntitiesT ON ActionsT.Action_ID = LinkAttributeValueToEntitiesT.Entity_ID) " & _
         "left join AttributesT on LinkAttributeValueToEntitiesT.Attribute_ID = AttributesT.Attribute_ID) " & _
         "ON CurrentEntityAttributesQ.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID) " & _
         "WHERE CurrentEntityAttributesQ.Attribute_ID Is Null AND LinkAttributeValueToEntitiesT.Entity_Type_ID = 8 " & _
         "AND AttributesT.Attribute_Description Like ""*" & TextToSearchArg & "*"" ORDER BY AttributesT.Attribute_Description;"
       
             CboxVar.ListWidth = CboxVar.ListWidth
           Exit Sub
         Else
         
        If AttributeFilterOptionsArg = 3 Then
       
             CboxVar.RowSource = "SELECT AttributesT.Attribute_Description, AttributesT.Attribute_ID " & _
             "FROM (SELECT DISTINCT LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "FROM " & _
         "AttributesT INNER JOIN (LinkAttributeValueToEntitiesT LEFT JOIN ActionsT ON LinkAttributeValueToEntitiesT.Entity_ID = ActionsT.Action_ID) " & _
         "ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "WHERE LinkAttributeValueToEntitiesT.Entity_Type_ID=8 " & _
         "AND ActionsT.Action_ID = " & EntityIDArg & _
         "AND AttributesT.Entities_May_Have_Multiple_Values=No) AS CurrentEntityAttributesQ " & _
         "RIGHT JOIN AttributesT ON CurrentEntityAttributesQ.Attribute_ID = AttributesT.Attribute_ID " & _
             "WHERE CurrentEntityAttributesQ.Attribute_ID Is Null and AttributesT.Attribute_Description Like ""*" & TextToSearchArg & "*"" order by AttributesT.Attribute_Description ;"
       
             CboxVar.ListWidth = CboxVar.ListWidth
            Exit Sub
           End If
          End If
         End If
         End If
Case 9 'Protocols
       If AttributeFilterOptionsArg = 0 Then
      
      ' CboxVar.RowSource = "SELECT  AttributesT.Attribute_Description, [LInkReferenceID(EntityTypeID)ToAttributeIDsT].Attribute_ID " & _
       "FROM AttributesT RIGHT JOIN ([LInkReferenceID(EntityTypeID)ToAttributeIDsT] LEFT JOIN (SELECT distinct LinkAttributeValueToEntitiesT.Link_Attribute_Value_To_Entity_ID, LinkAttributeValueToEntitiesT.Entity_Type_ID, " & _
       "LinkAttributeValueToEntitiesT.Attribute_ID " & _
       "FROM LinkAttributeValueToEntitiesT " & _
       "WHERE (((LinkAttributeValueToEntitiesT.Entity_Type_ID)=9 AND (LinkAttributeValueToEntitiesT.Entity_ID) = " & EntityIDArg & ")))  AS Q1 " & _
       "ON [LInkReferenceID(EntityTypeID)ToAttributeIDsT].Attribute_ID = Q1.Attribute_ID) ON AttributesT.Attribute_ID = [LInkReferenceID(EntityTypeID)ToAttributeIDsT].Attribute_ID " & _
       "WHERE ((Q1.Attribute_ID) Is Null) " & _
       "And  (([LInkReferenceID(EntityTypeID)ToAttributeIDsT].Entity_Type_ID) = 9) " & _
       "And  (([LInkReferenceID(EntityTypeID)ToAttributeIDsT].Reference_ID) = " & ReferenceEntityTypeIDArg & ") " & _
       "And (([LInkReferenceID(EntityTypeID)ToAttributeIDsT].Can_Have_Multiple_Values) = Yes) " & _
       "and AttributesT.Attribute_Description Like ""*" & TextToSearchArg & "*"" " & _
       "ORDER BY [LInkReferenceID(EntityTypeID)ToAttributeIDsT].View_ID;"
       
       'CboxVar.RowSource = "SELECT  AttributesT.Attribute_Description, [LInkReferenceID(EntityTypeID)ToAttributeIDsT].Attribute_ID " & _
       "FROM AttributesT RIGHT JOIN ([LInkReferenceID(EntityTypeID)ToAttributeIDsT] LEFT JOIN " & _
       "(SELECT DISTINCT LinkAttributeValueToEntitiesT.Attribute_ID, ProtocolsT.Protocol_Type_ID, AttributesT.Attribute_Description " & _
         "FROM " & _
         "AttributesT INNER JOIN (LinkAttributeValueToEntitiesT LEFT JOIN ProtocolsT ON LinkAttributeValueToEntitiesT.Entity_ID = ProtocolsT.Protocol_ID) " & _
         "ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "WHERE LinkAttributeValueToEntitiesT.Entity_Type_ID = 9 " & _
         "AND ProtocolsT.Protocol_ID =  " & EntityIDArg & _
         " AND AttributesT.Entities_May_Have_Multiple_Values = false)  AS Q1 " & _
       "ON [LInkReferenceID(EntityTypeID)ToAttributeIDsT].Attribute_ID = Q1.Attribute_ID) ON AttributesT.Attribute_ID = [LInkReferenceID(EntityTypeID)ToAttributeIDsT].Attribute_ID " & _
       "WHERE ((Q1.Attribute_ID) Is Null) " & _
       "And  (([LInkReferenceID(EntityTypeID)ToAttributeIDsT].Entity_Type_ID) = 9) " & _
       "And  (([LInkReferenceID(EntityTypeID)ToAttributeIDsT].Reference_ID) = " & ReferenceEntityTypeIDArg & ") " & _
       "and AttributesT.Attribute_Description Like ""*" & TextToSearchArg & "*"" " & _
       "ORDER BY [LInkReferenceID(EntityTypeID)ToAttributeIDsT].View_ID;"
       
       CboxVar.RowSource = "SELECT  AttributesT.Attribute_Description, [LInkReferenceID(EntityTypeID)ToAttributeIDsT].Attribute_ID, Q1.Entity_ID " & _
       "FROM (AttributesT RIGHT JOIN [LInkReferenceID(EntityTypeID)ToAttributeIDsT] ON AttributesT.Attribute_ID =  [LInkReferenceID(EntityTypeID)ToAttributeIDsT].Attribute_ID )left JOIN " & _
       "(SELECT DISTINCT LinkAttributeValueToEntitiesT.Attribute_ID, LinkAttributeValueToEntitiesT.Entity_ID " & _
         "FROM " & _
         "LinkAttributeValueToEntitiesT WHERE LinkAttributeValueToEntitiesT.Entity_Type_ID = 9 AND LinkAttributeValueToEntitiesT.Entity_ID = " & EntityIDArg & ") AS Q1 " & _
         "ON [LInkReferenceID(EntityTypeID)ToAttributeIDsT].Attribute_ID = Q1.Attribute_ID " & _
         "WHERE (Q1.Attribute_ID is null And [LInkReferenceID(EntityTypeID)ToAttributeIDsT].Reference_ID = 1 AND [LInkReferenceID(EntityTypeID)ToAttributeIDsT].Entity_Type_ID = 9) " & _
         "OR (Q1.Attribute_ID is not null AND [LInkReferenceID(EntityTypeID)ToAttributeIDsT].Can_Have_Multiple_Values = true And [LInkReferenceID(EntityTypeID)ToAttributeIDsT].Reference_ID = 1 " & _
         "AND [LInkReferenceID(EntityTypeID)ToAttributeIDsT].Entity_Type_ID = 9) ORDER BY [LInkReferenceID(EntityTypeID)ToAttributeIDsT].View_ID "
         
         '"WHERE (Q1.Attribute_ID is null OR (Q1.Attribute_ID is not null AND AttributesT.Entities_May_Have_Multiple_Values = true)) And [LInkReferenceID(EntityTypeID)ToAttributeIDsT].Reference_ID = " & ReferenceEntityTypeIDArg & " ORDER BY [LInkReferenceID(EntityTypeID)ToAttributeIDsT].View_ID "

         '"WHERE Q1.Attribute_ID is null OR (Q1.Attribute_ID is not null AND AttributesT.Entities_May_Have_Multiple_Values = true) ORDER BY [LInkReferenceID(EntityTypeID)ToAttributeIDsT].View_ID "
 'Debug.Print "CboxVar.RowSource = " & CboxVar.RowSource
 'να αντικταστησω την τελευταια γραμμη του ανωτερω query με αυτη "WHERE (Q1.Attribute_ID is null OR (Q1.Attribute_ID is not null AND AttributesT.Entities_May_Have_Multiple_Values = true) And (([LInkReferenceID(EntityTypeID)ToAttributeIDsT].Reference_ID) = " & ReferenceEntityTypeIDArg & ") ORDER BY [LInkReferenceID(EntityTypeID)ToAttributeIDsT].View_ID "

            CboxVar.ListWidth = CboxVar.ListWidth
          Exit Sub
        Else
      If AttributeFilterOptionsArg = 1 Then

       CboxVar.RowSource = "SELECT DISTINCT AttributesT.Attribute_Description, LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "FROM " & _
         "((SELECT DISTINCT LinkAttributeValueToEntitiesT.Attribute_ID, ProtocolsT.Protocol_Type_ID, AttributesT.Attribute_Description " & _
         "FROM " & _
         "AttributesT INNER JOIN (LinkAttributeValueToEntitiesT LEFT JOIN ProtocolsT ON LinkAttributeValueToEntitiesT.Entity_ID = ProtocolsT.Protocol_ID) " & _
         "ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "WHERE LinkAttributeValueToEntitiesT.Entity_Type_ID = 9 " & _
         "AND ProtocolsT.Protocol_ID =  " & EntityIDArg & _
         "AND AttributesT.Entities_May_Have_Multiple_Values = No)  AS CurrentEntityAttributesQ " & _
         "RIGHT JOIN ((ProtocolsT LEFT JOIN LinkAttributeValueToEntitiesT ON ProtocolsT.Protocol_ID = LinkAttributeValueToEntitiesT.Entity_ID) " & _
         "left join AttributesT on LinkAttributeValueToEntitiesT.Attribute_ID = AttributesT.Attribute_ID) " & _
         "ON CurrentEntityAttributesQ.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID) " & _
         "WHERE CurrentEntityAttributesQ.Attribute_ID Is Null AND LinkAttributeValueToEntitiesT.Entity_Type_ID = 9 " & _
         "AND ProtocolsT.Protocol_Type_ID = " & DLookup("ProtocolsT.Protocol_Type_ID", "ProtocolsT", "ProtocolsT.Protocol_ID = " & EntityIDArg) & _
         "and AttributesT.Attribute_Description Like ""*" & TextToSearchArg & "*"" ORDER BY AttributesT.Attribute_Description;"
            
            CboxVar.ListWidth = CboxVar.ListWidth
          Exit Sub
        Else

        If AttributeFilterOptionsArg = 2 Then
               
         CboxVar.RowSource = "SELECT DISTINCT AttributesT.Attribute_Description, LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "FROM " & _
         "((SELECT DISTINCT LinkAttributeValueToEntitiesT.Attribute_ID, ProtocolsT.Protocol_Type_ID, AttributesT.Attribute_Description " & _
         "FROM " & _
         "AttributesT INNER JOIN (LinkAttributeValueToEntitiesT LEFT JOIN ProtocolsT ON LinkAttributeValueToEntitiesT.Entity_ID = ProtocolsT.Protocol_ID) " & _
         "ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "WHERE LinkAttributeValueToEntitiesT.Entity_Type_ID = 9 " & _
         "AND ProtocolsT.Protocol_ID =  " & EntityIDArg & _
         "AND AttributesT.Entities_May_Have_Multiple_Values = No)  AS CurrentEntityAttributesQ " & _
         "RIGHT JOIN ((ProtocolsT LEFT JOIN LinkAttributeValueToEntitiesT ON ProtocolsT.Protocol_ID = LinkAttributeValueToEntitiesT.Entity_ID) " & _
         "left join AttributesT on LinkAttributeValueToEntitiesT.Attribute_ID = AttributesT.Attribute_ID) " & _
         "ON CurrentEntityAttributesQ.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID) " & _
         "WHERE CurrentEntityAttributesQ.Attribute_ID Is Null AND LinkAttributeValueToEntitiesT.Entity_Type_ID = 9 " & _
         "AND AttributesT.Attribute_Description Like ""*" & TextToSearchArg & "*"" ORDER BY AttributesT.Attribute_Description;"
       
             CboxVar.ListWidth = CboxVar.ListWidth
           Exit Sub
         Else
         
        If AttributeFilterOptionsArg = 3 Then
       
             CboxVar.RowSource = "SELECT AttributesT.Attribute_Description, AttributesT.Attribute_ID " & _
             "FROM (SELECT DISTINCT LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "FROM " & _
         "AttributesT INNER JOIN (LinkAttributeValueToEntitiesT LEFT JOIN ProtocolsT ON LinkAttributeValueToEntitiesT.Entity_ID = ProtocolsT.Protocol_ID) " & _
         "ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID " & _
         "WHERE LinkAttributeValueToEntitiesT.Entity_Type_ID=9 " & _
         "AND ProtocolsT.Protocol_ID = " & EntityIDArg & _
         "AND AttributesT.Entities_May_Have_Multiple_Values=No) AS CurrentEntityAttributesQ " & _
         "RIGHT JOIN AttributesT ON CurrentEntityAttributesQ.Attribute_ID = AttributesT.Attribute_ID " & _
             "WHERE CurrentEntityAttributesQ.Attribute_ID Is Null and AttributesT.Attribute_Description Like ""*" & TextToSearchArg & "*"" order by AttributesT.Attribute_Description ;"
       
             CboxVar.ListWidth = CboxVar.ListWidth
            Exit Sub
           End If
          End If
         End If
        End If
Case 11 'Countries

End Select

ExitProcedure:
Exit Sub
   
ErrorHandler:
Select Case Err.Number
        Case 2185
        Response = acDataErrContinue
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: RowsourceWithAttributeID" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
End Sub

Public Sub RowsourceWithAttributeValue(frm As Form, Cbox As Control, TextToSearchArg As String, EntityTypeIDArg As Integer, EntityIDArg As Long, AttributrIDArg As Long, Optional AttributeValueFilterOptionsArg As Integer)
                                      
Debug.Print "Exec Priority - " & "Attributes Module - " & "RowsourceWithAttributeValue " & Time()

'On Error GoTo ErrorHandler

Dim CboxVar As ComboBox

Set CboxVar = Cbox


Select Case DLookup("Data_type", "AttributesT", "AttributesT.Attribute_ID = " & AttributrIDArg)
  
   Case 9 ' text
        
        CboxVar.RowSource = "Select Attribute_Value_String, Attribute_Value_ID from AttributeValuesListsT where AttributeValuesListsT.Attribute_ID = " & frm!AttributeIDCbo & _
        " AND AttributeValuesListsT.Attribute_Value_String LIKE ""*" & TextToSearchArg & "*"" order by Attribute_Value_String"
                    
    Case 6 ' number
      
    If IsNull(frm!EntityTypeIDForRelevantTablePKFieldTbox) Then
       
        CboxVar.RowSource = "Select Attribute_Value_Number, Attribute_Value_ID from AttributeValuesListsT where AttributeValuesListsT.Attribute_ID = " & frm!AttributeIDCbo & _
        " AND AttributeValuesListsT.Attribute_Value_Number LIKE ""*" & TextToSearchArg & "*"" order by Attribute_Value_Number"
    
   Else
       
        If DLookup("Feed_list_Available", "AttributesT", "AttributesT.Attribute_ID = AttributeIDCbo.value") = True Then
        CboxVar.RowSource = "Select query1.Attribute_Value_Number, query1.DESCRIPTION, query1.[Type ID] from " & _
        " (Select AttributeValuesListsT.Attribute_Value_Number, SingleRelatedEntityDescription(Frm!EntityTypeIDForRelevantTablePKFieldTbox, Attribute_Value_Number) AS DESCRIPTION, SingleRelatedEntityTypeDescription(Frm!EntityTypeIDForRelevantTablePKFieldTbox, Attribute_Value_Number) AS [Type ID], View_Sequence_Number " & _
        " from AttributeValuesListsT where AttributeValuesListsT.Attribute_ID = " & frm!AttributeIDCbo & ") as query1 " & _
        " Where query1.DESCRIPTION LIKE ""*" & TextToSearchArg & "*"" order by query1.View_Sequence_Number"
    
        Else
        
         Select Case frm!EntityTypeIDForRelevantTablePKFieldTbox
         Case 1
         CboxVar.RowSource = "Select Product_ID, Product_Description, Type_Description from ProductsSimpleQ where ProductsSimpleQ.Product_Description LIKE ""*" & TextToSearchArg & "*""  " & _
         "AND ProductsSimpleQ.Product_Type_ID " & IIf(IsNull(frm!TypeOfEntityTypeIDTbox), " Like ""**"" ", " = " & frm!TypeOfEntityTypeIDTbox) & " order by ProductsSimpleQ.Product_Description"
         Case 2
         CboxVar.RowSource = "Select TransactorsWithBasicTransactorsDescriptionQ.Transactor_ID, TransactorsWithBasicTransactorsDescriptionQ.Basic_Transactor_Description, " & _
         "TransactorsWithBasicTransactorsDescriptionQ.Transactor_Type_Desription from TransactorsWithBasicTransactorsDescriptionQ " & _
         "Where TransactorsWithBasicTransactorsDescriptionQ.Basic_Transactor_Description LIKE ""*" & TextToSearchArg & "*"" AND TransactorsWithBasicTransactorsDescriptionQ.Transactor_Type_ID " & _
         IIf(IsNull(frm!TypeOfEntityTypeIDTbox), " Like ""**"" ", " = " & frm!TypeOfEntityTypeIDTbox) & " order by TransactorsWithBasicTransactorsDescriptionQ.Basic_Transactor_Description"
         End Select
        End If
 

    End If
           
    'Case 12 ' boolean, all that is needed have been arranged by AttributeValuesCboProperties Sub
        
              
    Case 8 ' date
       
        CboxVar.RowSource = "Select Attribute_Value_Date, Attribute_Value_ID from AttributeValuesListsT where AttributeValuesListsT.Attribute_ID = " & frm!AttributeIDCbo & " order by Attribute_Value_Date"
       
        
        
    Case 19 ' Time
        
        CboxVar.RowSource = "Select Attribute_Value_Time, Attribute_Value_ID from AttributeValuesListsT where AttributeValuesListsT.Attribute_ID = " & frm!AttributeIDCbo & " order by Attribute_Value_Time"
        
        
        
    Case 20 ' Timestamp
        
        CboxVar.RowSource = "Select Attribute_Value_Timestamp, Attribute_Value_ID from AttributeValuesListsT where AttributeValuesListsT.Attribute_ID = " & frm!AttributeIDCbo & " order by Attribute_Value_Timestamp"
    
        
        
  End Select
ExitProcedure:
Exit Sub
   
ErrorHandler:
Select Case Err.Number
        Case 2185
        Response = acDataErrContinue
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: RowsourceWithAttributeValue" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select

End Sub

Public Sub AttributeValuesCboProperties(frm As Form, AttributeIDArg As Long, Cbox As Control, EntityTypeIDArg As Integer, EntityIDArg As Long, Optional AttributeValueFilterOptionsArg As Integer)
Debug.Print "Exec Priority - " & "Attributes Module - " & "AttributeValuesCboProperties " & Time()

'On Error GoTo ErrorHandler

Dim CboxVar As Control

Set CboxVar = Cbox


Select Case DLookup("Data_type", "AttributesT", "AttributesT.Attribute_ID = " & AttributeIDArg)
  
     Case 9 ' text
        CboxVar.RowSourceType = "Table/Query"
        CboxVar.ColumnCount = 1
        CboxVar.BoundColumn = 1
        CboxVar.ColumnWidths = ""
        'Frm!StringCbo = Null
        frm!NumberCbo = Null
        frm!BooleanCheckbox = Null
        frm!DateCbo = Null
        frm!TimeCbo = Null
        frm!TImestampCbo = Null
        CboxVar.ControlSource = "Attribute_Value_String"
        CboxVar.LimitToList = False
        If DLookup("Feed_list_Available", "AttributesT", "AttributesT.Attribute_ID = " & AttributeIDArg) = True Then
          If DLookup("Limited_To_List", "AttributesT", "AttributesT.Attribute_ID = " & AttributeIDArg) = True Then
          CboxVar.LimitToList = True
        End If
        End If
                
                         
    Case 6 ' number
      
    If IsNull(frm!EntityTypeIDForRelevantTablePKFieldTbox) Then
        CboxVar.RowSourceType = "Table/Query"
        CboxVar.ColumnCount = 1
        CboxVar.BoundColumn = 1
        CboxVar.ColumnWidths = ""
        frm!StringCbo = Null
        'Frm!NumberCbo = Null
        frm!BooleanCheckbox = Null
        frm!DateCbo = Null
        frm!TimeCbo = Null
        frm!TImestampCbo = Null
        CboxVar.ControlSource = "Attribute_Value_Number"
        CboxVar.LimitToList = False
        If DLookup("Feed_list_Available", "AttributesT", "AttributesT.Attribute_ID = " & AttributeIDArg) = True Then
          If DLookup("Limited_To_List", "AttributesT", "AttributesT.Attribute_ID = " & AttributeIDArg) = True Then
            CboxVar.LimitToList = True
          End If
        End If
        
   Else
       
       CboxVar.RowSourceType = "Table/Query"
       CboxVar.ColumnCount = 3
       CboxVar.BoundColumn = 1
       CboxVar.ColumnWidths = "1,5 cm;15 cm;10 cm"
       'CboxVar.ListWidth = ""
        frm!StringCbo = Null
       'Frm!NumberCbo = Null
        frm!BooleanCheckbox = Null
        frm!DateCbo = Null
        frm!TimeCbo = Null
        frm!TImestampCbo = Null
        CboxVar.ControlSource = "Attribute_Value_Number"
        CboxVar.LimitToList = False
  End If
        

    Case 12 ' boolean
        CboxVar.RowSourceType = "value list"
        CboxVar.ColumnCount = 1
        CboxVar.BoundColumn = 1
        CboxVar.ColumnWidths = ""
        CboxVar.RowSource = "yes;no"
        CboxVar.SetFocus
        CboxVar.ControlSource = "Attribute_Value_Boolean"
        CboxVar.LimitToList = True
        frm!StringCbo = Null
        frm!NumberCbo = Null
       ' Frm!BooleanCheckbox = Null
        frm!DateCbo = Null
        frm!TimeCbo = Null
        frm!TImestampCbo = Null
      
              
    Case 8 ' date
        CboxVar.RowSourceType = "Table/Query"
        CboxVar.ColumnCount = 1
        CboxVar.BoundColumn = 1
        CboxVar.ColumnWidths = ""
        frm!StringCbo = Null
        frm!NumberCbo = Null
        frm!BooleanCheckbox = Null
        'Frm!DateCbo = Null
        frm!TimeCbo = Null
        frm!TImestampCbo = Null
        CboxVar.ControlSource = "Attribute_Value_Date"
        frm!DateCbo.LimitToList = False
        CboxVar.LimitToList = False
        If DLookup("Feed_list_Available", "AttributesT", "AttributesT.Attribute_ID = AttributeIDArg") = True Then
        If DLookup("Limited_To_List", "AttributesT", "AttributesT.Attribute_ID = AttributeIDArg") = True Then
        CboxVar.LimitToList = True
        End If
        End If
    
        
        
    Case 19 ' Time
        CboxVar.RowSourceType = "Table/Query"
        CboxVar.ColumnCount = 1
        CboxVar.BoundColumn = 1
        CboxVar.ColumnWidths = ""
        frm!StringCbo = Null
        frm!NumberCbo = Null
        frm!BooleanCheckbox = Null
        frm!DateCbo = Null
        'Frm!TimeCbo = Null
        frm!TImestampCbo = Null
        CboxVar.ControlSource = "Attribute_Value_Time"
        frm!TimeCbo.LimitToList = False
        CboxVar.LimitToList = False
        If DLookup("Feed_list_Available", "AttributesT", "AttributesT.Attribute_ID = AttributeIDArg") = True Then
        If DLookup("Limited_To_List", "AttributesT", "AttributesT.Attribute_ID = AttributeIDArg") = True Then
        frm!TimeCbo.LimitToList = True
        CboxVar.LimitToList = True
        End If
        End If

        
        
    Case 20 ' Timestamp
        CboxVar.RowSourceType = "Table/Query"
        CboxVar.ColumnCount = 1
        CboxVar.BoundColumn = 1
        CboxVar.ColumnWidths = ""
        frm!StringCbo = Null
        frm!NumberCbo = Null
        frm!BooleanCheckbox = Null
        frm!DateCbo = Null
        frm!TimeCbo = Null
        'Frm!TImestampCbo = Null
        CboxVar.ControlSource = "Attribute_Value_TImestamp"
        frm!TimeCbo.LimitToList = False
        CboxVar.LimitToList = False
        If DLookup("Feed_list_Available", "AttributesT", "AttributesT.Attribute_ID = AttributeIDArg") = True Then
        If DLookup("Limited_To_List", "AttributesT", "AttributesT.Attribute_ID = AttributeIDArg") = True Then
        frm!TImestampCbo.LimitToList = True
        CboxVar.LimitToList = True
        End If
        End If
    
        
        
  End Select

ExitProcedure:
Exit Sub
   
ErrorHandler:
Select Case Err.Number
        Case 2185
        Response = acDataErrContinue
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: AttributeValuesCboProperties" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
End Sub

Public Sub RowsourceForAttributeIDFilterOptionsCbo(Cbox As Control, EntityTypeIDArg As Integer)
Debug.Print "Exec Priority - " & "Attributes Module - " & "RowsourceForAttributeIDFilterOptionsCbo " & Time()
On Error GoTo ErrorHandler

Select Case EntityTypeIDArg
  Case 1
   Cbox.RowSource = "1;Only Attributes from the same Product Type Category and Subcategory. ;2;Only Attributes from the same Product Type and Category. ;3;Only Attributes from the same Product Type.;4;Only Attributes from Products (Of All Types);5;All Attributes"
  Case 2
   Cbox.RowSource = "1;Only Attributes from the same Transactor Type.;2;Only Attributes from Transactors (Of All Types).;3;All Attributes."
  Case 3
  Cbox.RowSource = "1;Only Attributes from the same Issued Document Type.;2;Only Attributes from Issued Documents (Of All Types).;3;All Attributes."
  Case 4
  Cbox.RowSource = "1;Only Attributes from the same Transactor's previous records;2;Only Attributes from same Type of Transactors's previous records(Of All Document Types).;3;Only Attributes from All Types of Transactors's previous records;4;All Attributes."
  Case 5
  Cbox.RowSource = "1;Only Attributes from previous records of the same Product Type,Category and Subcategory.;2;Only Attributes from previous records of the same Product Type and Category.;3;Only Attributes from previous records of the same Product Type.;4;Only Attributes from previous records of Products.;5;All Attributes."
  Case 6
  Cbox.RowSource = "1;Only Attributes from the same Transaction Type.;2;Only Attributes from Transactions (Of All Types).;3;All Attributes."
  Case 8
  Cbox.RowSource = "0;Attributes from help list.;1;Only Attributes from the same Action Type.;2;Only Attributes from Actions (Of All Types).;3;All Attributes."
  Case 9
  Cbox.RowSource = "0;Attributes from help list.;1;Only Attributes from the same Protocol Type.;2;Only Attributes from Protocols (Of All Types).;3;All Attributes."
  Case 10
  Cbox.RowSource = "1;Only Attributes from the same Installation Type.;2;Only Attributes from Installations (Of All Types).;3;All Attributes"
 
End Select

ExitProcedure:
Exit Sub
   
ErrorHandler:
Select Case Err.Number
        Case 2185
        Response = acDataErrContinue
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: RowsourceForAttributeIDFilterOptionsCbo" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select

End Sub

Public Sub RowsourceForAttributeValuesFilterOptionsCbo(Cbox As Control, EntityTypeIDArg As Integer)
' when I will decide how I would like the AttributeValuesFilterOptionCbo (to each different AttributeValuesToEntityBasicWithAttrFiltersDSF that I create for every entity case) to be filtered, _
I will fill the code. Now, this routine is just to be hitted by all the different attribute subforms _
that I am about to create, so when I will decide to implement filter on Attribute values, I will not have to visit all these forms and make changes. I just fill this routine which already been hitted _
by each form to request rowsource for the AttributeValuesFilterOptionCbo'


Select Case EntityTypeIDArg
  Case 1
   Cbox.RowSource = ""
  Case 2
   Cbox.RowSource = ""
  Case 3
  Cbox.RowSource = ""
  Case 4
  Cbox.RowSource = ""
  Case 5
  Cbox.RowSource = ""
  Case 6
  Cbox.RowSource = ""
  Case 8
  Cbox.RowSource = ""
  Case 9
  Cbox.RowSource = ""
  Case 10
  Cbox.RowSource = ""
End Select
End Sub

Public Function CheckIfEntityMayHaveMultipleAttributeValues(AttributeIDArg As Long, EntityIDArg As Long, EntityTpeIDArg As Integer) As Long
'this function checks 1st the state of "Entities_May_Have_Multiple_Values" field of AttributesT, and then 2nd checks if the entity has already a value for the specific attribute. _
If the field is true then it returns 0 (which means that new record in the table LinkAttributeValueToEntitiesT is permitted). _
If the field is false and entity has not any record in the LinkAttributeValueToEntitiesT for this attribute then it returns again 0. _
In all other cases, it returns the LinkAttributeValueToEntitiesT.Link_Attribute_Value_To_Entity_ID as it is may needed for edit this record.
Debug.Print "Exec Priority - " & "Attributes Module - " & "CheckIfEntityMayHaveMultipleAttributeValues " & Time()
On Error GoTo ErrorHandler

Dim db As DAO.Database
Dim rst As DAO.Recordset

CheckIfEntityMayHaveMultipleAttributeValues = 0

If DLookup("AttributesT.Entities_May_Have_Multiple_Values", "AttributesT", "Attribute_ID = " & AttributeIDArg) = True Then
CheckIfEntityMayHaveMultipleAttributeValues = 0
GoTo ExitProcedure
Else
Set db = CurrentDb
Set rst = db.OpenRecordset("Select Link_Attribute_Value_To_Entity_ID from LinkAttributeValueToEntitiesT where Entity_Type_ID = " & EntityTpeIDArg & _
" AND Entity_ID = " & EntityIDArg & " AND Attribute_ID = " & AttributeIDArg)

   If Not rst.EOF Then
   rst.MoveLast
   rst.MoveFirst
   CheckIfEntityMayHaveMultipleAttributeValues = rst("Link_Attribute_Value_To_Entity_ID")
   Else
   CheckIfEntityMayHaveMultipleAttributeValues = 0
   End If
End If

ExitProcedure:
If Not rst Is Nothing Then
rst.Close
Set rst = Nothing
End If

If Not db Is Nothing Then
db.Close
Set db = Nothing
End If

Exit Function
   
ErrorHandler:
Select Case Err.Number

        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: CheckIfEntityMayHaveMultipleAttributeValues" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select

End Function

Public Sub LinkAttributeValueToEntitiesFinalizeRecordAfterSave(EntityIDArg As Long, EntityTypeIDArg As Integer)
Debug.Print "Exec Priority - " & "Attributes Module - " & "LinkAttributeValueToEntitiesFinalizeRecordAfterSave " & Time()
'On Error GoTo ErrorHandler

Dim db As DAO.Database
Dim rst As DAO.Recordset

Set db = CurrentDb
Set rst = db.OpenRecordset("Select * from LinkAttributeValueToEntitiesT Where Entity_Type_ID = " & EntityTypeIDArg & " And Entity_ID = " & EntityIDArg)

If Not rst.EOF Then
rst.MoveLast
rst.MoveFirst
   Do Until rst.EOF
     rst.Edit
     rst("Is_New") = False
     rst("LinkAttributeValueToEntityID_Backup_ID") = Null
     rst.Update
     LinkSubAttributeValueToEntitiesFinalizeRecordAfterSave (rst("Link_Attribute_Value_To_Entity_ID"))
     rst.MoveNext
   Loop
End If

ExitProcedure:
If Not rst Is Nothing Then
rst.Close
Set rst = Nothing
End If

If Not db Is Nothing Then
db.Close
Set db = Nothing
End If

Exit Sub
   
ErrorHandler:
Select Case Err.Number

        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: LinkAttributeValueToEntitiesFinalizeRecordAfterSave" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select

End Sub

Public Sub LinkSubAttributeValueToEntitiesFinalizeRecordAfterSave(EntityIDArg As Long)
Debug.Print "Exec Priority - " & "Attributes Module - " & "LinkSubAttributeValueToEntitiesFinalizeRecordAfterSave " & Time()
'On Error GoTo ErrorHandler

Dim db As DAO.Database
Dim rst As DAO.Recordset

Set db = CurrentDb
Set rst = db.OpenRecordset("Select * from LinkAttributeValueToEntitiesT Where Entity_Type_ID = 7 And Entity_ID = " & EntityIDArg)

If Not rst.EOF Then
rst.MoveLast
rst.MoveFirst
   Do Until rst.EOF
     rst.Edit
     rst("Is_New") = False
     rst("LinkAttributeValueToEntityID_Backup_ID") = Null
     rst.Update
     rst.MoveNext
   Loop
End If

ExitProcedure:
If Not rst Is Nothing Then
rst.Close
Set rst = Nothing
End If

If Not db Is Nothing Then
db.Close
Set db = Nothing
End If

Exit Sub
   
ErrorHandler:
Select Case Err.Number

        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: LinkSubAttributeValueToEntitiesFinalizeRecordAfterSave" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select

End Sub