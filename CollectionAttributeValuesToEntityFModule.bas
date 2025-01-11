Option Compare Database
Public AttributeValuesToEntityMainFCollection As New Collection


Public Sub OpenAttributeValuesToMainEntityFClient(EntityIDArg As Long, EntityTypeIDArg As Integer)
Debug.Print "Exec Priority - " & "CollectionAttributeValuesToEntityFModule - " & "OpenAttributeValuesToEntityMainFClient" & Time()
 
    'On Error GoTo Errorhandler

    ' Purpose: Open an independent instance of form AttributeValuesToMainEntityF.
    Dim frm As Form
    Dim VarEntityTypeDescription As String
             
        
        ' Open a new instance, show it, and set a caption.
        Set frm = New Form_AttributeValuesToEntityMainF
        frm.RecordSource = AttributeValuesToMainEntityFRecordsource(EntityTypeIDArg, EntityIDArg)
        frm!AttributeValuesToEntityF.Form.RecordSource = AttributeValuesToEntityFRecordsource(EntityTypeIDArg)
        frm.Visible = True
        frm.Caption = "Attribute values to entities Form, opened " & Now() & ", (ID = " & frm.Hwnd & ")"

        ' Append it to our collection.
        AttributeValuesToEntityMainFCollection.Add Item:=frm, Key:=CStr(frm.Hwnd)

       
         frm.Requery
          Dim rst As DAO.Recordset
          Set rst = frm.RecordsetClone
                  rst.FindFirst "Entity_ID = " & EntityIDArg & " AND Entity_Type_ID = " & EntityTypeIDArg
                If Not frm.RecordsetClone.NoMatch Then
                     frm.Bookmark = frm.RecordsetClone.Bookmark
                Else
                     VarEntityTypeDescription = DLookup("EntitiesTypesToHaveAttributesT.Entities_To_Have_Attributes_Description", "EntitiesTypesToHaveAttributesT", "EntitiesTypesToHaveAttributesT.Entities_To_Have_Attributes_ID = " & EntityTypeIDArg)
                     MsgBox VarEntityTypeDescription & " with ID " & EntityIDArg & " not found having any attributes.", vbExclamation
                     GoTo ExitProcedure
                End If


ExitProcedure:

If Not frm Is Nothing Then
Set frm = Nothing
End If

If Not rst Is Nothing Then
rst.Close
Set rst = Nothing
End If

Exit Sub

ErrorHandler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: OpenAttributeValuesToMainEntityFClient" & vbCrLf & _
            "Error Description: " & Err.Description, vbOKOnly + vbCritical, "An Error has Occurred!"
    Resume ExitProcedure
 
End Sub
Function AttributeValuesToMainEntityFRecordsource(EntityTypeIDArg As Integer, EntityIDArg As Long) As String
Debug.Print "Exec Priority - " & "CollectionAttributeValuesToEntityFModule - " & "AttributeValuesToMainEntityFRecordsource" & Time()
 
On Error GoTo ErrorHandler

Select Case EntityTypeIDArg

  Case 1  ' products
   
    AttributeValuesToMainEntityFRecordsource = "SELECT ProductsSimpleQ1.Product_ID AS Entity_ID, ProductsSimpleQ1.EntityTypesToHaveAttributesID AS Entity_Type_ID, EntitiesTypesToHaveAttributesT.Entities_To_Have_Attributes_Description " & _
" from " & _
"(SELECT ProductsSimpleQ.*, 1 AS EntityTypesToHaveAttributesID " & _
"FROM ProductsSimpleQ where ProductsSimpleQ.Product_ID = " & EntityIDArg & ") as ProductsSimpleQ1 " & _
"INNER JOIN EntitiesTypesToHaveAttributesT on ProductsSimpleQ1.EntityTypesToHaveAttributesID = EntitiesTypesToHaveAttributesT.Entities_To_Have_Attributes_ID;"

    Case 2  ' transactors
    
   AttributeValuesToMainEntityFRecordsource = "SELECT TransactorsSimpleQ1.Transactor_ID AS Entity_ID, TransactorsSimpleQ1.EntityTypesToHaveAttributesID AS Entity_Type_ID, EntitiesTypesToHaveAttributesT.Entities_To_Have_Attributes_Description " & _
" from " & _
"(SELECT TransactorsWithBasicTransactorsDescriptionQ.*, 2 AS EntityTypesToHaveAttributesID " & _
"FROM TransactorsWithBasicTransactorsDescriptionQ where TransactorsWithBasicTransactorsDescriptionQ.Transactor_ID = " & EntityIDArg & ") as TransactorsSimpleQ1 " & _
"INNER JOIN EntitiesTypesToHaveAttributesT on TransactorsSimpleQ1.EntityTypesToHaveAttributesID = EntitiesTypesToHaveAttributesT.Entities_To_Have_Attributes_ID;"

    Case 3   ' Issued Document
    
      AttributeValuesToMainEntityFRecordsource = "SELECT IssuedDocumentSimpleQ1.Issued_Document_ID AS Entity_ID,  " & _
    " IssuedDocumentSimpleQ1.EntityTypesToHaveAttributesID AS Entity_Type_ID, " & _
    " EntitiesTypesToHaveAttributesT.Entities_To_Have_Attributes_Description " & _
    " from " & _
    "(SELECT IssuedDocumentSimpleQ.*, 3 AS EntityTypesToHaveAttributesID " & _
    "FROM IssuedDocumentSimpleQ where IssuedDocumentSimpleQ.Issued_Document_ID = " & EntityIDArg & ") as IssuedDocumentSimpleQ1 " & _
    "INNER JOIN EntitiesTypesToHaveAttributesT on IssuedDocumentSimpleQ1.EntityTypesToHaveAttributesID = EntitiesTypesToHaveAttributesT.Entities_To_Have_Attributes_ID;"

    Case 4   ' Document Financial Details
   

      AttributeValuesToMainEntityFRecordsource = "SELECT IssuedDocumentFinancialDetailsSimpleQ1.IssuedDocumentFinancialDetailsT.Issued_Document_Financial_Details_ID AS Entity_ID, " & _
    "IssuedDocumentFinancialDetailsSimpleQ1.EntityTypesToHaveAttributesID AS Entity_Type_ID, " & _
    " EntitiesTypesToHaveAttributesT.Entities_To_Have_Attributes_Description " & _
    " from " & _
    "(SELECT IssuedDocumentFinancialDetailsSimpleQ.*, 4 AS EntityTypesToHaveAttributesID " & _
    "FROM IssuedDocumentFinancialDetailsSimpleQ where IssuedDocumentFinancialDetailsSimpleQ.IssuedDocumentFinancialDetailsT.Issued_Document_Financial_Details_ID = " & EntityIDArg & ") as IssuedDocumentFinancialDetailsSimpleQ1 " & _
    "INNER JOIN EntitiesTypesToHaveAttributesT on IssuedDocumentFinancialDetailsSimpleQ1.EntityTypesToHaveAttributesID = EntitiesTypesToHaveAttributesT.Entities_To_Have_Attributes_ID;"

    
    Case 5   ' Document Product Details
    
      AttributeValuesToMainEntityFRecordsource = "SELECT IssuedDocumentProductDetailsSimpleQ1.IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID AS Entity_ID, " & _
    "IssuedDocumentProductDetailsSimpleQ1.EntityTypesToHaveAttributesID AS Entity_Type_ID, " & _
    " EntitiesTypesToHaveAttributesT.Entities_To_Have_Attributes_Description " & _
    " from " & _
    "(SELECT IssuedDocumentProductDetailsSimpleQ.*, 5 AS EntityTypesToHaveAttributesID " & _
    "FROM IssuedDocumentProductDetailsSimpleQ where IssuedDocumentProductDetailsSimpleQ.IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID = " & EntityIDArg & ") as IssuedDocumentProductDetailsSimpleQ1 " & _
    "INNER JOIN EntitiesTypesToHaveAttributesT on IssuedDocumentProductDetailsSimpleQ1.EntityTypesToHaveAttributesID = EntitiesTypesToHaveAttributesT.Entities_To_Have_Attributes_ID;"
    
    Case 6   ' Transaction Details
    
   AttributeValuesToMainEntityFRecordsource = "SELECT TransactionsQ1.Transaction_ID AS Entity_ID, TransactionsQ1.EntityTypesToHaveAttributesID AS Entity_Type_ID, EntitiesTypesToHaveAttributesT.Entities_To_Have_Attributes_Description " & _
" from " & _
"(SELECT TransactionsT.*, 6 AS EntityTypesToHaveAttributesID " & _
"FROM TransactionsT where TransactionsT.Transaction_ID = " & EntityIDArg & ") as TransactionsQ1 " & _
"INNER JOIN EntitiesTypesToHaveAttributesT on TransactionsQ1.EntityTypesToHaveAttributesID = EntitiesTypesToHaveAttributesT.Entities_To_Have_Attributes_ID;"
    
      Case 8   ' Actions
    
  AttributeValuesToMainEntityFRecordsource = "SELECT ActionsSimpleQ1.ActionsT.Action_ID AS Entity_ID, " & _
"ActionsSimpleQ1.EntityTypesToHaveAttributesID AS Entity_Type_ID, EntitiesTypesToHaveAttributesT.Entities_To_Have_Attributes_Description " & _
" from " & _
"(SELECT  ActionsT.*, 8 AS EntityTypesToHaveAttributesID " & _
"FROM ActionsT where ActionsT.Action_ID = " & EntityIDArg & ") as ActionsSimpleQ1 " & _
"INNER JOIN EntitiesTypesToHaveAttributesT on ActionsSimpleQ1.EntityTypesToHaveAttributesID = EntitiesTypesToHaveAttributesT.Entities_To_Have_Attributes_ID;"

     Case 9   ' Protocols
    
   AttributeValuesToMainEntityFRecordsource = "SELECT ProtocolsSimpleQ1.ProtocolsT.Protocol_ID AS Entity_ID, " & _
"ProtocolsSimpleQ1.EntityTypesToHaveAttributesID AS Entity_Type_ID, EntitiesTypesToHaveAttributesT.Entities_To_Have_Attributes_Description " & _
" from " & _
"(SELECT ProtocolsT.*, 9 AS EntityTypesToHaveAttributesID " & _
"FROM ProtocolsT where ProtocolsT.Protocol_ID = " & EntityIDArg & ") as ProtocolsSimpleQ1 " & _
"INNER JOIN EntitiesTypesToHaveAttributesT on ProtocolsSimpleQ1.EntityTypesToHaveAttributesID = EntitiesTypesToHaveAttributesT.Entities_To_Have_Attributes_ID;"

     Case 10   ' Installations
    
   AttributeValuesToMainEntityFRecordsource = "SELECT InstallationsSimpleQ1.Installation_ID AS Entity_ID, " & _
"InstallationsSimpleQ1.EntityTypesToHaveAttributesID AS Entity_Type_ID, EntitiesTypesToHaveAttributesT.Entities_To_Have_Attributes_Description " & _
" from " & _
"(SELECT InstallationsT.*, 10 AS EntityTypesToHaveAttributesID " & _
"FROM InstallationsT where InstallationsT.Installation_ID = " & EntityIDArg & ") as InstallationsSimpleQ1 " & _
"INNER JOIN EntitiesTypesToHaveAttributesT on InstallationsSimpleQ1.EntityTypesToHaveAttributesID = EntitiesTypesToHaveAttributesT.Entities_To_Have_Attributes_ID;"


End Select
  
ExitProcedure:

Exit Function

ErrorHandler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: AttributeValuesToMainEntityFRecordsource" & vbCrLf & _
            "Error Description: " & Err.Description, vbOKOnly + vbCritical, "An Error has Occurred!"
    Resume ExitProcedure
  
 End Function
Public Function AttributeValuesToEntityFRecordsource(EntityTypeIDArg As Integer)
  Debug.Print "Exec Priority - " & "CollectionAttributeValuesToEntityFModule - " & "AttributeValuesToEntityFRecordsource " & Time()
   'On Error GoTo Errorhandler
   
 AttributeValuesToEntityFRecordsource = "SELECT LinkAttributeValueToEntitiesT.*, AttributesT.Attribute_Description, AttributesT.EntityTypeID_For_RelevantTablePKField, " & _
 "AttributesT.TypeOfEntityType_ID, " & _
 "CStr(Nz([Attribute_Value_String], """")) & IIf([EntityTypeID_For_RelevantTablePKField] Is Null, CStr(Nz([Attribute_Value_Number], """")), " & _
 "CStr(Nz(SingleRelatedEntityDescription([EntityTypeID_For_RelevantTablePKField], Nz([Attribute_Value_Number], 0)), """"))) & " & _
 "CStr(IIf([Attribute_Value_Boolean] = 0, ""No"", IIf([Attribute_Value_Boolean] = -1, ""Yes"", """"))) & CStr(Nz([Attribute_Value_Date], """")) & " & _
 "CStr(Nz([Attribute_Value_Time], """")) & CStr(Nz([Attribute_Value_TImestamp], """")) AS [Value], " & _
 "LinkAttributeValueToEntitiesT.Entity_Type_ID AS EntityTypeIDHeloColumn " & _
"FROM AttributesT INNER JOIN LinkAttributeValueToEntitiesT ON AttributesT.Attribute_ID = LinkAttributeValueToEntitiesT.Attribute_ID " & _
"WHERE LinkAttributeValueToEntitiesT.Entity_Type_ID = " & EntityTypeIDArg & " AND LinkAttributeValueToEntitiesT.Is_Deleted=False"

ExitProcedure:

Exit Function
   
ErrorHandler:
Select Case Err.Number
        
        
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: AttributeValuesToEntityFRecordsource" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select

End Function
Function CloseOneAttributeValuesToMainEntityFClient(HwndArg As Long)
  Debug.Print "Exec Priority - " & "CollectionAttributeValuesToEntityFModule - " & "CloseOneAttributeValuesToMainEntityFClient " & Time()
   'On Error GoTo Errorhandler
   
    'Purpose: Remove this instance from the clnClient collection.
    Dim obj As Object           'Object in clnClient
    Dim blnRemove As Boolean    'Flag to remove it.
    Dim frm As Form
    Dim i As Form
    
    For Each i In Forms
      If i.Hwnd = HwndArg Then
      Set frm = i
      Exit For
      End If
    Next i
    
    'Check if this instance is in the collection.
    '   (It won't be if form was opened directly, or code was reset.)
    For Each obj In AttributeValuesToEntityMainFCollection
        If obj.Hwnd = HwndArg Then
            blnRemove = True
            Exit For
        End If
    Next
    
       
    'Deassign the object before removing from collection.
    Set obj = Nothing
    If blnRemove Then
        AttributeValuesToEntityMainFCollection.Remove CStr(HwndArg)
        If CheckIfFormIsOpen(frm.Name) Then
            DoCmd.Close acForm, frm.Name, acSaveNo
        End If
    Else
        DoCmd.Close acForm, "AttributeValuesToEntityMainF", acSaveNo
    End If

ExitProcedure:
If Not obj Is Nothing Then
Set obj = Nothing
End If

Exit Function
   
ErrorHandler:
Select Case Err.Number
        
        
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: CloseOneAttributeValuesToMainEntityFClient" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select

End Function



Function CloseAllAttributeValuesToMainEntityFClient()
    'Purpose: Close all instances in the clnClient collection.
    'Note: Leaves the copy opened directly from database window.
Debug.Print "Exec Priority - " & "CollectionAttributeValuesToEntityFModule - " & "CloseAllAttributeValuesToMainEntityFClient" & Time()
On Error GoTo ErrorHandler

    Dim NumberOfMembers As Long
    Dim i As Long
    
    NumberOfMembers = AttributeValuesToEntityMainFCollection.Count
    For i = 1 To NumberOfMembers
        AttributeValuesToEntityMainFCollection.Remove 1
    Next
    
If CheckIfFormIsOpen("AttributeValuesToEntityMainF") Then
DoCmd.Close acForm, "AttributeValuesToEntityMainF", acSaveNo
End If
    
ExitProcedure:
Exit Function
   
ErrorHandler:
Select Case Err.Number
        
        
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: CloseAllAttributeValuesToMainEntityFClient" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select

End Function