Option Compare Database

Public Sub Delete_Transaction(TransactionIDToDelete As Long, OldOrNew As Integer, AskConfirmation As Boolean) ' 1=Only New, 2=Only Old, 3=All
Debug.Print "Exec Priority - " & "Delete Module - " & "Delete_Transaction " & Time()
'On Error GoTo ErrorHandler
Dim Response As Integer


If AskConfirmation = True Then
Response = MsgBox("You are about to DELETE a TRANSACTION record(s)! Do you want to proceed?", vbExclamation + vbYesNoCancel, "ATTENTION")
If Response <> vbYes Then
Exit Sub
End If
End If


 Dim db As DAO.Database
 Dim rstTransactionToDelete As DAO.Recordset
 Dim VarTransactionID As Long
 
  Set db = CurrentDb
  
 Select Case OldOrNew
Case 1
   varNewOrOldRecord = " Is_New = True AND "
Case 2
   varNewOrOldRecord = " Is_New = False AND "
Case 3
   varNewOrOldRecord = ""
End Select
   
 Set rstTransactionToDelete = db.OpenRecordset("Select * from TransactionsT where " & varNewOrOldRecord & " Transaction_ID = " & TransactionIDToDelete)
   If Not rstTransactionToDelete.EOF Then
   rstTransactionToDelete.MoveLast
   rstTransactionToDelete.MoveFirst
   
   Do Until rstTransactionToDelete.EOF
     VarTransactionID = rstTransactionToDelete("Transaction_ID")
      Call Delete_IssuedDocuments(VarTransactionID, 3, False)
      Call Delete_LinkAttributeValuesToEntities(VarTransactionID, 6, False)
      db.Execute "Insert INTO TransactionsT_DeletedRecords Select * from TransactionsT where Is_New = False AND Transaction_ID = " & VarTransactionID, dbFailOnError
      db.Execute "Delete * from TransactionsT where Transaction_ID = " & VarTransactionID, dbFailOnError
     rstTransactionToDelete.MoveNext
   Loop
   
  End If
  
  rstTransactionToDelete.Close
   db.Close
   Set rstTransactionToDelete = Nothing
   Set db = Nothing


ExitProcedure:
If Not rstTransactionToDelete Is Nothing Then
  rstTransactionToDelete.Close
   Set rstTransactionToDelete = Nothing
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
            "Error Source: Delete_Module Delete_Transaction" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select

End Sub

Public Sub Delete_IssuedDocuments(TransactionIDtoDeleteRecords As Long, OldOrNew As Integer, AskConfirmation As Boolean)  ' 1=Only New, 2=Only Old, 3=All
Debug.Print "Exec Priority - " & "Delete Module - " & "Delete_IssuedDocuments " & Time()

'On Error GoTo ErrorHandler
Dim Response As Integer

If AskConfirmation = True Then
Response = MsgBox("You are about to DELETE a DOCUMENT record(s)! Do you want to proceed?", vbExclamation + vbYesNoCancel, "ATTENTION")
If Response <> vbYes Then
Exit Sub
End If
End If

Dim db As DAO.Database
Dim rstIssuedDocumentsToDelete As DAO.Recordset
Dim varIssuedDocumentsID As Long
Dim varNewOrOldRecord As String

Set db = CurrentDb

Select Case OldOrNew
Case 1
   varNewOrOldRecord = " Is_New = True AND "
Case 2
   varNewOrOldRecord = " Is_New = False AND "
Case 3
   varNewOrOldRecord = ""
End Select


 Set rstIssuedDocumentsToDelete = db.OpenRecordset("Select * from IssuedDocumentT where " & varNewOrOldRecord & " Transaction_ID = " & TransactionIDtoDeleteRecords)
   If Not rstIssuedDocumentsToDelete.EOF Then
   rstIssuedDocumentsToDelete.MoveLast
   rstIssuedDocumentsToDelete.MoveFirst
   
   Do Until rstIssuedDocumentsToDelete.EOF
   varIssuedDocumentsID = rstIssuedDocumentsToDelete("Issued_Document_ID")
   
   Call Delete_IssuedDocumentsFinancialDetails(varIssuedDocumentsID, 3, False)
   Call Delete_IssuedDocumentsProductDetails(varIssuedDocumentsID, 3, False)
   Call Delete_LinkAttributeValuesToEntities(varIssuedDocumentsID, 3, False)
   Call Delete_Discounts_Logs(varIssuedDocumentsID, 3, 0)
   db.Execute "Insert INTO IssuedDocumentT_DeletedRecords Select * from IssuedDocumentT where  Is_New = False AND Issued_Document_ID = " & varIssuedDocumentsID, dbFailOnError
   db.Execute "Delete * from IssuedDocumentT where  Issued_Document_ID = " & varIssuedDocumentsID, dbFailOnError
   
   rstIssuedDocumentsToDelete.MoveNext
   Loop
   End If
   
   

ExitProcedure:

If Not rstIssuedDocumentsToDelete Is Nothing Then
  rstIssuedDocumentsToDelete.Close
   Set rstIssuedDocumentsToDelete = Nothing
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
            "Error Source: Delete_Module Delete_IssuedDocuments" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
    End Sub

Public Sub Delete_One_IssuedDocument(IssuedDocumentIDtoDelete As Long, OldOrNew As Integer, AskConfirmation As Boolean)  ' 1=Only New, 2=Only Old, 3=All
Debug.Print "Exec Priority - " & "Delete Module - " & "Delete_One_IssuedDocument " & Time()

'On Error GoTo ErrorHandler

Dim db As DAO.Database
Dim rstIssuedDocumentToDelete As DAO.Recordset
Dim varIssuedDocumentsID As Long
Dim varNewOrOldRecord As String

Set db = CurrentDb

If IsNull(IssuedDocumentIDtoDelete) Or IssuedDocumentIDtoDelete = 0 Then
GoTo ExitProcedure
End If

Dim Response As Integer

If AskConfirmation = True Then
Response = MsgBox("You are about to DELETE a DOCUMENT record(s)! Do you want to proceed?", vbExclamation + vbYesNoCancel, "ATTENTION")
If Response <> vbYes Then
Exit Sub
End If
End If

Select Case OldOrNew
Case 1
   varNewOrOldRecord = " Is_New = True AND "
Case 2
   varNewOrOldRecord = " Is_New = False AND "
Case 3
   varNewOrOldRecord = ""
End Select
      
  Set rstIssuedDocumentToDelete = db.OpenRecordset("Select * from IssuedDocumentT where " & varNewOrOldRecord & " Issued_Document_ID = " & IssuedDocumentIDtoDelete)
   If Not rstIssuedDocumentToDelete.EOF Then
     rstIssuedDocumentToDelete.MoveLast
     rstIssuedDocumentToDelete.MoveFirst
   
     Do Until rstIssuedDocumentToDelete.EOF
      varIssuedDocumentsID = rstIssuedDocumentToDelete("Issued_Document_ID")
    
      Call Delete_IssuedDocumentsFinancialDetails(IssuedDocumentIDtoDelete, 3, False)
      Call Delete_IssuedDocumentsProductDetails(IssuedDocumentIDtoDelete, 3, False)
      Call Delete_LinkAttributeValuesToEntities(IssuedDocumentIDtoDelete, 3, False)
      CurrentDb.Execute "Insert INTO IssuedDocumentT_DeletedRecords Select * from IssuedDocumentT where Is_New = False AND Issued_Document_ID = " & IssuedDocumentIDtoDelete, dbFailOnError
      CurrentDb.Execute "Delete * from IssuedDocumentT where " & varNewOrOldRecord & " Issued_Document_ID = " & IssuedDocumentIDtoDelete, dbFailOnError
 
      rstIssuedDocumentToDelete.MoveNext
     Loop
   End If
   
ExitProcedure:
If Not rstIssuedDocumentToDelete Is Nothing Then
  rstIssuedDocumentToDelete.Close
   Set rstIssuedDocumentToDelete = Nothing
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
            "Error Source: Delete_Module Delete_One_IssuedDocument" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
    End Sub
Public Sub Delete_IssuedDocumentsFinancialDetails(IssuedDocumentIDtoDeleteRecords As Long, OldOrNew As Integer, AskConfirmation As Boolean)  ' 1=Only New, 2=Only Old, 3=All
Debug.Print "Exec Priority - " & "Delete Module - " & "Delete_IssuedDocumentsFinancialDetails " & Time()

'On Error GoTo ErrorHandler
Dim Response As Integer

If AskConfirmation = True Then
Response = MsgBox("You are about to DELETE a DOCUMENT FINANCIAL DETAILS record(s)! Do you want to proceed?", vbExclamation + vbYesNoCancel, "ATTENTION")
If Response <> vbYes Then
Exit Sub
End If
End If

Dim db As DAO.Database
Dim rstIssuedDocumentsFinancialDetailsToDelete As DAO.Recordset
Dim varIssuedDocumentsFinancialDetailsID As Long
Dim varNewOrOldRecord As String

Set db = CurrentDb

Select Case OldOrNew
Case 1
   varNewOrOldRecord = " Is_New = True AND "
Case 2
   varNewOrOldRecord = " Is_New = False AND "
Case 3
   varNewOrOldRecord = ""
End Select

 

Set rstIssuedDocumentsFinancialDetailsToDelete = db.OpenRecordset("Select * from IssuedDocumentFinancialDetailsT where " & varNewOrOldRecord & " Issued_Document_ID = " & IssuedDocumentIDtoDeleteRecords)

          
   If Not rstIssuedDocumentsFinancialDetailsToDelete.EOF Then
   rstIssuedDocumentsFinancialDetailsToDelete.MoveLast
   rstIssuedDocumentsFinancialDetailsToDelete.MoveFirst
   
    Do Until rstIssuedDocumentsFinancialDetailsToDelete.EOF
     varIssuedDocumentsFinancialDetailsID = rstIssuedDocumentsFinancialDetailsToDelete("Issued_Document_Financial_Details_ID")
   
     Call Delete_LinkAttributeValuesToEntities(varIssuedDocumentsFinancialDetailsID, 4, False)

     db.Execute "Insert INTO IssuedDocumentFinancialDetailsT_DeletedRecords Select * from IssuedDocumentFinancialDetailsT where  Is_New = False AND Issued_Document_Financial_Details_ID = " & varIssuedDocumentsFinancialDetailsID, dbFailOnError
     db.Execute "Delete * from IssuedDocumentFinancialDetailsT where Issued_Document_Financial_Details_ID = " & varIssuedDocumentsFinancialDetailsID, dbFailOnError
   
     rstIssuedDocumentsFinancialDetailsToDelete.MoveNext
    Loop
   End If
   
   
ExitProcedure:
If Not rstIssuedDocumentsFinancialDetailsToDelete Is Nothing Then
  rstIssuedDocumentsFinancialDetailsToDelete.Close
   Set rstIssuedDocumentsFinancialDetailsToDelete = Nothing
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
            "Error Source: Delete_Module Delete_IssuedDocumentsFinancialDetails" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
    End Sub
Public Sub Delete_One_IssuedDocumentsFinancialDetails(IssuedDocumentFinancialDetailIDtoDelete As Long, OldOrNew As Integer, AskConfirmation As Boolean)  ' 1=Only New, 2=Only Old, 3=All
Debug.Print "Exec Priority - " & "Delete Module - " & "Delete_One_IssuedDocumentsFinancialDetails " & Time()
'On Error GoTo ErrorHandler

If IsNull(IssuedDocumentFinancialDetailIDtoDelete) Or IssuedDocumentFinancialDetailIDtoDelete = 0 Then
GoTo ExitProcedure
End If

Dim Response As Integer

If AskConfirmation = True Then
Response = MsgBox("You are about to DELETE a DOCUMENT FINANCIAL DETAILS record(s)! Do you want to proceed?", vbExclamation + vbYesNoCancel, "ATTENTION")
If Response <> vbYes Then
Exit Sub
End If
End If

Dim db As DAO.Database
Dim rstIssuedDocumentsFinancialDetailsToDelete As DAO.Recordset
Dim varIssuedDocumentsFinancialDetailsID As Long
Dim varNewOrOldRecord As String

Set db = CurrentDb

Select Case OldOrNew
Case 1
   varNewOrOldRecord = " Is_New = True AND "
Case 2
   varNewOrOldRecord = " Is_New = False AND "
Case 3
   varNewOrOldRecord = ""
End Select
   
 Set rstIssuedDocumentsFinancialDetailsToDelete = db.OpenRecordset("Select * from IssuedDocumentFinancialDetailsT where " & varNewOrOldRecord & " Issued_Document_Financial_Details_ID = " & IssuedDocumentFinancialDetailIDtoDelete)
   If Not rstIssuedDocumentsFinancialDetailsToDelete.EOF Then
     rstIssuedDocumentsFinancialDetailsToDelete.MoveLast
     rstIssuedDocumentsFinancialDetailsToDelete.MoveFirst
   
     Do Until rstIssuedDocumentsFinancialDetailsToDelete.EOF
      varIssuedDocumentsFinancialDetailsID = rstIssuedDocumentsFinancialDetailsToDelete("Issued_Document_Financial_Details_ID")
    
      Call Delete_LinkAttributeValuesToEntities(varIssuedDocumentsFinancialDetailsID, 4, False)
       CurrentDb.Execute "Insert INTO IssuedDocumentFinancialDetailsT_DeletedRecords Select * from IssuedDocumentFinancialDetailsT where Is_New = False AND Issued_Document_Financial_Details_ID = " & varIssuedDocumentsFinancialDetailsID, dbFailOnError
       CurrentDb.Execute "Delete * from IssuedDocumentFinancialDetailsT where Issued_Document_Financial_Details_ID = " & varIssuedDocumentsFinancialDetailsID, dbFailOnError
 
      rstIssuedDocumentsFinancialDetailsToDelete.MoveNext
     Loop
   End If
   
ExitProcedure:
If Not rstIssuedDocumentsFinancialDetailsToDelete Is Nothing Then
  rstIssuedDocumentsFinancialDetailsToDelete.Close
   Set rstIssuedDocumentsFinancialDetailsToDelete = Nothing
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
            "Error Source: Delete_Module Delete_One_IssuedDocumentsFinancialDetails" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
    End Sub
 Public Sub Delete_IssuedDocumentsProductDetails(IssuedDocumentIDtoDeleteRecords As Long, OldOrNew As Integer, AskConfirmation As Boolean) ' 1=Only New, 2=Only Old, 3=All
Debug.Print "Exec Priority - " & "Delete Module - " & "Delete_IssuedDocumentsProductDetails " & Time()

'On Error GoTo ErrorHandler
Dim Response As Integer

If AskConfirmation = True Then
Response = MsgBox("You are about to DELETE a DOCUMENT PRODUCT DETAILS record(s)! Do you want to proceed?", vbExclamation + vbYesNoCancel, "ATTENTION")
If Response <> vbYes Then
Exit Sub
End If
End If

Dim db As DAO.Database
Dim rstIssuedDocumentsProductDetailsToDelete As DAO.Recordset
Dim varIssuedDocumentsProductDetailsID As Long
Dim varNewOrOldRecord As String

Set db = CurrentDb

Select Case OldOrNew
Case 1
   varNewOrOldRecord = " Is_New = True AND "
Case 2
   varNewOrOldRecord = " Is_New = False AND "
Case 3
   varNewOrOldRecord = ""
End Select

 Set rstIssuedDocumentsProductDetailsToDelete = db.OpenRecordset("Select * from IssuedDocumentProductDetailsT where " & varNewOrOldRecord & " Issued_Document_ID = " & IssuedDocumentIDtoDeleteRecords)
   
   If Not rstIssuedDocumentsProductDetailsToDelete.EOF Then
    rstIssuedDocumentsProductDetailsToDelete.MoveLast
    rstIssuedDocumentsProductDetailsToDelete.MoveFirst
   
    Do Until rstIssuedDocumentsProductDetailsToDelete.EOF
     varIssuedDocumentsProductDetailsID = rstIssuedDocumentsProductDetailsToDelete("Issued_Document_Product_Details_ID")
  
     Call Delete_LinkAttributeValuesToEntities(varIssuedDocumentsProductDetailsID, 5, False)
     Call Delete_Discounts_Logs(varIssuedDocumentsProductDetailsID, 3, False)
     db.Execute "Insert INTO IssuedDocumentProductDetailsT_DeletedRecords Select * from IssuedDocumentProductDetailsT where  Is_New = False AND Issued_Document_Product_Details_ID = " & varIssuedDocumentsProductDetailsID, dbFailOnError
     db.Execute "Delete * from IssuedDocumentProductDetailsT where Issued_Document_Product_Details_ID = " & varIssuedDocumentsProductDetailsID, dbFailOnError
   
     rstIssuedDocumentsProductDetailsToDelete.MoveNext
    Loop
   End If
   
   
ExitProcedure:
If Not rstIssuedDocumentsProductDetailsToDelete Is Nothing Then
  rstIssuedDocumentsProductDetailsToDelete.Close
   Set rstIssuedDocumentsProductDetailsToDelete = Nothing
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
            "Error Source: Delete_Module Delete_IssuedDocumentsProductDetails" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
    End Sub
Public Sub Delete_One_IssuedDocumentsProductDetails(IssuedProductDetailIDtoDelete As Long, OldOrNew As Integer, AskConfirmation As Boolean) ' 1=Only New, 2=Only Old, 3=All
Debug.Print "Exec Priority - " & "Delete Module - " & "Delete_One_IssuedDocumentsProductDetails " & Time()
'On Error GoTo ErrorHandler

If IsNull(IssuedProductDetailIDtoDelete) Or IssuedProductDetailIDtoDelete = 0 Then
Exit Sub
End If

Dim Response As Integer

If AskConfirmation = True Then
Response = MsgBox("You are about to DELETE a DOCUMENT PRODUCT DETAILS record(s)! Do you want to proceed?", vbExclamation + vbYesNoCancel, "ATTENTION")
If Response <> vbYes Then
Exit Sub
End If
End If

Dim db As DAO.Database
Dim rstIssuedDocumentProductDetailToDelete As DAO.Recordset
Dim varIssuedDocumentProductDetailID As Long
Dim varNewOrOldRecord As String

Set db = CurrentDb

Select Case OldOrNew
Case 1
   varNewOrOldRecord = " Is_New = True AND "
Case 2
   varNewOrOldRecord = " Is_New = False AND "
Case 3
   varNewOrOldRecord = ""
End Select

 Set rstIssuedDocumentProductDetailToDelete = db.OpenRecordset("Select * from IssuedDocumentProductDetailsT where " & varNewOrOldRecord & " Issued_Document_Product_Details_ID = " & IssuedProductDetailIDtoDelete)
   
   If Not rstIssuedDocumentProductDetailToDelete.EOF Then
    rstIssuedDocumentProductDetailToDelete.MoveLast
    rstIssuedDocumentProductDetailToDelete.MoveFirst
   
    Do Until rstIssuedDocumentProductDetailToDelete.EOF
     varIssuedDocumentProductDetailID = rstIssuedDocumentProductDetailToDelete("Issued_Document_Product_Details_ID")
  
     Call Delete_LinkAttributeValuesToEntities(varIssuedDocumentProductDetailID, 5, False)
     Call Delete_Discounts_Logs(varIssuedDocumentProductDetailID, 3, False)
     db.Execute "Insert INTO IssuedDocumentProductDetailsT_DeletedRecords Select * from IssuedDocumentProductDetailsT where  Is_New = False AND Issued_Document_Product_Details_ID = " & varIssuedDocumentProductDetailID, dbFailOnError
     db.Execute "Delete * from IssuedDocumentProductDetailsT where Issued_Document_Product_Details_ID = " & varIssuedDocumentProductDetailID, dbFailOnError
   
     rstIssuedDocumentProductDetailToDelete.MoveNext
    Loop
   End If
    
ExitProcedure:
If Not rstIssuedDocumentProductDetailToDelete Is Nothing Then
  rstIssuedDocumentProductDetailToDelete.Close
   Set rstIssuedDocumentProductDetailToDelete = Nothing
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
            "Error Source: Delete_Module - Delete_One_IssuedDocumentsProductDetails" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
    End Sub


Public Sub Delete_LinkAttributeValuesToEntities(EntityIDtoDeleteRecords As Long, EntityTypeID As Long, AskConfirmation As Boolean)
Debug.Print "Exec Priority - " & "Delete Module - " & "Delete_LinkAttributeValuesToEntities " & Time()

'On Error GoTo ErrorHandler
Dim Response As Integer

If AskConfirmation = True Then
Response = MsgBox("You are about to DELETE a LINKED ATTRIBUTE VALUE TO ENTITY record(s)! Do you want to proceed?", vbExclamation + vbYesNoCancel, "ATTENTION")
If Response <> vbYes Then
Exit Sub
End If
End If


Dim db As DAO.Database
Dim rstLinkAttributeValuesToEntities As DAO.Recordset
Dim VarLinkAttributeValueToEntitiesID As Long

Set db = CurrentDb

Set rstLinkAttributeValuesToEntities = db.OpenRecordset("Select * from LinkAttributeValueToEntitiesT where Entity_ID = " & EntityIDtoDeleteRecords & " And Entity_Type_ID = " & EntityTypeID)
   If Not rstLinkAttributeValuesToEntities.EOF Then
   rstLinkAttributeValuesToEntities.MoveLast
   rstLinkAttributeValuesToEntities.MoveFirst
      
   Do Until rstLinkAttributeValuesToEntities.EOF
   VarLinkAttributeValueToEntitiesID = rstLinkAttributeValuesToEntities("Link_Attribute_Value_To_Entity_ID")
   Call Delete_LinkAttributeValuesForAttributeValues(VarLinkAttributeValueToEntitiesID, False)
   db.Execute "Insert INTO LinkAttributeValueToEntitiesT_DeletedRecords Select * from LinkAttributeValueToEntitiesT where Entity_Type_ID = " & EntityTypeID & " AND Link_Attribute_Value_To_Entity_ID = " & VarLinkAttributeValueToEntitiesID, dbFailOnError
   db.Execute "Delete * from LinkAttributeValueToEntitiesT where Entity_Type_ID = " & EntityTypeID & " AND Link_Attribute_Value_To_Entity_ID = " & VarLinkAttributeValueToEntitiesID, dbFailOnError
   rstLinkAttributeValuesToEntities.MoveNext
   Loop
   End If

   
ExitProcedure:
If Not rstLinkAttributeValuesToEntities Is Nothing Then
  rstLinkAttributeValuesToEntities.Close
   Set rstLinkAttributeValuesToEntities = Nothing
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
            "Error Source: Delete_Module Delete_LinkAttributeValuesToEntities" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
End Sub
Public Sub Delete_LinkAttributeValuesToEntitiesByAttributeID(AttributeIDtoDeleteRecords As Long, EntityTypeID As Long, AskConfirmation As Boolean)
Debug.Print "Exec Priority - " & "Delete Module - " & "Delete_LinkAttributeValuesToEntitiesByAttributeID " & Time()
'On Error GoTo ErrorHandler

Dim Response As Integer

If AskConfirmation = True Then
Response = MsgBox("You are about to DELETE a LINKED ATTRIBUTE VALUE TO ENTITY record(s)! Do you want to proceed?", vbExclamation + vbYesNoCancel, "ATTENTION")
If Response <> vbYes Then
Exit Sub
End If
End If


Dim db As DAO.Database
Dim rstLinkAttributeValuesToEntities As DAO.Recordset
Dim VarLinkAttributeValueToEntitiesID As Long

Set db = CurrentDb

Set rstLinkAttributeValuesToEntities = db.OpenRecordset("Select * from LinkAttributeValueToEntitiesT where Entity_ID = " & AttributeIDtoDeleteRecords & " And Entity_Type_ID = " & EntityTypeID)
   If Not rstLinkAttributeValuesToEntities.EOF Then
    rstLinkAttributeValuesToEntities.MoveLast
    rstLinkAttributeValuesToEntities.MoveFirst
      
    Do Until rstLinkAttributeValuesToEntities.EOF
     VarLinkAttributeValueToEntitiesID = rstLinkAttributeValuesToEntities("Link_Attribute_Value_To_Entity_ID")
     Call Delete_LinkAttributeValuesForAttributeValues(VarLinkAttributeValueToEntitiesID, False)
     db.Execute "Insert INTO LinkAttributeValueToEntitiesT_DeletedRecords Select * from LinkAttributeValueToEntitiesT where Entity_Type_ID = " & EntityTypeID & " AND Link_Attribute_Value_To_Entity_ID = " & VarLinkAttributeValueToEntitiesID, dbFailOnError
     db.Execute "Delete * from LinkAttributeValueToEntitiesT where Entity_Type_ID = " & EntityTypeID & " AND Link_Attribute_Value_To_Entity_ID = " & VarLinkAttributeValueToEntitiesID, dbFailOnError
     rstLinkAttributeValuesToEntities.MoveNext
    Loop
   End If
   
ExitProcedure:
If Not rstLinkAttributeValuesToEntities Is Nothing Then
  rstLinkAttributeValuesToEntities.Close
   Set rstLinkAttributeValuesToEntities = Nothing
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
            "Error Source: Delete_Module Delete_LinkAttributeValuesToEntitiesByAttributeID" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
End Sub
Public Sub Delete_One_LinkAttributeValuesToEntities(LinkAttributeValueToEntityIDToDelete As Long, EntityTypeID As Integer, AskConfirmation As Boolean) ' the use of EntityTypeID is redundant. I cannot recall why i placed it there, but i leave it.
Debug.Print "Exec Priority - " & "Delete Module - " & "Delete_One_LinkAttributeValuesToEntities " & Time()
'On Error GoTo ErrorHandler

If IsNull(LinkAttributeValueToEntityIDToDelete) Or LinkAttributeValueToEntityIDToDelete = 0 Then
Exit Sub
End If

Dim Response As Integer

If AskConfirmation = True Then
Response = MsgBox("You are about to DELETE a LINKED ATTRIBUTE VALUE TO ENTITY record(s)! Do you want to proceed?", vbExclamation + vbYesNoCancel, "ATTENTION")
If Response <> vbYes Then
Exit Sub
End If
End If

    Call Delete_LinkAttributeValuesForAttributeValues(LinkAttributeValueToEntityIDToDelete, False)
    CurrentDb.Execute "Insert INTO LinkAttributeValueToEntitiesT_DeletedRecords Select * from LinkAttributeValueToEntitiesT where Entity_Type_ID " & IIf(EntityTypeID = 0, " Like ""**"" ", " = " & EntityTypeID) & " AND Link_Attribute_Value_To_Entity_ID = " & LinkAttributeValueToEntityIDToDelete, dbFailOnError
    CurrentDb.Execute "Delete * from LinkAttributeValueToEntitiesT where Entity_Type_ID " & IIf(EntityTypeID = 0, " Like ""**"" ", " = " & EntityTypeID) & " AND Link_Attribute_Value_To_Entity_ID = " & LinkAttributeValueToEntityIDToDelete, dbFailOnError
     
ExitProcedure:
Exit Sub
   
ErrorHandler:
Select Case Err.Number
        
        
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: Delete_Module Delete_One_LinkAttributeValuesToEntities" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
End Sub

Public Sub Delete_LinkAttributeValuesForAttributeValues(EntityIDtoDeleteRecords As Long, AskConfirmation As Boolean)
Debug.Print "Exec Priority - " & "Delete Module - " & "Delete_LinkAttributeValuesForAttributeValues " & Time()
'On Error GoTo ErrorHandler

Dim Response As Integer

If AskConfirmation = True Then
Response = MsgBox("You are about to DELETE a LINKED ATTRIBUTE VALUE FOR ATTRIBUTES record(s)! Do you want to proceed?", vbExclamation + vbYesNoCancel, "ATTENTION")
If Response <> vbYes Then
Exit Sub
End If
End If


 Dim db As DAO.Database
 Dim rstLinkAttributeValuesForAttributeValues As DAO.Recordset
  
 Set db = CurrentDb

 Set rstLinkAttributeValuesForAttributeValues = db.OpenRecordset("Select * from LinkAttributeValueToEntitiesT where Entity_ID = " & EntityIDtoDeleteRecords & " And Entity_Type_ID = 7")
   If Not rstLinkAttributeValuesForAttributeValues.EOF Then
      db.Execute "Insert INTO LinkAttributeValueToEntitiesT_DeletedRecords Select * from LinkAttributeValueToEntitiesT where Entity_Type_ID = 7 AND Entity_ID = " & EntityIDtoDeleteRecords, dbFailOnError
      db.Execute "Delete * from LinkAttributeValueToEntitiesT where Entity_Type_ID = 7 AND Entity_ID = " & EntityIDtoDeleteRecords, dbFailOnError
  End If
  
   
ExitProcedure:
If Not rstLinkAttributeValuesForAttributeValues Is Nothing Then
  rstLinkAttributeValuesForAttributeValues.Close
   Set rstLinkAttributeValuesForAttributeValues = Nothing
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
            "Error Source: Delete_Module Delete_LinkAttributeValuesForAttributeValues" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select

End Sub

Public Sub Delete_Unlinked_Data(myform As Form, AskConfirmation As Boolean)
Debug.Print "Delete Module - " & "Delete_Unlinked_Data " & Time()
'On Error GoTo ErrorHandler

Dim Response As Integer

If AskConfirmation = True Then
Response = MsgBox("You are about to DELETE ALL THE FORM'S RECORDS! Do you want to proceed?", vbExclamation + vbYesNoCancel, "ATTENTION")
If Response <> vbYes Then
Exit Sub
End If
End If

Dim db As DAO.Database
Dim rst As DAO.Recordset

Set db = CurrentDb
Set rst = myform.Recordset
If Not rst.EOF Then
rst.MoveLast
rst.MoveFirst
Do Until rst.EOF
'Debug.Print "rst(0) = " & rst(0)
Select Case myform.Name
    Case "CheckForTransactionsWithoutDocumentsF"
      Call Delete_Transaction(rst(0), 3, False)
    Case "CheckForDocumentsWithoutTransactionF"
      Call Delete_One_IssuedDocument(rst(0), 3, False)
    Case "CheckForAccountingDocumentsWithoutFinancialDetailsF"
      Call Delete_One_IssuedDocument(rst(0), 3, False)
    Case "CheckForCommercialDocumentsWithoutProductDetailsF"
      Call Delete_One_IssuedDocument(rst(0), 3, False)
    Case "CheckForFinancialDetailsWithoutDocumentF"
      Call Delete_One_IssuedDocumentsFinancialDetails(rst(0), 3, False)
    Case "CheckForProductDetailsWithoutDocumentF"
      Call Delete_One_IssuedDocumentsProductDetails(rst(0), 3, False)
    Case "CheckForTransactorsAttributesWithoutTransactorF"
      Call Delete_One_LinkAttributeValuesToEntities(rst(0), 2, False)
    Case "CheckForTransactionsAttributesWithoutTransactionsF"
      Call Delete_One_LinkAttributeValuesToEntities(rst(0), 6, False)
    Case "CheckForProductAttributesWithoutProductsF"
       Call Delete_One_LinkAttributeValuesToEntities(rst(0), 1, False)
    Case "CheckForIssuedDocumentAttributesWithoutIssuedDocumentF"
       Call Delete_One_LinkAttributeValuesToEntities(rst(0), 3, False)
    Case "CheckForIssuedDocFinDetailsAttributesWithoutIssuedDocFinDetailsF"
       Call Delete_One_LinkAttributeValuesToEntities(rst(0), 4, False)
    Case "CheckForIssDocProdDetailsAttributesWithoutIssDocProdDetailsF"
       Call Delete_One_LinkAttributeValuesToEntities(rst(0), 5, False)
    Case "CheckForAttributesSubAttributesWithoutAttributesF"
       Call Delete_One_LinkAttributeValuesToEntities(rst(0), 7, False)
    Case "CheckForActionsWithoutProtocolsF"
       Call Delete_One_Action(rst(0), 3, False)
    Case "CheckForDocumentFinancialDetailsWithoutTransactorsF"
       Call Delete_One_IssuedDocumentsFinancialDetails(rst(0), 3, False)
    Case "CheckForDocumentProductDetailsWithoutProductsF"
       Call Delete_One_IssuedDocumentsProductDetails(rst(0), 3, False)
    Case "CheckForActionAttributesWithoutActionsF"
       Call Delete_One_LinkAttributeValuesToEntities(rst(0), 8, False)
    Case "CheckForProtocolAttributesWithoutProtocolsF"
       Call Delete_One_LinkAttributeValuesToEntities(rst(0), 9, False)
    Case "CheckForLinkAttrValToEntitiesWithoutAttributesF"
       Call Delete_One_LinkAttributeValuesToEntities(rst(0), 0, False)
    Case "CheckCommercialDocumentForMissingOrNonExistingFinTransactorF"
       Call Delete_One_IssuedDocument(rst(0), 3, False)
    Case "CheckForCommercialDocumentsWithoutProductTransactorsF"
       Call Delete_One_IssuedDocument(rst(0), 3, False)
    Case "CheckForDocumentsWithNonExistingOrMissingIssuableDocumentIDF"
       Call Delete_One_IssuedDocument(rst(0), 3, False)
    Case "CheckForDocumentsWithoutDateIssuableDocumentIDIntentionOUUserIDF"
       Call Delete_One_IssuedDocument(rst(0), 3, False)
    Case "CheckForDocumentsWithoutFinancialOrProductDetailsFToDelete"
       Call Delete_One_IssuedDocument(rst(0), 3, False)
    Case "CheckForFinancialDetailsFaulslyLeftMarkedAsNewF"
      Call Delete_One_IssuedDocumentsFinancialDetails(rst(0), 3, False)
    Case "CheckForFinancialDetailsWithZeroOrNullDebitAndCreditF"
      Call Delete_One_IssuedDocumentsFinancialDetails(rst(0), 3, False)
    Case "CheckForIssuedDocumentsFaulslyLeftMarkedAsNewF"
      Call Delete_One_IssuedDocument(rst(0), 3, False)
    Case "CheckForProductDetailsFaulslyLeftMarkedAsNewF"
      Call Delete_One_IssuedDocumentsProductDetails(rst(0), 3, False)
    Case "CheckForProductDetailsWithoutFinancialTransactorsF"
      Call Delete_One_IssuedDocumentsProductDetails(rst(0), 3, False)
    Case "CheckForProductDetailsWithoutPricesF"
      Call Delete_One_IssuedDocumentsProductDetails(rst(0), 3, False)
    Case "CheckForProductDetailsWithoutQuantityF"
      Call Delete_One_IssuedDocumentsProductDetails(rst(0), 3, False)
    Case "CheckForProductDetailsWithoutVatTransactorsF"
      Call Delete_One_IssuedDocumentsProductDetails(rst(0), 3, False)
    Case "CheckForTransactionsFaulslyLeftMarkedAsNewF"
      Call Delete_Transaction(rst(0), 3, False)

      
End Select
rst.MoveNext
Loop
End If


ExitProcedure:
Exit Sub
   
ErrorHandler:
Select Case Err.Number
       
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: Delete_Unlinked_Data " & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
End Sub

Public Sub Delete_Discounts_Logs(DocumentIDtoDeleteDiscountLogRecords As Long, OldOrNew As Integer, AskConfirmation As Boolean)  ' 1=Only New, 2=Only Old, 3=All
Debug.Print "Delete Module - " & "Delete_Discounts_Logs " & Time()
'On Error GoTo ErrorHandler

Dim Response As Integer

If AskConfirmation = True Then
Response = MsgBox("You are about to DELETE DISCOUNT LOG(s)! Do you want to proceed?", vbExclamation + vbYesNoCancel, "ATTENTION")
If Response <> vbYes Then
Exit Sub
End If
End If

Dim db As DAO.Database
Dim rstDiscountLogsToDelete As DAO.Recordset
Dim VarDiscountLogID As Long
Dim varNewOrOldRecord As String

Set db = CurrentDb

Select Case OldOrNew
Case 1
   varNewOrOldRecord = " Is_New = True AND "
Case 2
   varNewOrOldRecord = " Is_New = False AND "
Case 3
   varNewOrOldRecord = ""
End Select


  Set rstDiscountLogsToDelete = db.OpenRecordset("Select * from DiscountLogsT where  " & varNewOrOldRecord & " Issued_Document_ID = " & DocumentIDtoDeleteDiscountLogRecords)
   If Not rstDiscountLogsToDelete.EOF Then
   rstDiscountLogsToDelete.MoveLast
   rstDiscountLogsToDelete.MoveFirst
   
   Do Until rstDiscountLogsToDelete.EOF
   VarDiscountLogID = rstDiscountLogsToDelete("Discount_Logs_ID")
   
   Call Delete_Discounts_Logs_Details(VarDiscountLogID, 3, False)
   db.Execute "Insert INTO DiscountLogsT_DeletedRecords Select * from DiscountLogsT where  Is_New = False AND Discount_Logs_ID = " & VarDiscountLogID, dbFailOnError
   db.Execute "Delete * from DiscountLogsT where Discount_Logs_ID = " & VarDiscountLogID, dbFailOnError
   
   rstDiscountLogsToDelete.MoveNext
   Loop
   End If


ExitProcedure:
If Not rstDiscountLogsToDelete Is Nothing Then
  rstDiscountLogsToDelete.Close
   Set rstDiscountLogsToDelete = Nothing
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
            "Error Source: Delete_Module Delete_Discounts_Logs" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select

End Sub

Public Sub Delete_Discounts_Logs_Details(DiscountLogIDtoDeleteDiscountLogDetailsRecords As Long, OldOrNew As Integer, AskConfirmation As Boolean)  ' 1=Only New, 2=Only Old, 3=All
Debug.Print "Delete Module - " & "Delete_Discounts_Logs_Details " & Time()
'On Error GoTo ErrorHandler
Dim Response As Integer

If AskConfirmation = True Then
Response = MsgBox("You are about to DELETE DELETE DISCOUNT LOG DETAILS(s)!! Do you want to proceed?", vbExclamation + vbYesNoCancel, "ATTENTION")
If Response <> vbYes Then
Exit Sub
End If
End If

Dim db As DAO.Database
Dim rstDiscountLogsDetailsToDelete As DAO.Recordset
Dim VarDiscountLogDetailsID As Long
Dim varNewOrOldRecord As String

Set db = CurrentDb

Select Case OldOrNew
Case 1
   varNewOrOldRecord = " Is_New = True AND "
Case 2
   varNewOrOldRecord = " Is_New = False AND "
Case 3
   varNewOrOldRecord = ""
End Select

 Set rstDiscountLogsDetailsToDelete = db.OpenRecordset("Select * from DiscountLogsDetailsT where " & varNewOrOldRecord & " Discount_Logs_ID = " & DiscountLogIDtoDeleteDiscountLogDetailsRecords)
   
   If Not rstDiscountLogsDetailsToDelete.EOF Then
   rstDiscountLogsDetailsToDelete.MoveLast
   rstDiscountLogsDetailsToDelete.MoveFirst
   
   Do Until rstDiscountLogsDetailsToDelete.EOF
   VarDiscountLogDetailsID = rstDiscountLogsDetailsToDelete("Discounts_Logs_Details_ID")
  
   db.Execute "Insert INTO DiscountLogsDetailsT_DeletedRecords Select * from DiscountLogsDetailsT where  Is_New = False AND Discounts_Logs_Details_ID = " & VarDiscountLogDetailsID, dbFailOnError
   db.Execute "Delete * from DiscountLogsDetailsT where Discounts_Logs_Details_ID = " & VarDiscountLogDetailsID, dbFailOnError
   
   rstDiscountLogsDetailsToDelete.MoveNext
   Loop
   End If
   
   rstDiscountLogsDetailsToDelete.Close
   db.Close
   Set rstDiscountLogsDetailsToDelete = Nothing
   Set db = Nothing
   
ExitProcedure:
Exit Sub
   
ErrorHandler:
Select Case Err.Number
        
        
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: Delete_Module Delete_Discounts_Logs_Details" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select

End Sub

Public Sub Delete_One_Discount_Log(DiscountLogIDtoDelete As Long, OldOrNew As Integer, AskConfirmation As Boolean)  ' 1=Only New, 2=Only Old, 3=All)
Debug.Print "Delete Module - " & "Delete_One_Discount_Log " & Time()

'On Error GoTo ErrorHandler
If IsNull(DiscountLogIDtoDelete) Or DiscountLogIDtoDelete = 0 Then
Exit Sub
End If

Dim Response As Integer

If AskConfirmation = True Then
Response = MsgBox("You are about to DELETE 1 DISCOUNT LOG! Do you want to proceed?", vbExclamation + vbYesNoCancel, "ATTENTION")
If Response <> vbYes Then
Exit Sub
End If
End If

Dim db As DAO.Database
Dim rstDiscountLogsToDelete As DAO.Recordset
Dim VarDiscountLogID As Long
Dim varNewOrOldRecord As String

Set db = CurrentDb

Select Case OldOrNew
Case 1
   varNewOrOldRecord = " Is_New = True AND "
Case 2
   varNewOrOldRecord = " Is_New = False AND "
Case 3
   varNewOrOldRecord = ""
End Select
   
   
 Set rstDiscountLogsToDelete = db.OpenRecordset("Select * from DiscountLogsT where  " & varNewOrOldRecord & " Discount_Logs_ID = " & DiscountLogIDtoDelete)
   If Not rstDiscountLogsToDelete.EOF Then
    rstDiscountLogsToDelete.MoveLast
    rstDiscountLogsToDelete.MoveFirst
   
    Do Until rstDiscountLogsToDelete.EOF
     VarDiscountLogID = rstDiscountLogsToDelete("Discount_Logs_ID")
   
     Call Delete_Discounts_Logs_Details(VarDiscountLogID, 3, False)
     db.Execute "Insert INTO DiscountLogsT_DeletedRecords Select * from DiscountLogsT where  Is_New = False AND Discount_Logs_ID = " & VarDiscountLogID, dbFailOnError
     db.Execute "Delete * from DiscountLogsT where Discount_Logs_ID = " & VarDiscountLogID, dbFailOnError
   
     rstDiscountLogsToDelete.MoveNext
    Loop
   End If
 
ExitProcedure:
If Not rstDiscountLogsToDelete Is Nothing Then
  rstDiscountLogsToDelete.Close
   Set rstDiscountLogsToDelete = Nothing
End If
      
If Not db Is Nothing Then
  db.Close
   Set db = Nothing
End If

Exit Sub
   
ErrorHandler:
Select Case Err.Number
        Case 3022
        Resume Next
        
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: Delete_Module Delete_One_Discount_Log" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
End Sub

Public Sub Delete_One_Discount_Log_Detail(DiscountLogDetailIDtoDelete As Long, OldOrNew As Integer, AskConfirmation As Boolean)  ' 1=Only New, 2=Only Old, 3=All)
Debug.Print "Delete Module - " & "Delete_One_Discount_Log_Detail " & Time()

'On Error GoTo ErrorHandler
If IsNull(DiscountLogDetailIDtoDelete) Or DiscountLogDetailIDtoDelete = 0 Then
Exit Sub
End If

Dim Response As Integer

If AskConfirmation = True Then
Response = MsgBox("You are about to DELETE 1 DISCOUNT LOG DETAIL! Do you want to proceed?", vbExclamation + vbYesNoCancel, "ATTENTION")
If Response <> vbYes Then
Exit Sub
End If
End If

Dim varNewOrOldRecord As String

Select Case OldOrNew
Case 1
   varNewOrOldRecord = " Is_New = True AND "
Case 2
   varNewOrOldRecord = " Is_New = False AND "
Case 3
   varNewOrOldRecord = ""
End Select
   
   CurrentDb.Execute "Insert INTO DiscountLogsDetailsT_DeletedRecords Select * from DiscountLogsDetailsT where Is_New = False AND Discounts_Logs_Details_ID = " & DiscountLogDetailIDtoDelete, dbFailOnError
   CurrentDb.Execute "Delete * from DiscountLogsDetailsT where " & varNewOrOldRecord & " Discounts_Logs_Details_ID = " & DiscountLogDetailIDtoDelete, dbFailOnError
 
ExitProcedure:
Exit Sub
   
ErrorHandler:
Select Case Err.Number
        Case 3022
        Resume Next
        
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: Delete_Module Delete_One_Discount_Log_Detail" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
End Sub

Public Sub Delete_One_Action(ActionIDtoDelete As Long, OldOrNew As Integer, AskConfirmation As Boolean)
Debug.Print "Exec Priority - " & "Delete Module - " & "Delete_One_Action " & Time()
On Error GoTo ErrorHandler

If IsNull(ActionIDtoDelete) Or ActionIDtoDelete = 0 Then
Exit Sub
End If

Dim Response As Integer

If AskConfirmation = True Then
 Response = MsgBox("You are about to DELETE an Action record(s)! Do you want to proceed?", vbExclamation + vbYesNoCancel, "ATTENTION")
 If Response <> vbYes Then
  Exit Sub
 End If
End If



   Call Delete_LinkAttributeValuesToEntities(ActionIDtoDelete, 8, False)
   CurrentDb.Execute "Insert INTO ActionsT_DeletedRecords Select Action_ID,Action_Type_ID,Protocol_ID,Timestamp_Assigned,Timestamp_Planned_To_Be_Executed_From, " & _
   "Timestamp_Planned_To_Be_Executed_Until,Timestamp_Executed,Notes, Action_Insert_Timestamp,Action_Insert_User_ID  from ActionsT Where ActionsT.Action_ID = " & ActionIDtoDelete, dbFailOnError
   CurrentDb.Execute "Delete * from ActionsT where ActionsT.Action_ID = " & ActionIDtoDelete, dbFailOnError
   
     
ExitProcedure:
Exit Sub
   
ErrorHandler:
Select Case Err.Number
        
        
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: Delete_Module Delete_One_Action" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
            
End Sub

Public Sub Delete_One_Protocol(ProtocolIDtoDelete As Long, OldOrNew As Integer, AskConfirmation As Boolean)
Debug.Print "Exec Priority - " & "Delete Module - " & "Delete_One_Protocol " & Time()
'On Error GoTo ErrorHandler

If IsNull(ProtocolIDtoDelete) Or ProtocolIDtoDelete = 0 Then
Exit Sub
End If

Dim Response As Integer

If AskConfirmation = True Then
Response = MsgBox("You are about to DELETE a Protocol record(s)! Do you want to proceed?", vbExclamation + vbYesNoCancel, "ATTENTION")
If Response <> vbYes Then
Exit Sub
End If
End If


   Call Delete_LinkAttributeValuesToEntities(ProtocolIDtoDelete, 9, False)
   Call Delete_Actions(ProtocolIDtoDelete, OldOrNew, AskConfirmation)
   CurrentDb.Execute "Insert INTO ProtocolsT_DeletedRecords Select * from ProtocolsT Where ProtocolsT.Protocol_ID = " & ProtocolIDtoDelete, dbFailOnError
   CurrentDb.Execute "Delete * from ProtocolsT where ProtocolsT.Protocol_ID = " & ProtocolIDtoDelete, dbFailOnError
       
ExitProcedure:
Exit Sub
   
ErrorHandler:
Select Case Err.Number
        
        
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: Delete_Module Delete_One_Protocol" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
            
End Sub

Public Sub Delete_Actions(ProtocolIDtoDeleteRecords As Long, OldOrNew As Integer, AskConfirmation As Boolean)
Debug.Print "Exec Priority - " & "Delete Module - " & "Delete_Actions " & Time()
'On Error GoTo ErrorHandler

Dim Response As Integer

If AskConfirmation = True Then
Response = MsgBox("You are about to DELETE an Action record(s)! Do you want to proceed?", vbExclamation + vbYesNoCancel, "ATTENTION")
If Response <> vbYes Then
Exit Sub
End If
End If

Dim db As DAO.Database
Dim rstActionsToDelete As DAO.Recordset
Dim varActionID As Long


Set db = CurrentDb

Set rstActionsToDelete = db.OpenRecordset("Select * from ActionsT where Protocol_ID = " & ProtocolIDtoDeleteRecords)

          
   If Not rstActionsToDelete.EOF Then
   rstActionsToDelete.MoveLast
   rstActionsToDelete.MoveFirst
   
   Do Until rstActionsToDelete.EOF
    varActionID = rstActionsToDelete("Action_ID")
   
    Call Delete_LinkAttributeValuesToEntities(varActionID, 8, False)

    db.Execute "Insert INTO ActionsT_DeletedRecords Select * from ActionsT where Action_ID = " & varActionID, dbFailOnError
    db.Execute "Delete * from ActionsT where Action_ID = " & varActionID, dbFailOnError
   
    rstActionsToDelete.MoveNext
   Loop
   End If
   
   rstActionsToDelete.Close
   db.Close
   Set rstActionsToDelete = Nothing
   Set db = Nothing
   
ExitProcedure:
Exit Sub
   
ErrorHandler:
Select Case Err.Number
        
        
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: Delete_Module Delete_Actions" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
    End Sub

Public Sub Delete_One_Transactor(TransactorIDToDelete As Long, AskConfirmation As Boolean)
Debug.Print "Module Delere Module - " & "Delete_One_Transactor" & Time()
'On Error GoTo ErrorHandler

Dim Response As Integer

Dim db As DAO.Database
   Set db = CurrentDb
 Dim rstIssuedDocumentsRelatedToTransactor As DAO.Recordset
 Dim rstFinancialDetailsRelatedToTransactor As DAO.Recordset
 Dim rstTransactorToDelete As DAO.Recordset
 
 Set rstIssuedDocumentsRelatedToTransactor = db.OpenRecordset("Select * from IssuedDocumentT where [Transactor_Financial_ID(Other_Entity)] = " & TransactorIDToDelete & _
 " OR [Transactor_Product_ID(Main_Entity)] = " & TransactorIDToDelete & " OR [Transactor_Product_ID(Other_Entity)] = " & TransactorIDToDelete)
 
 Set rstFinancialDetailsRelatedToTransactor = db.OpenRecordset("Select * from IssuedDocumentFinancialDetailsT where Transactor_ID = " & TransactorIDToDelete)
 
 If Not rstIssuedDocumentsRelatedToTransactor.EOF Or Not rstFinancialDetailsRelatedToTransactor.EOF Then
 MsgBox "The Transactor you are trying to delete has been moved. Please delete all the documents in which participates and try again.", vbExclamation + vbOKOnly, "ATTENTION"
 Exit Sub
 End If
 
If AskConfirmation = True Then
 Response = MsgBox("You are about to DELETE a TRANSACTOR(s)! Do you want to proceed?", vbExclamation + vbYesNoCancel, "ATTENTION")
 If Response <> vbYes Then
  Exit Sub
 End If
End If

 Set rstTransactorToDelete = db.OpenRecordset("Select * from TransactorsT where Transactor_ID = " & TransactorIDToDelete)
   If Not rstTransactorToDelete.EOF Then
      Call Delete_LinkAttributeValuesToEntities(TransactorIDToDelete, 2, False)
      db.Execute "Delete * from TransactorsT where Transactor_ID = " & TransactorIDToDelete, dbFailOnError
   End If


ExitProcedure:
If Not rstTransactorToDelete Is Nothing Then
   rstTransactorToDelete.Close
   Set rstTransactorToDelete = Nothing
End If
    
If Not rstIssuedDocumentsRelatedToTransactor Is Nothing Then
   rstIssuedDocumentsRelatedToTransactor.Close
   Set rstIssuedDocumentsRelatedToTransactor = Nothing
End If

If Not rstFinancialDetailsRelatedToTransactor Is Nothing Then
   rstFinancialDetailsRelatedToTransactor.Close
   Set rstFinancialDetailsRelatedToTransactor = Nothing
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
            "Error Source: Delete_One_Transactor" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select

End Sub

Public Sub Soft_Delete_One_Transaction(TransactionIDToDeleteArg As Long, AskConfirmationArg As Boolean)
Debug.Print "Exec Priority - " & "Delete Module - " & "Soft_Delete_One_Transaction " & Time()
'On Error GoTo ErrorHandler

Dim db As DAO.Database
Dim rstTransaction As DAO.Recordset
Dim rstDocument As DAO.Recordset
Dim rstDocumentFinancialDetails As DAO.Recordset
Dim rstDocumentProductDetails As DAO.Recordset
Dim rstDocumentFinancialDetailsBeforeAndAfterEdit As DAO.Recordset
Dim rstForUpdateTranBalanceAndInvBalanceByProductDoc As DAO.Recordset
Dim rstDocumentOtherEntityFinancialTransactor As DAO.Recordset
Dim VarIssuedDocumentID As Long
Dim VarIntentionID As Integer
Dim Response As Integer
Dim DebitAndCreditForTranFromFinancialDetailsRecQDef As QueryDef
Dim DebitAndCreditForTranFromProductDetailsRecQDef As QueryDef

Set db = CurrentDb
'we construct one recordset the transactions (it must be only one)
Set rstTransaction = db.OpenRecordset("select * from TransactionsT where Transaction_ID = " & TransactionIDToDeleteArg)
If rstTransaction.recordcount = 1 Then
   If AskConfirmationArg = True Then
   Response = MsgBox("Do you want to delete this transaction?", vbExclamation + vbYesNo, "Delete Confirmation")
    If Response = vbNo Then
    GoTo ExitProcedure
    End If
   End If

rstTransaction.MoveLast
rstTransaction.MoveFirst
'Now we soft delete the transaction
rstTransaction.Edit
rstTransaction("Is_Deleted") = True
rstTransaction.Update
' we proceed with the document, constructing a recordset with all the documents (could be more than one document per transaction)
Set rstDocument = db.OpenRecordset("select * from IssuedDocumentT where Transaction_ID = " & TransactionIDToDeleteArg)
  If Not rstDocument.EOF Then
   rstDocument.MoveLast
   rstDocument.MoveFirst
   'We iterrate the document recordset (rstDocument)
   Do Until rstDocument.EOF
    VarIssuedDocumentID = rstDocument("Issued_Document_ID")
    VarIntentionID = rstDocument("Intention_ID")
    'we recover the document record to its initial state, just in case it was edited by the user. We want to soft delete the record at its initial state,
    'as user may has done meaningless changes before he decided to soft delete
      Call RecoverIssuedDocument(Nz(rstDocument("Issued_Document_Backup_ID"), 0))
'-----------------------------------------------------------------------------------------------------------------------------------------
       ' we proceed with the documentFinancialDetails, constructing a recordset with all the documentFinancialDetails (of the specific document)
       Set rstDocumentFinancialDetails = db.OpenRecordset("select * from IssuedDocumentFinancialDetailsT where Issued_Document_ID = " & rstDocument("Issued_Document_ID"))
         If Not rstDocumentFinancialDetails.EOF Then
           rstDocumentFinancialDetails.MoveLast
           rstDocumentFinancialDetails.MoveFirst
           Do Until rstDocumentFinancialDetails.EOF
           'we recover the documentFinancialdetails record to its initial state, just in case it was edited by the user. We want to soft delete the record at its initial state,
           'as user may has done meaningless changes before he decided to soft delete
             Call RecoverIssuedDocumentFinancialDetails(Nz(rstDocumentFinancialDetails("IssuedDocumentFinancialDetails_Backup_ID"), 0), False)
             rstDocumentFinancialDetails.MoveNext
           Loop
         End If
         
         'we call the query that makes calculations and produces all the necessary fields in order to update the transactors TotalDebit and TotalCredit field
         'for the soft delete of the record
         Set DebitAndCreditForTranFromFinancialDetailsRecQDef = db.QueryDefs("DebitAndCreditForTransactorsFromFinancialDetailsRecordsQ")
             DebitAndCreditForTranFromFinancialDetailsRecQDef.Parameters("IssuedDocumentIDPar").Value = VarIssuedDocumentID
         Set rstDocumentFinancialDetailsBeforeAndAfterEdit = DebitAndCreditForTranFromFinancialDetailsRecQDef.OpenRecordset()
          Call IterateRecordsets(rstDocumentFinancialDetailsBeforeAndAfterEdit, "Follows rstDocumentFinancialDetailsBeforeAndAfterEdit iteration for UpdateTransactorsTotalDebitAndTotalCreditByDocumentFinancialDetailsRecordset")
         'we call the function which executes the update. 3d argument is "True" as the action is deleting and not inserting new record.
         Call UpdateTransactorsTotalDebitAndTotalCreditByDocumentFinancialDetailsRecordset(rstDocumentFinancialDetailsBeforeAndAfterEdit, VarIntentionID, True)
'---------------------------------------------------------------------------------------------------------------------------------------------
        ' we proceed with the documentProductDetails, constructing a recordset with all the documentProductDetails (of the specific document)
       Set rstDocumentProductDetails = db.OpenRecordset("select * from IssuedDocumentProductDetailsT where Issued_Document_ID = " & rstDocument("Issued_Document_ID"))
         If Not rstDocumentProductDetails.EOF Then
           rstDocumentProductDetails.MoveLast
           rstDocumentProductDetails.MoveFirst
           Do Until rstDocumentProductDetails.EOF
           'we recover the documentProductDetails record to its initial state, just in case it was edited by the user. We want to soft delete the record at its initial state,
           'as user may has done meaningless changes before he decided to soft delete
             Call RecoverIssuedDocumentProductDetails(Nz(rstDocumentProductDetails("IssuedDocumentProductDetails_Backup_ID"), 0), False)
             rstDocumentProductDetails.MoveNext
           Loop
         End If
         
         'we call the query that makes calculations and produces all the necessary fields in order to update the transactors TotalDebit and TotalCredit field
         'for the soft delete of the record
         Set DebitAndCreditForTranFromProductDetailsRecQDef = db.QueryDefs("DebitAndCreditForTransactorsFromProductDetailsRecordsQ")
             DebitAndCreditForTranFromProductDetailsRecQDef.Parameters("IssuedDocumentIDPar").Value = VarIssuedDocumentID
         Set rstForUpdateTranBalanceAndInvBalanceByProductDoc = DebitAndCreditForTranFromProductDetailsRecQDef.OpenRecordset()
         Call IterateRecordsets(rstForUpdateTranBalanceAndInvBalanceByProductDoc, "Follows rstForUpdateTranBalanceAndInvBalanceByProductDoc iteration for UpdateTransactorsBalanceAndInventoryBalanceByProductDocumentDetailsRecordset")
         'we call the function which executes the update. 3d argument is "True" as the action is deleting and not inserting new record.
         Call UpdateTransactorsBalanceAndInventoryBalanceByProductDocumentDetailsRecordset(rstForUpdateTranBalanceAndInvBalanceByProductDoc, VarIntentionID, True)
'----------------------------------------------------------------------------------------------------------------------------------------------
   ' we proceed with the OtherEntityFinancialTransactor, constructing a recordset with the specific transactor (of the specific document) - this routine needs to be executed only in ProductDocuments
       Set rstOtherEntityFinancialTransactor = db.OpenRecordset("select * from IssuedDocumentT where Issued_Document_ID = " & rstDocument("Issued_Document_ID"))
         If Not rstOtherEntityFinancialTransactor.EOF Then
           rstOtherEntityFinancialTransactor.MoveLast
           rstOtherEntityFinancialTransactor.MoveFirst
           Do Until rstOtherEntityFinancialTransactor.EOF
           'we recover the IssuedDocument record to its initial state, just in case it was edited by the user. We want to soft delete the record at its initial state,
           'as user may has done meaningless changes before he decided to soft delete
             Call RecoverIssuedDocument(Nz(rstOtherEntityFinancialTransactor("Issued_Document_Backup_ID"), 0))
             rstOtherEntityFinancialTransactor.MoveNext
           Loop
         End If
         
         'we call the query that makes calculations and produces all the necessary fields in order to update the transactors TotalDebit and TotalCredit field
         'for the soft delete of the record
         Set DebitAndCreditforOtherEntityFinTranNewAndOldForProductDocumentQDef = db.QueryDefs("DebitAndCreditForOtherEntityFinTranForProductDocumentQ")
             DebitAndCreditforOtherEntityFinTranNewAndOldForProductDocumentQDef.Parameters("IssuedDocumentIDPar").Value = VarIssuedDocumentID
         Set rstDocumentOtherEntityFinancialTransactor = DebitAndCreditforOtherEntityFinTranNewAndOldForProductDocumentQDef.OpenRecordset()
         Call IterateRecordsets(rstDocumentOtherEntityFinancialTransactor, "Follows rstDocumentOtherEntityFinancialTransactor iteration for UpdateTransactorsTotalDebitAndTotalCreditByDocumentFinancialDetailsRecordset")
         'we call the function which executes the update. 3d argument is "True" as the action is deleting and not inserting new record.
         Call UpdateTransactorsTotalDebitAndTotalCreditByDocumentFinancialDetailsRecordset(rstDocumentOtherEntityFinancialTransactor, VarIntentionID, True)
'----------------------------------------------------------------------------------------------------------------------------------------------
         
    rstDocument.MoveNext
   Loop
   
  End If
End If

ExitProcedure:

If Not rstTransaction Is Nothing Then
rstTransaction.Close
Set rstTransaction = Nothing
End If

If Not rstDocument Is Nothing Then
rstDocument.Close
Set rstDocument = Nothing
End If

If Not rstDocumentFinancialDetails Is Nothing Then
rstDocumentFinancialDetails.Close
Set rstDocumentFinancialDetails = Nothing
End If

If Not rstDocumentProductDetails Is Nothing Then
rstDocumentProductDetails.Close
Set rstDocumentProductDetails = Nothing
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
            "Error Source: Soft_Delete_One_Transaction" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select


End Sub

Public Sub EmptyDatabase()
Debug.Print "Module Delere Module - " & "EmptyDatabase" & Time()
'On Error GoTo ErrorHandler

Dim db As DAO.Database
Dim rstTransactions As DAO.Recordset

Set db = CurrentDb
Set rstTransactions = db.OpenRecordset("Select * from transactionsT")

If Not rstTransactions.EOF Then
rstTransactions.MoveLast
rstTransactions.MoveFirst

Do Until rstTransactions.EOF
Call Delete_Transaction(rstTransactions(0), 3, False)
rstTransactions.MoveNext
Loop
End If

ExitProcedure:
If Not rstTransactions Is Nothing Then
rstTransactions.Close
Set rstTransactions = Nothing
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
            "Error Source: EmptyDatabase" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
End Sub