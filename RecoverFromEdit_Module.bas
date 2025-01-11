Option Compare Database
Public Sub RecoverFullTransaction(TransactionIDArg As Long)
Debug.Print "RecoverFromEditModule - " & "RecoverFullTransaction " & Time()
'On Error GoTo Errorhandler

Dim db As DAO.Database
Dim rstDocuments As DAO.Recordset

Set db = CurrentDb
Set rstDocuments = db.OpenRecordset("select * from issuedDocumentT where Transaction_ID = " & TransactionIDArg)

If Not rstDocuments.EOF Then
 rstDocuments.MoveLast
 rstDocuments.MoveFirst
 Do Until rstDocuments.EOF
   If rstDocuments("Is_New") = True Then
     Call Delete_One_IssuedDocument(rstDocuments("Issued_Document_ID"), 1, False)
   Else
     If Not IsNull(rstDocuments("Issued_Document_Backup_ID")) Then
       Call RecoverIssuedDocument(rstDocuments("Issued_Document_Backup_ID"))
     End If
   End If
 rstDocuments.MoveNext
 Loop
End If


ExitProcedure:
If Not rstDocuments Is Nothing Then
 rstDocuments.Close
 Set rstDocuments = Nothing
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
            "Error Source: RecoverFullTransaction " & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
            
End Sub

Public Sub RecoverTransaction(TransactionBackupIDArg As Long)
Debug.Print "RecoverFromEditModule - " & "RecoverTransaction " & Time()
'On Error GoTo ErrorHandler

Dim db As DAO.Database
Dim rstTransactionDataToRecover As DAO.Recordset
Dim RecoverTransactionFromBackUpTableQ As QueryDef

If IsNull(TransactionBackupIDArg) Then Exit Sub

' we make the recordset (actually a single record recordset) from backup table which contains only the information that we want to recover and update to the main table.
Set db = CurrentDb
Set rstTransactionDataToRecover = db.OpenRecordset("Select Transaction_Type_ID, IS_Deleted FROM TransactionsBackupT WHERE TransactionID_Backup_ID = " & TransactionBackupIDArg)
'Debug.Print "rstTransactionDataToRecover.RecordCount =  " & rstTransactionDataToRecover.RecordCount
If rstTransactionDataToRecover.recordcount = 1 Then

Set RecoverTransactionFromBackUpTableQ = CurrentDb.QueryDefs("TransactionRecoverQ")
'Debug.Print rstTransactionDataToRecover(0) & " - " & rstTransactionDataToRecover(1)
RecoverTransactionFromBackUpTableQ.Parameters(0) = TransactionBackupIDArg
RecoverTransactionFromBackUpTableQ.Parameters(1) = rstTransactionDataToRecover(0)
RecoverTransactionFromBackUpTableQ.Parameters(2) = rstTransactionDataToRecover(1)
RecoverTransactionFromBackUpTableQ.Execute dbFailOnError

rstTransactionDataToRecover.Delete

Else
Debug.Print "PROBLEM! MORE THAN 1 TRANSACTIONS RECORDS FOUND! rstTransactionDataToRecover.RecordCount =  " & rstTransactionDataToRecover.recordcount
Exit Sub
End If

ExitProcedure:

If Not rstTransactionDataToRecover Is Nothing Then
rstTransactionDataToRecover.Close
Set rstTransactionDataToRecover = Nothing
End If

If Not RecoverTransactionFromBackUpTableQ Is Nothing Then
RecoverTransactionFromBackUpTableQ.Close
Set RecoverTransactionFromBackUpTableQ = Nothing
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
            "Error Source: RecoverTransaction " & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
End Sub

Public Sub RecoverIssuedDocument(DocumentBackupIDArg As Long)
Debug.Print "RecoverFromEditModule - " & "RecoverIssuedDocument " & Time()
'On Error GoTo ErrorHandler

Debug.Print "DocumentBackupIDArg = " & DocumentBackupIDArg
Dim db As DAO.Database
Dim rstDocumentDataToRecover As DAO.Recordset
Dim RecoverDocumentFromBackUpTableQ As QueryDef

If IsNull(DocumentBackupIDArg) Then
GoTo ExitProcedure
End If
' we make the recordset from backup table which contains only the information that we want to recover and update to the main table.
Set db = CurrentDb
Set rstDocumentDataToRecover = db.OpenRecordset("Select Transaction_ID, Issued_Date, Issuable_Document_ID, Intention_ID, Document_Official_Number_Sequence, [Financial_Transaction_Point_ID(Main_Entity)], " & _
"[Transactor_Financial_ID(Other_Entity)], [Transactor_Product_ID(Main_Entity)], [Transactor_Product_ID(Other_Entity)], User_Notes_For_Other_Users, User_Notes_For_Accounting_Department, " & _
"Has_Been_Cancelled_By_Document_ID, O, U, Is_Deleted,Issued_Document_ID FROM IssuedDocumentBackupT WHERE IssuedDocumentID_Backup_ID = " & DocumentBackupIDArg)
'Debug.Print "rstTransactionDataToRecover.RecordCount =  " & rstTransactionDataToRecover.RecordCount
If Not rstDocumentDataToRecover.EOF Then
  rstDocumentDataToRecover.MoveLast
  rstDocumentDataToRecover.MoveFirst

   Call RecoverFullLinkAttributeValuesToEntities(3, rstDocumentDataToRecover("Issued_Document_ID"))
   Call RecoverFullIssuedDocumentFinancialDetails(rstDocumentDataToRecover("Issued_Document_ID"), False)
   Call RecoverFullIssuedDocumentProductDetails(rstDocumentDataToRecover("Issued_Document_ID"), False)
   Call RecoverFullDiscountLogsAndDiscountLogsDetails(rstDocumentDataToRecover("Issued_Document_ID"))
   
    Set RecoverDocumentFromBackUpTableQ = CurrentDb.QueryDefs("IssuedDocumentsRecoverQ")
    RecoverDocumentFromBackUpTableQ.Parameters(0) = rstDocumentDataToRecover(0)
    RecoverDocumentFromBackUpTableQ.Parameters(1) = rstDocumentDataToRecover(1)
    RecoverDocumentFromBackUpTableQ.Parameters(2) = rstDocumentDataToRecover(2)
    RecoverDocumentFromBackUpTableQ.Parameters(3) = rstDocumentDataToRecover(3)
    RecoverDocumentFromBackUpTableQ.Parameters(4) = rstDocumentDataToRecover(4)
    RecoverDocumentFromBackUpTableQ.Parameters(5) = rstDocumentDataToRecover(5)
    RecoverDocumentFromBackUpTableQ.Parameters(6) = rstDocumentDataToRecover(6)
    RecoverDocumentFromBackUpTableQ.Parameters(7) = rstDocumentDataToRecover(7)
    RecoverDocumentFromBackUpTableQ.Parameters(8) = rstDocumentDataToRecover(8)
    RecoverDocumentFromBackUpTableQ.Parameters(9) = rstDocumentDataToRecover(9)
    RecoverDocumentFromBackUpTableQ.Parameters(10) = rstDocumentDataToRecover(10)
    RecoverDocumentFromBackUpTableQ.Parameters(11) = rstDocumentDataToRecover(11)
    RecoverDocumentFromBackUpTableQ.Parameters(12) = rstDocumentDataToRecover(12)
    RecoverDocumentFromBackUpTableQ.Parameters(13) = rstDocumentDataToRecover(13)
    RecoverDocumentFromBackUpTableQ.Parameters(14) = rstDocumentDataToRecover(14)
    RecoverDocumentFromBackUpTableQ.Parameters(15) = DocumentBackupIDArg
    RecoverDocumentFromBackUpTableQ.Execute dbFailOnError

    rstDocumentDataToRecover.Delete

Else
Debug.Print "SUB ""RecoverIssuedDocument"" error message: PROBLEM! NO RECORDS FOUND!"
GoTo ExitProcedure
End If

ExitProcedure:

If Not rstDocumentDataToRecover Is Nothing Then
rstDocumentDataToRecover.Close
Set rstDocumentDataToRecover = Nothing
End If

If Not RecoverDocumentFromBackUpTableQ Is Nothing Then
RecoverDocumentFromBackUpTableQ.Close
Set RecoverDocumentFromBackUpTableQ = Nothing
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
            "Error Source: RecoverIssuedDocument " & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
End Sub


Public Sub RecoverIssuedDocumentFinancialDetails(DocumentFinancialDetailsBackupIDArg As Long, KeepRecordsBackedUp As Boolean)  'KeepRecordsBackedUp means that we do not delete the backed up records from table IssuedDocumentFinancialDetailsBackUpT and we do not delete the flag in field IssuedDocumentFinancialDetails_Backup_ID of IssuedDocumentFinancialDetailsT, which shoews that records have been backed up
Debug.Print "RecoverFromEditModule - " & "RecoverIssuedDocumentFinancialDetails " & Time()
'On Error GoTo ErrorHandler

Dim db As DAO.Database
Dim rstDocumentFinancialDetailsDataToRecover As DAO.Recordset
Dim DocumentFinancialDetailsRecoverQ As QueryDef

If IsNull(DocumentFinancialDetailsBackupIDArg) Then Exit Sub

' we make the recordset from backup table which contains only the information that we want to recover and update to the main table.
Set db = CurrentDb
Set rstDocumentFinancialDetailsDataToRecover = db.OpenRecordset("Select Issued_Document_ID, Transactor_ID, Debit, Credit,  Notes, " & _
"Is_Deleted, Issued_Document_Financial_Details_ID  FROM IssuedDocumentFinancialDetailsBackUpT WHERE IssuedDocumentFinancialDetails_Backup_ID = " & DocumentFinancialDetailsBackupIDArg)

If Not rstDocumentFinancialDetailsDataToRecover.EOF Then
  rstDocumentFinancialDetailsDataToRecover.MoveLast
  rstDocumentFinancialDetailsDataToRecover.MoveFirst
  
  Call RecoverFullLinkAttributeValuesToEntities(4, rstDocumentFinancialDetailsDataToRecover("Issued_Document_Financial_Details_ID"))
   
    Set DocumentFinancialDetailsRecoverQ = CurrentDb.QueryDefs("DocumentFinancialDetailsRecoverQ")
   
    DocumentFinancialDetailsRecoverQ.Parameters(0) = rstDocumentFinancialDetailsDataToRecover(0)
    DocumentFinancialDetailsRecoverQ.Parameters(1) = rstDocumentFinancialDetailsDataToRecover(1)
    DocumentFinancialDetailsRecoverQ.Parameters(2) = rstDocumentFinancialDetailsDataToRecover(2)
    DocumentFinancialDetailsRecoverQ.Parameters(3) = rstDocumentFinancialDetailsDataToRecover(3)
    DocumentFinancialDetailsRecoverQ.Parameters(4) = rstDocumentFinancialDetailsDataToRecover(4)
    DocumentFinancialDetailsRecoverQ.Parameters(5) = rstDocumentFinancialDetailsDataToRecover(5)
    DocumentFinancialDetailsRecoverQ.Parameters(6) = DocumentFinancialDetailsBackupIDArg
    If KeepRecordsBackedUp = False Then
      DocumentFinancialDetailsRecoverQ.Parameters(7) = Null
    Else
      DocumentFinancialDetailsRecoverQ.Parameters(7) = DocumentFinancialDetailsBackupIDArg
    End If
    DocumentFinancialDetailsRecoverQ.Execute dbFailOnError

    If KeepRecordsBackedUp = False Then
    rstDocumentFinancialDetailsDataToRecover.Delete
    End If
    
Else
Debug.Print "PROBLEM! NO RECORDS FOUND! Argument brought " & DocumentFinancialDetailsBackupIDArg & " as IssuedDocumentFinancialDetails_Backup_ID, but corresponding recordset is empty!!!"
Exit Sub
End If

ExitProcedure:
If Not DocumentFinancialDetailsRecoverQ Is Nothing Then
DocumentFinancialDetailsRecoverQ.Close
Set DocumentFinancialDetailsRecoverQ = Nothing
End If

If Not rstDocumentFinancialDetailsDataToRecover Is Nothing Then
rstDocumentFinancialDetailsDataToRecover.Close
Set rstDocumentFinancialDetailsDataToRecover = Nothing
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
            "Error Source: RecoverIssuedDocumentFinancialDetails" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select

End Sub

Public Sub RecoverFullIssuedDocumentFinancialDetails(IssuedDocumentIDArg As Long, KeepRecordsBackedUp As Boolean)  'KeepRecordsBackedUp means that we do not delete the backed up records from table IssuedDocumentFinancialDetailsBackUpT and we do not delete the flag in field IssuedDocumentFinancialDetails_Backup_ID of IssuedDocumentFinancialDetailsT, which shoews that records have been backed up
Debug.Print "RecoverFromEditModule - " & "RecoverFullIssuedDocumentFinancialDetails " & Time()
'On Error GoTo ErrorHandler

Dim db As DAO.Database
Dim rstFullDocumentFinancialDetailsRecords As DAO.Recordset
Dim VarDocumentFinancialDetailsID As Long
Dim VarDocumentFinancialDetailsBackUpID As Long

If IsNull(IssuedDocumentIDArg) Then Exit Sub

Set db = CurrentDb

'We make a recordset with all the Financial details records of the document
Set rstFullDocumentFinancialDetailsRecords = db.OpenRecordset("Select * from IssuedDocumentFinancialDetailsT where Issued_Document_ID = " & IssuedDocumentIDArg)

If Not rstFullDocumentFinancialDetailsRecords.EOF Then
  rstFullDocumentFinancialDetailsRecords.MoveLast
  rstFullDocumentFinancialDetailsRecords.MoveFirst

  Do Until rstFullDocumentFinancialDetailsRecords.EOF
  VarDocumentFinancialDetailsID = rstFullDocumentFinancialDetailsRecords("Issued_Document_Financial_Details_ID")

    If rstFullDocumentFinancialDetailsRecords("Is_New") = True Then
           Call Delete_One_IssuedDocumentsFinancialDetails(VarDocumentFinancialDetailsID, 1, False)
        Else
            If Not IsNull(rstFullDocumentFinancialDetailsRecords("IssuedDocumentFinancialDetails_Backup_ID")) Then
            VarDocumentFinancialDetailsBackUpID = rstFullDocumentFinancialDetailsRecords("IssuedDocumentFinancialDetails_Backup_ID")
            Call RecoverIssuedDocumentFinancialDetails(VarDocumentFinancialDetailsBackUpID, False)
            End If
        End If
  
  rstFullDocumentFinancialDetailsRecords.MoveNext
  Loop
  
End If

ExitProcedure:

If Not rstFullDocumentFinancialDetailsRecords Is Nothing Then
rstFullDocumentFinancialDetailsRecords.Close
Set rstFullDocumentFinancialDetailsRecords = Nothing
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
            "Error Source: RecoverFullIssuedDocumentFinancialDetails" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select

End Sub

Public Sub RecoverIssuedDocumentProductDetails(DocumentProductDetailsBackupIDArg As Long, KeepRecordsBackedUp As Boolean)  'KeepRecordsBackedUp means that we do not delete the backed up records from table IssuedDocumentProductDetailsBackUpT and we do not delete the flag in field IssuedDocumentProductDetails_Backup_ID of IssuedDocumentProductDetailsT, which shoews that records have been backed up
Debug.Print "RecoverFromEditModule - " & "RecoverIssuedDocumentProductDetails " & Time()
'On Error GoTo ErrorHandler

Dim db As DAO.Database
Dim rstDocumentProductDetailsDataToRecover As DAO.Recordset
Dim DocumentProductDetailsRecoverQ As QueryDef

If IsNull(DocumentProductDetailsBackupIDArg) Then Exit Sub

' we make the recordset from backup table which contains only the information that we want to recover and update to the main table.
Set db = CurrentDb
Set rstDocumentProductDetailsDataToRecover = db.OpenRecordset("Select Issued_Document_ID, Product_ID, Quantity, Unit_Price_Before_Discount,  [Total # Unit Discount], " & _
"[VAT%], VAT_ID, Notes, [Transactor_Product_ID(Main_Entity)], Is_Deleted, [Transactor_Product_ID(Other_Entity)], Accounting_Behavior_ID, Unit_Price_After_Discount,Financial_Transactor_ID, " & _
"Vat_Transactor_ID, Issued_Document_Product_Details_ID  FROM IssuedDocumentProductDetailsBackUpT WHERE IssuedDocumentProductDetailsID_Backup_ID = " & DocumentProductDetailsBackupIDArg)

If Not rstDocumentProductDetailsDataToRecover.EOF Then
  rstDocumentProductDetailsDataToRecover.MoveLast
  rstDocumentProductDetailsDataToRecover.MoveFirst
   
   Call RecoverFullDiscountLogsAndDiscountLogsDetails(rstDocumentProductDetailsDataToRecover("Issued_Document_ID"))
   Call RecoverFullLinkAttributeValuesToEntities(5, rstDocumentProductDetailsDataToRecover("Issued_Document_Product_Details_ID"))
   
    Set DocumentProductDetailsRecoverQ = CurrentDb.QueryDefs("DocumentProductDetailsRecoverQ")
   
    DocumentProductDetailsRecoverQ.Parameters(0) = rstDocumentProductDetailsDataToRecover(0)
    DocumentProductDetailsRecoverQ.Parameters(1) = rstDocumentProductDetailsDataToRecover(1)
    DocumentProductDetailsRecoverQ.Parameters(2) = rstDocumentProductDetailsDataToRecover(2)
    DocumentProductDetailsRecoverQ.Parameters(3) = rstDocumentProductDetailsDataToRecover(3)
    DocumentProductDetailsRecoverQ.Parameters(4) = rstDocumentProductDetailsDataToRecover(4)
    DocumentProductDetailsRecoverQ.Parameters(5) = rstDocumentProductDetailsDataToRecover(5)
    DocumentProductDetailsRecoverQ.Parameters(6) = rstDocumentProductDetailsDataToRecover(6)
    DocumentProductDetailsRecoverQ.Parameters(7) = rstDocumentProductDetailsDataToRecover(7)
    DocumentProductDetailsRecoverQ.Parameters(8) = rstDocumentProductDetailsDataToRecover(8)
    DocumentProductDetailsRecoverQ.Parameters(9) = rstDocumentProductDetailsDataToRecover(9)
    DocumentProductDetailsRecoverQ.Parameters(10) = DocumentProductDetailsBackupIDArg
      
           If KeepRecordsBackedUp = False Then
             DocumentProductDetailsRecoverQ.Parameters(11) = Null
           Else
             DocumentProductDetailsRecoverQ.Parameters(11) = DocumentProductDetailsBackupIDArg
           End If
      
    DocumentProductDetailsRecoverQ.Parameters(12) = rstDocumentProductDetailsDataToRecover(10)
    DocumentProductDetailsRecoverQ.Parameters(13) = rstDocumentProductDetailsDataToRecover(11)
    DocumentProductDetailsRecoverQ.Parameters(14) = rstDocumentProductDetailsDataToRecover(12)
    DocumentProductDetailsRecoverQ.Parameters(15) = rstDocumentProductDetailsDataToRecover(13)
    DocumentProductDetailsRecoverQ.Parameters(16) = rstDocumentProductDetailsDataToRecover(14)

    DocumentProductDetailsRecoverQ.Execute dbFailOnError
     
    If KeepRecordsBackedUp = False Then
     rstDocumentProductDetailsDataToRecover.Delete
    End If
    
Else
Debug.Print "PROBLEM! NO RECORDS FOUND! Argument brought " & DocumentProductDetailsBackupIDArg & " as IssuedDocumentProductDetails_Backup_ID, but corresponding recordset is empty!!!"
Exit Sub
End If

ExitProcedure:
If Not DocumentProductDetailsRecoverQ Is Nothing Then
DocumentProductDetailsRecoverQ.Close
Set DocumentProductDetailsRecoverQ = Nothing
End If

If Not rstDocumentProductDetailsDataToRecover Is Nothing Then
rstDocumentProductDetailsDataToRecover.Close
Set rstDocumentProductDetailsDataToRecover = Nothing
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
            "Error Source: RecoverIssuedDocumentProductDetails" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select

End Sub
Public Sub RecoverFullIssuedDocumentProductDetails(IssuedDocumentIDArg As Long, KeepRecordsBackedUp As Boolean)  'KeepRecordsBackedUp means that we do not delete the backed up records from table IssuedDocumentProductDetailsBackUpT and we do not delete the flag in field IssuedDocumentProductDetails_Backup_ID of IssuedDocumentProductDetailsT, which shoews that records have been backed up
Debug.Print "RecoverFromEditModule - " & "RecoverFullIssuedDocumentProductDetails " & Time()
'On Error GoTo ErrorHandler

Dim db As DAO.Database
Dim rstFullDocumentProductDetailsRecords As DAO.Recordset
Dim VarDocumentProductDetailsID As Long

Dim VarDocumentProductDetailsBackUpID As Long

If IsNull(IssuedDocumentIDArg) Then Exit Sub

Set db = CurrentDb

'We make a recordset with all the Product details records of the document
Set rstFullDocumentProductDetailsRecords = db.OpenRecordset("Select * from IssuedDocumentProductDetailsT where Issued_Document_ID = " & IssuedDocumentIDArg)

If Not rstFullDocumentProductDetailsRecords.EOF Then
  rstFullDocumentProductDetailsRecords.MoveLast
  rstFullDocumentProductDetailsRecords.MoveFirst

  Do Until rstFullDocumentProductDetailsRecords.EOF
   VarDocumentProductDetailsID = rstFullDocumentProductDetailsRecords("Issued_Document_Product_Details_ID")
   
        If rstFullDocumentProductDetailsRecords("Is_New") = True Then
           Call Delete_One_IssuedDocumentsProductDetails(VarDocumentProductDetailsID, 1, False)
        Else
            If Not IsNull(rstFullDocumentProductDetailsRecords("IssuedDocumentProductDetails_Backup_ID")) Then
            VarDocumentProductDetailsBackUpID = rstFullDocumentProductDetailsRecords("IssuedDocumentProductDetails_Backup_ID")
            Call RecoverIssuedDocumentProductDetails(VarDocumentProductDetailsBackUpID, False)
            End If
        End If

  rstFullDocumentProductDetailsRecords.MoveNext
  Loop
End If

ExitProcedure:
If Not rstFullDocumentProductDetailsRecords Is Nothing Then
rstFullDocumentProductDetailsRecords.Close
Set rstFullDocumentProductDetailsRecords = Nothing
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
            "Error Source: RecoverFullIssuedDocumentProductDetails" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select

End Sub

Public Sub RecoverDiscountLogs(DiscountLogsBackupIDArg As Long)
Debug.Print "RecoverFromEditModule - " & "RecoverDiscountLogs " & Time()
'On Error GoTo ErrorHandler

Dim db As DAO.Database
Dim rstDiscountLogsDataToRecover As DAO.Recordset
Dim DiscountLogsRecoverQ As QueryDef

If IsNull(DiscountLogsBackupIDArg) Then Exit Sub

' we make the recordset from backup table which contains only the information that we want to recover and update to the main table.
Set db = CurrentDb
Set rstDiscountLogsDataToRecover = db.OpenRecordset("Select Issued_Document_ID, Discount_OR_Offer_ID, [%Discount_Percentage],  [#Discount_Value], " & _
"Is_Deleted, Discount_For_Single_ProductDetailsID_Only FROM DiscountLogsBackupT WHERE DiscountLogsID_Backup_ID = " & DiscountLogsBackupIDArg)

If Not rstDiscountLogsDataToRecover.EOF Then
  rstDiscountLogsDataToRecover.MoveLast
  rstDiscountLogsDataToRecover.MoveFirst

    Set DiscountLogsRecoverQ = CurrentDb.QueryDefs("DiscountLogsRecoverQ")
    
    DiscountLogsRecoverQ.Parameters(0) = rstDiscountLogsDataToRecover(0)
    DiscountLogsRecoverQ.Parameters(1) = rstDiscountLogsDataToRecover(1)
    DiscountLogsRecoverQ.Parameters(2) = rstDiscountLogsDataToRecover(2)
    DiscountLogsRecoverQ.Parameters(3) = rstDiscountLogsDataToRecover(3)
    DiscountLogsRecoverQ.Parameters(4) = rstDiscountLogsDataToRecover(4)
    DiscountLogsRecoverQ.Parameters(5) = DiscountLogsBackupIDArg
    DiscountLogsRecoverQ.Parameters(6) = rstDiscountLogsDataToRecover(5)
    DiscountLogsRecoverQ.Execute dbFailOnError
    
    rstDiscountLogsDataToRecover.Delete

Else
Debug.Print "PROBLEM! NO RECORDS FOUND! Argument brought " & DiscountLogsBackupIDArg & " as DiscountLogsID_Backup_ID, but corresponding recordset is empty!!!"
Exit Sub
End If

ExitProcedure:
If Not rstDiscountLogsDataToRecover Is Nothing Then
rstDiscountLogsDataToRecover.Close
Set rstDiscountLogsDataToRecover = Nothing
End If

If Not DiscountLogsRecoverQ Is Nothing Then
DiscountLogsRecoverQ.Close
Set DiscountLogsRecoverQ = Nothing
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
            "Error Source: RecoverDiscountLogs" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select

End Sub

Public Sub RecoverDiscountLogsDetails(DiscountLogsDetailsBackupIDArg As Long)
Debug.Print "RecoverFromEditModule - " & "RecoverDiscountLogsDetails" & Time()
'On Error GoTo ErrorHandler

Dim db As DAO.Database
Dim rstDiscountLogsDetailsDataToRecover As DAO.Recordset
Dim DiscountLogsDetailsRecoverQ As QueryDef

If IsNull(DiscountLogsDetailsBackupIDArg) Then Exit Sub

' we make the recordset from backup table which contains only the information that we want to recover and update to the main table.
Set db = CurrentDb
Set rstDiscountLogsDetailsDataToRecover = db.OpenRecordset("Select Discount_Logs_ID, Product_Details_ID, Unit_Price_Before_This_Discount, Unit_Price_After_This_Discount,  Is_Deleted " & _
"FROM DiscountLogsDetailsBackupT WHERE DiscountsLogsDetailsID_backup_ID = " & DiscountLogsDetailsBackupIDArg)

If Not rstDiscountLogsDetailsDataToRecover.EOF Then
  rstDiscountLogsDetailsDataToRecover.MoveLast
  rstDiscountLogsDetailsDataToRecover.MoveFirst

    Set DiscountLogsDetailsRecoverQ = CurrentDb.QueryDefs("DiscountLogsDetailsRecoverQ")
   
    DiscountLogsDetailsRecoverQ.Parameters(0) = rstDiscountLogsDetailsDataToRecover(0)
    DiscountLogsDetailsRecoverQ.Parameters(1) = rstDiscountLogsDetailsDataToRecover(1)
    DiscountLogsDetailsRecoverQ.Parameters(2) = rstDiscountLogsDetailsDataToRecover(2)
    DiscountLogsDetailsRecoverQ.Parameters(3) = rstDiscountLogsDetailsDataToRecover(3)
    DiscountLogsDetailsRecoverQ.Parameters(4) = rstDiscountLogsDetailsDataToRecover(4)
    DiscountLogsDetailsRecoverQ.Parameters(5) = DiscountLogsDetailsBackupIDArg
    DiscountLogsDetailsRecoverQ.Execute dbFailOnError

    rstDiscountLogsDetailsDataToRecover.Delete

Else
Debug.Print "PROBLEM! NO RECORDS FOUND! Argument brought " & DiscountLogsDetailsBackupIDArg & " as DiscountLogsID_Backup_ID, but corresponding recordset is empty!!!"
Exit Sub
End If

ExitProcedure:
If Not DiscountLogsDetailsRecoverQ Is Nothing Then
DiscountLogsDetailsRecoverQ.Close
Set DiscountLogsDetailsRecoverQ = Nothing
End If

If Not rstDiscountLogsDetailsDataToRecover Is Nothing Then
rstDiscountLogsDetailsDataToRecover.Close
Set rstDiscountLogsDetailsDataToRecover = Nothing
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
            "Error Source: RecoverDiscountLogsDetails" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select

End Sub


Public Sub RecoverFullDiscountLogsAndDiscountLogsDetails(IssuedDocumentIDArg)
Debug.Print "RecoverFromEditModule - " & "RecoverFullDiscountLogsAndDiscountLogsDetails " & Time()
'On Error GoTo ErrorHandler
   
Dim db As DAO.Database
Dim rstDiscountLogs As DAO.Recordset
Dim rstDiscoultLogsDetails As DAO.Recordset
Dim VarDiscountLogID As Long
Dim VarDiscountLogDetailsID As Long
Dim VarDiscountLogBackUpID As Long
Dim VarDiscountLogDetailsBackUpID As Long

Set db = CurrentDb
Set rstDiscountLogs = db.OpenRecordset("SELECT DiscountLogsT.* " & _
"FROM DiscountLogsT " & _
"WHERE Issued_Document_ID = " & IssuedDocumentIDArg)
      
       'We iterate the recordset of the DiscountlogsF and for every record we check if Is_New IsNull Or Not. If it is we delete the record. If it is not,_
   'we check if the DiscountLogs_Backup_ID IsNull Or Not. If it is not we then call the RecoverDiscountLogs for the specific DiscountLogID.
   
    If rstDiscountLogs.recordcount > 0 Then
      rstDiscountLogs.MoveLast
      rstDiscountLogs.MoveFirst
      Do Until rstDiscountLogs.EOF
      'Debug.Print rstDiscountLogs.recordCount
      VarDiscountLogID = rstDiscountLogs("Discount_Logs_ID")
        If rstDiscountLogs("Is_New") = True Then
           Call Delete_One_Discount_Log(VarDiscountLogID, 1, 0)
        Else
           If Not IsNull(rstDiscountLogs("DiscountLogs_Backup_ID")) Then
           VarDiscountLogBackUpID = rstDiscountLogs("DiscountLogs_Backup_ID")
           Call RecoverDiscountLogs(VarDiscountLogBackUpID)
           End If
         
         'We construct a recordset of the DiscountLogDetails that are childs of the rstDiscountLogs("Discount_Logs_ID")
           Set rstDiscoultLogsDetails = db.OpenRecordset("SELECT DiscountLogsDetailsT.* " & _
                "FROM DiscountLogsDetailsT " & _
                "WHERE  Discount_Logs_ID = " & VarDiscountLogID)
        
        'We iterate the recordset of the DiscountlogsDetails and for every record we check if Is_New IsNull Or Not. If it is we delete the record. If it is not,_
       'we check if the DiscountLogsDetails_Backup_ID IsNull Or Not. If it is not we then call the RecoverDiscountLogsDetails for the specific DiscountLogDetailsID.
   
           If Not rstDiscoultLogsDetails.EOF Then
             rstDiscoultLogsDetails.MoveLast
             rstDiscoultLogsDetails.MoveFirst
             Do Until rstDiscoultLogsDetails.EOF
              VarDiscountLogDetailsID = rstDiscoultLogsDetails("Discounts_Logs_Details_ID")
              If rstDiscoultLogsDetails("Is_New") = True Then
                Call Delete_One_Discount_Log_Detail(VarDiscountLogDetailsID, 1, 0)
              Else
                If Not IsNull(rstDiscoultLogsDetails("DiscountLogsDetails_Backup_ID")) Then
                  VarDiscountLogDetailsBackUpID = rstDiscoultLogsDetails("DiscountLogsDetails_Backup_ID")
                  Call RecoverDiscountLogsDetails(VarDiscountLogDetailsBackUpID)
                End If
             End If
            rstDiscoultLogsDetails.MoveNext
            Loop
          End If
       End If
     rstDiscountLogs.MoveNext
     Loop
   End If
         
     
ExitProcedure:
Exit Sub
   
If Not rstDiscountLogs Is Nothing Then
rstDiscountLogs.Close
Set rstDiscountLogs = Nothing
End If

If Not rstDiscoultLogsDetails Is Nothing Then
rstDiscoultLogsDetails.Close
Set rstDiscoultLogsDetails = Nothing
End If

If Not db Is Nothing Then
db.Close
Set db = Nothing
End If

ErrorHandler:
Select Case Err.Number
        
        
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: RecoverFullDiscountLogsAndDiscountLogsDetails" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select


End Sub

Public Sub RecoverFullLinkAttributeValuesToEntities(EntityTypeIDArg As Integer, EntityIDArg As Long)
Debug.Print "RecoverFromEditModule - " & "RecoverFullLinkAttributeValuesToEntities " & Time()
'On Error GoTo Errorhandler
   
Dim db As DAO.Database
Dim rstLinkAttributeValuesToEntities As DAO.Recordset
Dim rstLinkSubAttributeValuesToEntities As DAO.Recordset
Dim VarLinkAttributeValueToEntityID As Long
Dim VarLinkSubAttributeValueToEntityID As Long
Dim VarLinkAttributeValueToEntityIDBackupID As Long
Dim VarLinkSubAttributeValueToEntityIDBackupID As Long

Set db = CurrentDb
Set rstLinkAttributeValuesToEntities = db.OpenRecordset("SELECT * " & _
"FROM LinkAttributeValueToEntitiesT " & _
"WHERE Entity_Type_ID = " & EntityTypeIDArg & " AND Entity_ID = " & EntityIDArg)
      
       'We iterate the recordset rstLinkAttributeValuesToEntities and for every record we check if Is_New IsNull Or Not. If it is we delete the record. If it is not,_
   'we check if the LinkAttributeValueToEntityID_Backup_ID IsNull Or Not. If it is not we then call the RecoverLinkAttributeValueToEntities for the specific Link_Attribute_Value_To_Entity_ID.
   
    If rstLinkAttributeValuesToEntities.recordcount > 0 Then
      rstLinkAttributeValuesToEntities.MoveLast
      rstLinkAttributeValuesToEntities.MoveFirst
      Do Until rstLinkAttributeValuesToEntities.EOF
        VarLinkAttributeValueToEntityID = rstLinkAttributeValuesToEntities("Link_Attribute_Value_To_Entity_ID")
        If rstLinkAttributeValuesToEntities("Is_New") = True Then
           Call Delete_One_LinkAttributeValuesToEntities(VarLinkAttributeValueToEntityID, EntityTypeIDArg, False)
        Else
            If Not IsNull(rstLinkAttributeValuesToEntities("LinkAttributeValueToEntityID_Backup_ID")) Then
            VarLinkAttributeValueToEntityIDBackupID = rstLinkAttributeValuesToEntities("LinkAttributeValueToEntityID_Backup_ID")
            Call RecoverLinkAttributeValueToEntities(VarLinkAttributeValueToEntityIDBackupID)
            End If
        End If
       rstLinkAttributeValuesToEntities.MoveNext
       Loop
  
    End If
         
     
ExitProcedure:
Exit Sub
   
If Not rstLinkAttributeValuesToEntities Is Nothing Then
rstLinkAttributeValuesToEntities.Close
Set rstLinkAttributeValuesToEntities = Nothing
End If

If Not rstLinkSubAttributeValuesToEntities Is Nothing Then
rstLinkSubAttributeValuesToEntities.Close
Set rstLinkSubAttributeValuesToEntities = Nothing
End If

If Not db Is Nothing Then
db.Close
Set db = Nothing
End If

ErrorHandler:
Select Case Err.Number
        
        
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: RecoverFullLinkAttributeValuesToEntities" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select

End Sub

Public Sub RecoverLinkAttributeValueToEntities(LinkAttributeValueToEntityBackupIDArg As Long) ' it recovers one record for LinkAttributeValueToEntities and all records for its LinkSubAttributeValueToEntity
Debug.Print "RecoverFromEditModule - " & "RecoverLinkAttributeValueToEntities " & Time()
'On Error GoTo Errorhandler

Dim db As DAO.Database
Dim rstLinkAttributeValueToEntitiesToRecover As DAO.Recordset
Dim LinkAttributeValueToEntityRecoverQdef As QueryDef
Dim rstSubAttributesValuesToEntityToRecover As DAO.Recordset

If IsNull(LinkAttributeValueToEntityBackupIDArg) Then Exit Sub

' we make the recordset from backup table which contains only the information that we want to recover and update to the main table.
Set db = CurrentDb
Set rstLinkAttributeValueToEntitiesToRecover = db.OpenRecordset("Select * FROM LinkAttributeValueToEntitiesBackupT WHERE LinkAttributeValueToEntityID_Backup_ID = " & LinkAttributeValueToEntityBackupIDArg)

If Not rstLinkAttributeValueToEntitiesToRecover.EOF Then
  rstLinkAttributeValueToEntitiesToRecover.MoveLast
  rstLinkAttributeValueToEntitiesToRecover.MoveFirst

    Set LinkAttributeValueToEntityRecoverQdef = CurrentDb.QueryDefs("LinkAttributeValueToEntitiesRecoveQ")
    
    LinkAttributeValueToEntityRecoverQdef.Parameters(0) = rstLinkAttributeValueToEntitiesToRecover("LinkAttributeValueToEntityID_Backup_ID")
   ' Debug.Print "LinkAttributeValueToEntityID_Backup_ID = " & LinkAttributeValueToEntityRecoverQdef.Parameters(0)
    LinkAttributeValueToEntityRecoverQdef.Parameters(1) = rstLinkAttributeValueToEntitiesToRecover("Entity_Type_ID")
    LinkAttributeValueToEntityRecoverQdef.Parameters(2) = rstLinkAttributeValueToEntitiesToRecover("Entity_ID")
    LinkAttributeValueToEntityRecoverQdef.Parameters(3) = rstLinkAttributeValueToEntitiesToRecover("Attribute_ID")
    LinkAttributeValueToEntityRecoverQdef.Parameters(4) = rstLinkAttributeValueToEntitiesToRecover("Attribute_Value_String")
   ' Debug.Print "Attribute_Value_String = " & LinkAttributeValueToEntityRecoverQdef.Parameters(4)
    LinkAttributeValueToEntityRecoverQdef.Parameters(5) = rstLinkAttributeValueToEntitiesToRecover("Attribute_Value_Number")
   ' Debug.Print "Attribute_Value_Number = " & LinkAttributeValueToEntityRecoverQdef.Parameters(5)
    LinkAttributeValueToEntityRecoverQdef.Parameters(6) = rstLinkAttributeValueToEntitiesToRecover("Attribute_Value_Boolean")
    LinkAttributeValueToEntityRecoverQdef.Parameters(7) = rstLinkAttributeValueToEntitiesToRecover("Attribute_Value_Date")
    LinkAttributeValueToEntityRecoverQdef.Parameters(8) = rstLinkAttributeValueToEntitiesToRecover("Attribute_Value_Time")
    LinkAttributeValueToEntityRecoverQdef.Parameters(9) = rstLinkAttributeValueToEntitiesToRecover("Attribute_Value_TImestamp")
    LinkAttributeValueToEntityRecoverQdef.Parameters(10) = rstLinkAttributeValueToEntitiesToRecover("Notes")
    LinkAttributeValueToEntityRecoverQdef.Parameters(11) = rstLinkAttributeValueToEntitiesToRecover("Is_Included_To_Suggested_Attributes")
    LinkAttributeValueToEntityRecoverQdef.Parameters(12) = rstLinkAttributeValueToEntitiesToRecover("Is_Deleted")
    LinkAttributeValueToEntityRecoverQdef.Execute dbFailOnError
    
   
    
    Set rstSubAttributesValuesToEntityToRecover = db.OpenRecordset("Select * FROM LinkAttributeValueToEntitiesT WHERE Entity_Type_ID = 7 AND Entity_ID = " & rstLinkAttributeValueToEntitiesToRecover("Link_Attribute_Value_To_Entity_ID") & " AND LinkAttributeValueToEntityID_Backup_ID is not null")
    If Not rstSubAttributesValuesToEntityToRecover.EOF Then
        rstLinkAttributeValueToEntitiesToRecover.MoveLast
        rstLinkAttributeValueToEntitiesToRecover.MoveFirst
        Do Until rstSubAttributesValuesToEntityToRecover.EOF
        Call RecoverLinkSubAttributeValueToEntities(rstSubAttributesValuesToEntityToRecover("LinkAttributeValueToEntityID_Backup_ID"))
        rstSubAttributesValuesToEntityToRecover.MoveNext
        Loop
    End If
    
     rstLinkAttributeValueToEntitiesToRecover.Delete
Else
Debug.Print "PROBLEM! NO RECORDS FOUND! In sub RecoverLinkAttributeValueToEntities, Argument brought " & LinkAttributeValueToEntityBackupIDArg & " as LinkAttributeValueToEntityID_Backup_ID, but corresponding recordset is empty!!! Please show this message to IT"

Exit Sub
End If

ExitProcedure:
If Not rstLinkAttributeValueToEntitiesToRecover Is Nothing Then
rstLinkAttributeValueToEntitiesToRecover.Close
Set rstLinkAttributeValueToEntitiesToRecover = Nothing
End If

If Not LinkAttributeValueToEntityRecoverQdef Is Nothing Then
LinkAttributeValueToEntityRecoverQdef.Close
Set LinkAttributeValueToEntityRecoverQdef = Nothing
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
            "Error Source: RecoverLinkAttributeValueToEntities" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select

End Sub

Public Sub RecoverLinkSubAttributeValueToEntities(LinkSubAttributeValueToEntityBackupIDArg As Long) ' it recovers one record LinkSubAttributeValueToEntity
Debug.Print "RecoverFromEditModule - " & "RecoverLinkSubAttributeValueToEntities " & Time()
'On Error GoTo Errorhandler

Dim db As DAO.Database
Dim rstLinkSubAttributeValueToEntitiesToRecover As DAO.Recordset
Dim LinkSubAttributeValueToEntityRecoverQdef As QueryDef

If IsNull(LinkSubAttributeValueToEntityBackupIDArg) Then Exit Sub

' we make the recordset from backup table which contains only the information that we want to recover and update to the main table.
Set db = CurrentDb
Set rstLinkSubAttributeValueToEntitiesToRecover = db.OpenRecordset("Select * FROM LinkAttributeValueToEntitiesBackupT WHERE LinkAttributeValueToEntityID_Backup_ID = " & LinkSubAttributeValueToEntityBackupIDArg)

If Not rstLinkSubAttributeValueToEntitiesToRecover.EOF Then
  rstLinkSubAttributeValueToEntitiesToRecover.MoveLast
  rstLinkSubAttributeValueToEntitiesToRecover.MoveFirst

    Set LinkSubAttributeValueToEntityRecoverQdef = CurrentDb.QueryDefs("LinkAttributeValueToEntitiesRecoveQ")
    
    LinkSubAttributeValueToEntityRecoverQdef.Parameters(0) = rstLinkSubAttributeValueToEntitiesToRecover("LinkAttributeValueToEntityID_Backup_ID")
    'Debug.Print "LinkAttributeValueToEntityID_Backup_ID = " & LinkSubAttributeValueToEntityRecoverQdef.Parameters(0)
    LinkSubAttributeValueToEntityRecoverQdef.Parameters(1) = rstLinkSubAttributeValueToEntitiesToRecover("Entity_Type_ID")
    LinkSubAttributeValueToEntityRecoverQdef.Parameters(2) = rstLinkSubAttributeValueToEntitiesToRecover("Entity_ID")
    LinkSubAttributeValueToEntityRecoverQdef.Parameters(3) = rstLinkSubAttributeValueToEntitiesToRecover("Attribute_ID")
    LinkSubAttributeValueToEntityRecoverQdef.Parameters(4) = rstLinkSubAttributeValueToEntitiesToRecover("Attribute_Value_String")
    'Debug.Print "Attribute_Value_String = " & LinkSubAttributeValueToEntityRecoverQdef.Parameters(4)
    LinkSubAttributeValueToEntityRecoverQdef.Parameters(5) = rstLinkSubAttributeValueToEntitiesToRecover("Attribute_Value_Number")
   ' Debug.Print "Attribute_Value_Number = " & LinkSubAttributeValueToEntityRecoverQdef.Parameters(5)
    LinkSubAttributeValueToEntityRecoverQdef.Parameters(6) = rstLinkSubAttributeValueToEntitiesToRecover("Attribute_Value_Boolean")
    LinkSubAttributeValueToEntityRecoverQdef.Parameters(7) = rstLinkSubAttributeValueToEntitiesToRecover("Attribute_Value_Date")
    LinkSubAttributeValueToEntityRecoverQdef.Parameters(8) = rstLinkSubAttributeValueToEntitiesToRecover("Attribute_Value_Time")
    LinkSubAttributeValueToEntityRecoverQdef.Parameters(9) = rstLinkSubAttributeValueToEntitiesToRecover("Attribute_Value_TImestamp")
    LinkSubAttributeValueToEntityRecoverQdef.Parameters(10) = rstLinkSubAttributeValueToEntitiesToRecover("Notes")
    LinkSubAttributeValueToEntityRecoverQdef.Parameters(11) = rstLinkSubAttributeValueToEntitiesToRecover("Is_Included_To_Suggested_Attributes")
    LinkSubAttributeValueToEntityRecoverQdef.Parameters(12) = rstLinkSubAttributeValueToEntitiesToRecover("Is_Deleted")

    LinkSubAttributeValueToEntityRecoverQdef.Execute dbFailOnError
    
    rstLinkSubAttributeValueToEntitiesToRecover.Delete

Else
Debug.Print "PROBLEM! NO RECORDS FOUND! In sub RecoverLinkSubAttributeValueToEntities, Argument brought " & LinkSubAttributeValueToEntityBackupIDArg & " as LinkAttributeValueToEntityID_Backup_ID, but corresponding recordset is empty!!! Please show this message to IT"

Exit Sub
End If

ExitProcedure:
If Not rstLinkSubAttributeValueToEntitiesToRecover Is Nothing Then
rstLinkSubAttributeValueToEntitiesToRecover.Close
Set rstLinkSubAttributeValueToEntitiesToRecover = Nothing
End If

If Not LinkSubAttributeValueToEntityRecoverQdef Is Nothing Then
LinkSubAttributeValueToEntityRecoverQdef.Close
Set LinkSubAttributeValueToEntityRecoverQdef = Nothing
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
            "Error Source: RecoverLinkSubAttributeValueToEntities" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select

End Sub