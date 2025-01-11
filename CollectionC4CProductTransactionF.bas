Option Compare Database

Public C4CProductTransactionFCollection As New Collection  'Instances of C4CProductTransactionF.

Public Sub OpenAC4CProductTransactionFClient(Optional FormOpenDataModeAddArg As Boolean, Optional DocumentIDArg As Long, Optional TransactionIDArg As Long)
'On Error GoTo ErrorHandler
  Dim VarProcedureTitle As String
  Dim VarModuleName As String
  VarProcedureTitle = "OpenAC4CProductTransactionFClient"
  VarModuleName = "CollectionC4CProductTransactionF"
  Debug.Print VarModuleName & " - " & VarProcedureTitle & " - " & Time()
  '---------------------------------------------------------------------r
Application.Echo False
    'Purpose:   Open an independent instance of form ProductTransactionF.
    
    'we check initially if the document that user asks to open with this form is suitable.
   
    Dim db As DAO.Database
    Dim rstIssuedDocuments As DAO.Recordset
    Dim frm As Form
    Dim VarTransactionID As Integer
    Dim VarIssuedDocumentID As Integer
    Dim VarIssuableDocumentID As Integer
    
    Set db = CurrentDb
'here we check if user wants to add new record or edit old one. If he wants to edit, we must make a series of checks that the document he wants to open is suitable to be opened with this form
If FormOpenDataModeAddArg = False And (DocumentIDArg + TransactionIDArg > 0) Then
    'First we find the VarTransactionID
    If TransactionIDArg <> 0 Then
      VarTransactionID = TransactionIDArg
    Else
      If DocumentIDArg > 0 Then
        VarTransactionID = Nz(DLookup("IssuedDocumentT.Transaction_ID", "IssuedDocumentT", "IssuedDocumentT.Issued_Document_ID = " & DocumentIDArg), 0)
        If VarTransactionID = 0 Then
        MsgBox "The Document number you gave does not correspond to any transactions. Please select other document number!"
        GoTo ExitProcedure
        End If
      End If
    End If
    
     'After, we check if this transaction is a single document transaction, as this form can open single document transactions only.
     'At the same time, we find the VarIssuableDocumentID, which is usefull to find after the IssuableDocumentID and check if it has product detais, as this
     'form opens only documents that have product details
  If FormOpenDataModeAddArg = False And (DocumentIDArg + TransactionIDArg > 0) Then
    Set rstIssuedDocuments = db.OpenRecordset("Select Issued_Document_ID, Issuable_Document_ID from IssuedDocumentT where Transaction_ID = " & VarTransactionID)
    If Not rstIssuedDocuments.EOF And Not rstIssuedDocuments.BOF Then
    rstIssuedDocuments.MoveLast
    rstIssuedDocuments.MoveFirst
    End If
    Select Case rstIssuedDocuments.recordcount
       Case Is < 1
         MsgBox "No document found.", vbOKOnly + vbExclamation
         GoTo ExitProcedure
       Case Is = 1
         VarIssuableDocumentID = rstIssuedDocuments(1)
       Case Is > 1
         MsgBox "This is a multidocument transaction. The form you are trying to open is for single document transactions. Please use the correct form to open the document you are trying to open.", vbOKOnly + vbExclamation
         GoTo ExitProcedure
    End Select
  End If
End If

    TempVars.Add ("TVarOkToCreateNewTransactionForThisFormInstance"), False
    ' Open a new instance, show it, and set a caption.
    Set frm = New Form_C4CProductTransactionF

    frm.Visible = True
    frm.Caption = "Product Transactions Form, opened " & Now() & ", (ID = " & frm.Hwnd & ")"
                   
    'Append it to our collection.
    C4CProductTransactionFCollection.Add Item:=frm, Key:=CStr(frm.Hwnd)
     
    If FormOpenDataModeAddArg = False And (DocumentIDArg + TransactionIDArg > 0) Then
      'Here we check if Issuuable Docuemnt has product details and if yes, we proceed to open the form
      If DLookup("IssuableDocumentT.Has_Product_Details", "IssuableDocumentT", "IssuableDocumentT.Issuable_Document_ID = " & VarIssuableDocumentID) = True Then
         frm.DataEntry = False
         frm.AllowAdditions = False
         frm.AllowDeletions = True
         frm.AllowEdits = True
         
        frm.Requery
      
        Dim rst As DAO.Recordset
        Set rst = frm.RecordsetClone
          rst.FindFirst "[Transaction_ID] = " & VarTransactionID
        If Not frm.RecordsetClone.NoMatch Then
            frm.Bookmark = frm.RecordsetClone.Bookmark
        Else
            MsgBox "Record with ID " & RecordIDArg & " not found.", vbExclamation
            GoTo ExitProcedure
        End If
      Else
        MsgBox "The document you are trying to open, does not have product records. Please use a different suitable form to open this document.", vbOKOnly + vbExclamation
        GoTo ExitProcedure
      End If
    Else
      TempVars!TVarOkToCreateNewTransactionForThisFormInstance = True
      frm.DataEntry = True
      frm.AllowAdditions = True
      frm.AllowDeletions = False
      frm.AllowEdits = False
    End If


    

ExitProcedure:
If Not IsNull(TempVars!TVarOkToCreateNewTransactionForThisFormInstance) Then
 TempVars.Remove "TVarOkToCreateNewTransactionForThisFormInstance"
End If

If Not frm Is Nothing Then
 Set frm = Nothing
End If

If Not rst Is Nothing Then
 rst.Close
 Set rst = Nothing
End If

If Not rstIssuedDocuments Is Nothing Then
 rstIssuedDocuments.Close
 Set rstIssuedDocuments = Nothing
End If

If Not db Is Nothing Then
 db.Close
 Set db = Nothing
End If

Application.Echo True
Exit Sub
   
ErrorHandler:
  Dim VarErrorNum As Long
  Dim VarErrorDescription As String
  VarErrorNum = Err.Number
  VarErrorDescription = Err.Description
  Call ErrorLoging(Err.Number, Err.Description, VarModuleName, VarProcedureTitle)

  Select Case VarErrorNum
    Case Else
        Call ShowErrorMessage(VarErrorNum, VarErrorDescription, VarModuleName, VarProcedureTitle)
    Resume ExitProcedure
  End Select
End Sub

Function CloseAllC4CProductTransactionFClients()
    'Purpose: Close all instances in the clnClient collection.
    'Note: Leaves the copy opened directly from database window.
Debug.Print "Exec Priority - " & "CollectionC4CProductTransactionF - " & "CloseAllC4CProductTransactionFClients" & Time()
On Error GoTo ErrorHandler

    Dim NumberOfMembers As Long
    Dim i As Long
    
    NumberOfMembers = C4CProductTransactionFCollection.Count
    For i = 1 To NumberOfMembers
        C4CProductTransactionFCollection.Remove 1
    Next
    
If CheckIfFormIsOpen("C4CProductTransactionF") Then
DoCmd.Close acForm, "C4CProductTransactionF", acSaveNo
End If

ExitProcedure:
Exit Function
   
ErrorHandler:
Select Case Err.Number
        
        
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: CloseAllC4CProductTransactionFClients " & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select

End Function

Function CloseOneC4CProductTransactionFClient(HwndArg As Long)
Debug.Print "Exec Priority - " & "CollectionC4CProductTransactionF - " & "CloseOneC4CProductTransactionFClient" & Time();
'On Error GoTo Errorhandler

'Purpose: Remove this instance from the C4CProductTransactionFCollection.
    Dim obj As Object           'Object in C4CProductTransactionFCollection
    Dim blnRemove As Boolean    'Flag to remove it.
    Dim frm As Form
    
    For Each i In C4CProductTransactionFCollection
      If i.Hwnd = HwndArg Then
      Set frm = i
      Exit For
      End If
    Next i
    
  'Check if this instance is in the collection.
    '   (It won't be if form was opened directly, or code was reset.)
   ' Debug.Print "frm.Name = " & frm.Name
    For Each obj In C4CProductTransactionFCollection
        If obj.Hwnd = HwndArg Then
            blnRemove = True
            Exit For
        End If
    Next
    
       
    'Deassign the object before removing from collection.
    Set obj = Nothing
    If blnRemove Then
       C4CProductTransactionFCollection.Remove CStr(HwndArg)
           ' If CheckIfFormIsOpen(frm.Name) Then
               ' DoCmd.Close acForm, frm.Name, acSaveNo
           '    'DoCmd.Close acForm, Screen.ActiveForm.Name
           ' End If
  '  Else
   '     DoCmd.Close acForm, "C4CProductTransactionF", acSaveNo
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
            "Error Source: CloseOneC4CProductTransactionFClient" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select

End Function