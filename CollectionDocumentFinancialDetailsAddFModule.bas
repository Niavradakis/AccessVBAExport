Option Compare Database

Public DocumentFinancialAddFCollection As New Collection  'Instances of DocumentFinancialDetailsAddF.

Public Function OpenADocumentFinancialAddFClient(Optional FormOpenDataModeAddArg As Boolean, Optional DocumentIDArg As Long, Optional TransactionIDArg As Long)
Debug.Print "Exec Priority - " & "CollectionDocumentFinancialDetailsAddFModule - " & "OpenADocumentFinancialAddFClient" & Time()
 On Error GoTo ErrorHandler
    
      'Purpose:   Open an independent instance of form ProductTransactionF.
    Dim frm As Form
    Dim VarTransactionID
    
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
  
  TempVars.Add ("TempVars!TVarOkToCreateNewTransactionForThisFormInstance"), False
    ' Open a new instance, show it, and set a caption.
    Set frm = New Form_DocumentFinancialDetailsAddF
    
    frm.Visible = True
    frm.Caption = "Financial Transactions Form, opened " & Now() & ", (ID = " & frm.Hwnd & ")"
                   
    'Append it to our collection.
    DocumentFinancialAddFCollection.Add Item:=frm, Key:=CStr(frm.Hwnd)
     
     If FormOpenDataModeAddArg = False And (DocumentIDArg + TransactionIDArg > 0) Then
     
    
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
     TempVars!TVarOkToCreateNewTransactionForThisFormInstance = True
         frm.DataEntry = True
         frm.AllowAdditions = True
         frm.AllowDeletions = False
         frm.AllowEdits = False
     End If
    
    
ExitProcedure:
If Not IsNull(TempVars!TVarOkToCreateNewTransactionForThisFormInstance) Then
TempVars.Remove "TempVars!TVarOkToCreateNewTransactionForThisFormInstance"
End If

If Not frm Is Nothing Then
Set frm = Nothing
End If

If Not rst Is Nothing Then
rst.Close
Set rst = Nothing
End If

Exit Function
   
ErrorHandler:
Select Case Err.Number
        
        
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: OpenADocumentFinancialAddFClient" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
End Function

Function CloseAllDocumentFinancialAddFClients()
    'Purpose: Close all instances in the clnClient collection.
    'Note: Leaves the copy opened directly from database window.
Debug.Print "Exec Priority - " & "CollectionDocumentFinancialDetailsAddFModule - " & "CloseAllDocumentFinancialAddFClients" & Time()
    Dim NumberOfMembers As Long
    Dim i As Long
    
    NumberOfMembers = DocumentFinancialAddFCollection.Count
    For i = 1 To NumberOfMembers
        DocumentFinancialAddFCollection.Remove 1
    Next
    
If CheckIfFormIsOpen("DocumentFinancialDetailsAddF") Then
DoCmd.Close acForm, "DocumentFinancialDetailsAddF", acSaveNo
End If

ExitProcedure:
Exit Function
   
ErrorHandler:
Select Case Err.Number
        
        
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: CloseAllDocumentFinancialAddFClients" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select

End Function

Function CloseOneDocumentFinancialAddFClients()
   
Debug.Print "Exec Priority - " & "CollectionDocumentFinancialDetailsAddFModule - " & "CloseOneDocumentFinancialAddFClients " & Time()
On Error GoTo ErrorHandler

'Purpose: Remove this instance from the clnClient collection.
    Dim obj As Object           'Object in clnClient
    Dim blnRemove As Boolean    'Flag to remove it.
    Dim frm As Form
    
    For Each i In Forms
      If i.Hwnd = HwndArg Then
      Set frm = i
      Exit For
      End If
    Next i
    
  'Check if this instance is in the collection.
    '   (It won't be if form was opened directly, or code was reset.)
    For Each obj In DocumentFinancialAddFCollection
        If obj.Hwnd = HwndArg Then
            blnRemove = True
            Exit For
        End If
    Next
    
       
    'Deassign the object before removing from collection.
    Set obj = Nothing
    If blnRemove Then
       DocumentFinancialAddFCollection.Remove CStr(HwndArg)
        If CheckIfFormIsOpen(frm.Name) Then
            DoCmd.Close acForm, frm.Name, acSaveNo
        End If
    Else
        DoCmd.Close acForm, "DocumentFinancialDetailsAddF", acSaveNo
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
            "Error Source: CloseOneDocumentFinancialAddFClients" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select

End Function