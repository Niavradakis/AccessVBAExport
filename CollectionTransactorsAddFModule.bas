Option Compare Database
Option Explicit

Public TransactorsAddFCollection As New Collection  'Instances of TransactorsAddF.

Public Sub OpenATransactorsAddFClient(Optional FormOpenDataModeAddArg As Boolean, Optional RecordIDArg As Long)
Debug.Print "Exec Priority - " & "TransactorsAddFCollectionModule - " & "OpenATransactorsAddFClient" & Time()
 
    On Error GoTo ErrorHandler

    ' Purpose: Open an independent instance of form TransactorsAddF.
    Dim frm As Form
    
    If FormOpenDataModeAddArg = True Then
    TempVars.Add "TransactorsAddFOpenModeAdd", True
    Else
    TempVars.Add "TransactorsAddFOpenModeAdd", False
    End If
    
        ' Open a new instance, show it, and set a caption.
        Set frm = New Form_TransactorsAddF
        frm.Visible = True
        frm.Caption = "Transactors Form, opened " & Now() & ", (ID = " & frm.Hwnd & ")"

        ' Append it to our collection.
        TransactorsAddFCollection.Add Item:=frm, Key:=CStr(frm.Hwnd)

       If FormOpenDataModeAddArg = False And RecordIDArg > 0 Then
         frm.DataEntry = False
         frm.AllowAdditions = False
         frm.AllowDeletions = True
         frm.AllowEdits = True
         
         frm.Requery
          Dim rst As DAO.Recordset
          Set rst = frm.RecordsetClone
                  rst.FindFirst "Transactor_ID = " & RecordIDArg
                If Not frm.RecordsetClone.NoMatch Then
                     frm.Bookmark = frm.RecordsetClone.Bookmark
                Else
                     MsgBox "Record with ID " & RecordIDArg & " not found.", vbExclamation
                     GoTo ExitProcedure
                End If
         Else
         frm.DataEntry = True
         frm.AllowAdditions = True
         frm.AllowDeletions = False
         frm.AllowEdits = False
        
         
       End If


ExitProcedure:
If Not IsNull(TempVars!TransactorsAddFOpenModeAdd) Then
TempVars.Remove "TransactorsAddFOpenModeAdd"
End If

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
            "Error Source: OpenATransactosAddFClient" & vbCrLf & _
            "Error Description: " & Err.Description, vbOKOnly + vbCritical, "An Error has Occurred!"
    Resume ExitProcedure
 
 
End Sub

Function CloseAllTransactorsAddFClients()
    'Purpose: Close all instances in the clnClient collection.
    'Note: Leaves the copy opened directly from database window.
Debug.Print "Exec Priority - " & "TransactorsAddFCollectionModule - " & "CloseAllTransactorsAddFClients" & Time()
On Error GoTo ErrorHandler

    Dim NumberOfMembers As Long
    Dim i As Long
    
    NumberOfMembers = TransactorsAddFCollection.Count
    For i = 1 To NumberOfMembers
        TransactorsAddFCollection.Remove 1
    Next
    
If CheckIfFormIsOpen("TransactorsAddF") Then
DoCmd.Close acForm, "TransactorsAddF", acSaveNo
End If
    
ExitProcedure:
Exit Function
   
ErrorHandler:
Select Case Err.Number
        
        
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: CloseAllTransactorsAddFClients" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select

End Function

Function CloseOneTransactorsAddFClients(HwndArg As Long)
  Debug.Print "Exec Priority - " & "TransactorsAddFCollectionModule - " & "CloseOneTransactorsAddFClients " & Time()
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
    For Each obj In TransactorsAddFCollection
        If obj.Hwnd = HwndArg Then
            blnRemove = True
            Exit For
        End If
    Next
    
       
    'Deassign the object before removing from collection.
    Set obj = Nothing
    If blnRemove Then
        TransactorsAddFCollection.Remove CStr(HwndArg)
        If CheckIfFormIsOpen(frm.Name) Then
            DoCmd.Close acForm, frm.Name, acSaveNo
        End If
    Else
        DoCmd.Close acForm, "TransactorsAddF", acSaveNo
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
            "Error Source: CloseOneTransactorsAddFClients" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select

End Function