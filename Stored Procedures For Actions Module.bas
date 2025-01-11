Option Compare Database

Public Sub CallActionsStoredProcedures(LinkStoredProcToActionsIDArg As Long, ActionIDArg As Long, Optional FormArg As Form, Optional ActionTypeIDArg As Long)
Debug.Print "StoredProdeduresForActionsModule - " & "CallActionsStoredProcedures " & Time()
If Not IsNull(LinkStoredProcToActionsIDArg) Then
Select Case LinkStoredProcToActionsIDArg
  Case 2
   If FormArg.Name = "ProtocolsF" Then
   InsertC4COriginalToWarehouse (ActionIDArg)
   End If
  Case 3
   TransactorEditProcedure (NumberArg)
   FormArg.Refresh
End Select
End If
End Sub

Public Sub InsertC4COriginalToWarehouse(ActionIDArg As Long) 'Opens Document Product Form
Debug.Print "StoredProdeduresForActionsModule - " & "InsertC4COriginalToWarehouse " & Time()
On Error GoTo ErrorHandler

Dim VarC4CIndividualID As Long

VarC4CIndividualID = CLng(BringAnyAttributeValue(333, 8, ActionIDArg))


ExitProcedure:
Exit Sub
   
ErrorHandler:
Select Case Err.Number
  
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: InsertC4COriginalToWarehouse " & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select

End Sub