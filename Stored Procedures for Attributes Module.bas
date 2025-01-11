Option Compare Database

Public Sub CallAttributesStoredProcedures(LinkStoredProcToAttributesIDArg As Long, Optional FormArg As Form, Optional AttributeIDArg As Long, Optional EntityIDArg As Long, Optional EntityTypeIDArg As Long, Optional EntityTypeID_For_RelevantTablePKFieldArg As Long, Optional TypeOfEntityType_IDArg As Long, Optional Link_Attribute_Value_To_Entity_IDArg As Long, Optional StringArg As String, Optional NumberArg As Double, Optional BooleanArg As Long, Optional DateArg As Date, Optional TimeArg As Date, Optional TimestampArg As Date)
Debug.Print "StoredProdeduresForAttributesModule - " & "CallAttributesStoredProcedures " & Time()
If Not IsNull(LinkStoredProcToAttributesIDArg) Then
Select Case LinkStoredProcToAttributesIDArg
  Case 2
   Dim NewTransactorID As Long
   NewTransactorID = AddNewTransactorFunction(TypeOfEntityType_IDArg)
   If FormArg.Name = "AttrValToProtocWithAttrFiltForProtocolsFormDSF" Then
   Forms!ProtocolsF!AttrValToProtocWithAttrFiltForProtocolsFormDSF!InputAttributeValuesCbo1 = NewTransactorID
   End If
  Case 3
   TransactorEditProcedure (NumberArg)
   FormArg.Refresh
End Select
End If
End Sub

Public Function AddNewTransactorFunction(Optional TransactorTypeIDArg As Long, Optional BasicTransactorIDArg As Long, Optional TextForNewTransactorDescriptionArg As String)
Debug.Print "StoredProdeduresForAttributesModule - " & "AddNewTransactorFunction " & Time()
On Error GoTo ErrorHandler

If CheckIfFormIsOpen("TransactorAddF") Then
MsgBox "Form TransactorAddF is already open. Please close the form and try again.", vbInformation, vbOK
GoTo ExitProcedure
End If

If Not IsNull(TransactorTypeIDArg) Then
TempVars.Add "TempVarTransactorTypeID", TransactorTypeIDArg
End If

If Not IsNull(BasicTransactorIDArg) Then
TempVars.Add "TempVars!TempVarBasicTransactorID", BasicTransactorIDArg
End If

If Not IsNull(TextForNewTransactorDescriptionArg) Then
TempVars.Add "TempVarTextForTransactorDescription", TextForNewTransactorDescriptionArg
End If

TempVars.Add "TempVarTransactorAddLastInsertedID", 0

DoCmd.OpenForm "TransactorsAddF", acNormal, , , acFormAdd, acDialog


ExitProcedure:
If Not IsNull(TempVars!TempVarTransactorTypeID) Then
TempVars.Remove (TempVarTransactorTypeID)
End If
If Not IsNull(TempVars!TempVarBasicTransactorID) Then
TempVars.Remove (TempVars!TempVarBasicTransactorID)
End If
If Not IsNull(TempVars!TempVarTextForTransactorDescription) Then
TempVars.Remove (TempVarTextForTransactorDescription)
End If
If Not IsNull(TempVars!TempVarTransactorAddLastInsertedID) Then
AddNewTransactorFunction = TempVars!TempVarTransactorAddLastInsertedID
TempVars.Remove (TempVarTransactorAddLastInsertedID)
End If

Exit Function
   
ErrorHandler:
Select Case Err.Number
  
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: AddNewTransactorFunction " & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select

End Function

Public Sub TransactorEditProcedure(TransactorIDArg As Long)
Debug.Print "StoredProdeduresForAttributesModule - " & "TransactorEditProcedure(TransactorIDArg " & Time()
On Error GoTo ErrorHandler

DoCmd.OpenForm "TransactorsAddF", acNormal, , "TransactorsT.Transactor_ID = " & TransactorIDArg, acFormEdit, acHidden
If CInt(Forms!TransactorsAddF.Form!TransactorIDTbox) = CInt(TransactorIDArg) Then
Forms!TransactorsAddF.Form.Visible = True
Forms!TransactorsAddF.SetFocus
Else
MsgBox "Transactor Record was not found!"
DoCmd.Close acForm, "TransactorsAddF", acSaveNo
End If

ExitProcedure:
Exit Sub
   
ErrorHandler:
Select Case Err.Number
  
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: TransactorEditProcedure(TransactorIDArg " & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
End Sub