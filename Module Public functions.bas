Option Compare Database

Public Sub YellowWhenActive(MyControl As Control, myform As Access.Form)
Debug.Print "Module Public Functions - " & "YellowWhenActive " & Time()
On Error GoTo ErrorHandler
Dim ctl As Control

For Each ctl In myform.Controls
   Select Case ctl.ControlType
      Case acTextBox, acComboBox, acListBox, acLabel
      ctl.BackColor = RGB(255, 255, 255)
   End Select
Next



If (MyControl Is myform.ActiveControl) Then
Select Case MyControl.ControlType
      Case acTextBox, acComboBox, acListBox
      MyControl.BackColor = RGB(255, 255, 0)
   End Select
End If

ExitProcedure:
Exit Sub
   
ErrorHandler:
Select Case Err.Number
        
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: YellowWhenActive" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
            
End Sub
Public Sub YellowWhenActiveForDetailsSectionOnly(MyControl As Control, myform As Access.Form)
On Error GoTo ErrorHandler
Debug.Print "Module Public Functions - " & "YellowWhenActiveForDetailsSectionOnly " & Time()
Dim ctl As Control

For Each ctl In myform.Detail.Controls
   Select Case ctl.ControlType
      Case acTextBox, acComboBox, acListBox, acLabel
      ctl.BackColor = RGB(255, 255, 255)
   End Select
Next



If (MyControl Is myform.ActiveControl) Then
Select Case MyControl.ControlType
      Case acTextBox, acComboBox, acListBox
      MyControl.BackColor = RGB(255, 255, 0)
   End Select
End If


ExitProcedure:
Exit Sub
   
ErrorHandler:
Select Case Err.Number
        Case 2474
        Resume Next
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: YellowWhenActiveForDetailsSectionOnly" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
            
End Sub
Public Sub YellowWhenActiveForHeaderSectionOnly(MyControl As Control, myform As Access.Form)
On Error GoTo ErrorHandler
Debug.Print "Module Public Functions - " & "YellowWhenActiveForHeaderSectionOnly " & Time()
Dim ctl As Control

For Each ctl In myform.Section(acHeader).Controls
   Select Case ctl.ControlType
      Case acTextBox, acComboBox, acListBox, acLabel
      ctl.BackColor = RGB(255, 255, 255)
   End Select
Next



If (MyControl Is myform.ActiveControl) Then
Select Case MyControl.ControlType
      Case acTextBox, acComboBox, acListBox
      MyControl.BackColor = RGB(255, 255, 0)
   End Select
End If


ExitProcedure:
Exit Sub
   
ErrorHandler:
Select Case Err.Number
        Case 2474
        Resume Next
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: YellowWhenActiveForHeaderSectionOnly" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
            
End Sub
Public Sub YellowWhenActiveForFooterSectionOnly(MyControl As Control, myform As Access.Form)
On Error GoTo ErrorHandler
Debug.Print "Module Public Functions - " & "YellowWhenActiveForFooterSectionOnly " & Time()
Dim ctl As Control

For Each ctl In myform.Section(acFooter).Controls
   Select Case ctl.ControlType
      Case acTextBox, acComboBox, acListBox, acLabel
      ctl.BackColor = RGB(255, 255, 255)
   End Select
Next



If (MyControl Is myform.ActiveControl) Then
Select Case MyControl.ControlType
      Case acTextBox, acComboBox, acListBox
      MyControl.BackColor = RGB(255, 255, 0)
   End Select
End If


ExitProcedure:
Exit Sub
   
ErrorHandler:
Select Case Err.Number
        Case 2474
        Resume Next
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: YellowWhenActiveForFooterSectionOnly" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
            
End Sub
Public Sub WhiteWhenInactive(MyControl As Control, myform As Access.Form)
Debug.Print "Module Public Functions - " & "WhiteWhenInactive " & Time()
If (MyControl Is myform.ActiveControl) Then
MyControl.BackColor = RGB(255, 255, 255)
End If
End Sub

Public Sub zoom()
Debug.Print "Module Public Functions - " & "zoom " & Time()
RunCommand acCmdZoomBox
End Sub

Public Sub ShowQueryResults(ByVal sql As String)
Debug.Print "Module Public Functions - " & "ShowQueryResults " & Time()
    Const Query_Name As String = "test_query1"

    Dim db As DAO.Database
    Dim qd As DAO.QueryDef


    Set db = CurrentDb
    Set qd = db.CreateQueryDef(Query_Name, sql)

    DoCmd.OpenQuery Query_Name

    If MsgBox("Close Query?", vbYesNo) = vbYes Then
        DoCmd.Close acQuery, Query_Name, acSaveNo

        DeleteQuery Query_Name
    End If

End Sub

Public Sub DeleteQuery(ByVal QueryName As String)
Debug.Print "Module Public Functions - " & "DeleteQuery " & Time()
'On Error Resume Next
CurrentDb.QueryDefs.Delete QueryName
End Sub
Public Function FetchUserID() As Integer
Debug.Print "Module Public Functions - " & "FetchUserID() " & Time()
On Error GoTo ErrorHandler

CheckUsers
FetchUserID = DLookup("[CurrentUserT]![Current_User_ID]", "[CurrentUserT]")

ExitProcedure:
Exit Function
   
ErrorHandler:
Select Case Err.Number
        Case 3314
        MsgBox "You have left empty fields which must be filled", vbInformation, "������ ���������"
        Response = acDataErrContinue
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: FetchUserID" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select

End Function
Public Sub CheckUsers()
Debug.Print "Module Public Functions - " & "CheckUsers " & Time()
Debug.Print "CheckUsers Record Number = " & DCount("CurrentUserT![Current_User_ID]", "CurrentUserT")
If DCount("CurrentUserT![Current_User_ID]", "CurrentUserT") <> 1 Then
  MsgBox "There is a problem with user selection. Please contact IT. Application will shut down.", vbOKOnly
   'DoCmd.Quit acQuitSaveAll
   CloseDatabase (False)
End If
End Sub


'The Form 's Key Preview property must be set to True for this code to work.
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' iKeyCode  : Keycode from the source form from the KeyDown event
' frm       : form object to apply the new behavior to
'
' Usage:
' ~~~~~~
' KeyCode = EnableArrowsScroll(KeyCode, Me) 'This is placed in the KeyDown event
'                            'Dont forget to set the Key Preview property to Yes
'CODE TO CALL IT WITH ERROR HANDLING IS PLACED IN "REUSABLE_CODE" MODULE WITH NAME "Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)" WHICH IS THE NAME OF THE EVENT OF EVERY FORMS' FORM KEY DOWN EVENT

Public Function EnableArrowsScroll(ByVal iKeyCode As Integer, frm As Access.Form) As Integer
    On Error GoTo Error_Handler_Exit
 Debug.Print "Module Public Functions - " & "EnableArrowsScroll " & Time()
    If frm.DefaultView = 1 Then    'Only process for Continuous forms
        Select Case iKeyCode
            Case vbKeyDown
                '            If CurrentRecord <> RecordsetClone.RecordCount Then 'Restrict to existing records
                If frm.NewRecord = False Then    'Allow going to new record for data entry
                    DoCmd.GoToRecord , , acNext
                End If
                EnableArrowsScroll = 0
                
                Case vbKeyUp
                If frm.currentRecord <> 1 Then
                    DoCmd.GoToRecord , , acPrevious
                   
                   End If
                   EnableArrowsScroll = 0
             
            Case Else
                EnableArrowsScroll = iKeyCode
        End Select
    Else
        EnableArrowsScroll = iKeyCode
    End If
 
    
Error_Handler_Exit:
    On Error Resume Next
    If Not frm Is Nothing Then Set frm = Nothing
    Exit Function
   
ErrorHandler:
Select Case Err.Number
        Case 2105
        Response = acDataErrContinue
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: EnableArrowsScroll" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume Error_Handler_Exit
            End Select

End Function

Function ComboListDisplay()
 Debug.Print "Module Public Functions - " & "ComboListDisplay " & Time()
    Dim MyControl As Control
    Set MyControl = Screen.ActiveControl
 
        If IsNull(MyControl) Then
            MyControl.dropdown
        End If
 
End Function

Public Function CheckIfTableExists(TableName As String) As Boolean
Debug.Print "Exec Priority - " & "Module Public Functions - " & "CheckIfTableExists " & Time()

On Error Resume Next
Dim tdf As TableDef

Set tdf = CurrentDb.TableDefs(TableName)
If Err.Number = 0 Then
CheckIfTableExists = True
Else
CheckIfTableExists = False
End If

End Function

Public Function EntityInfoForSubattributes(EntityTypeIDint As Integer, EntityIDint As Integer)
Debug.Print "Exec Priority - " & "Module Public Functions - " & "EntityInfoForSubattributes " & Time()

'Dim EntityIDint As Integer
Dim valuetobring As String
Debug.Print "EntityInfoForSubattributes, EntityTypeIDint = "; EntityTypeIDint & " EntityIDint = " & EntityIDint
'EntityIDint = EntityIDint

Select Case EntityTypeIDint

Case 1 ' product
   EntityInfoForSubattributes = "��������� : " & Nz(DLookup("ProductsSimpleQ.Product_Description", "ProductsSimpleQ", "ProductsSimpleQ.Product_ID = " & EntityIDint), "-") & "," & _
   vbCrLf & "Marketing Label : " & Nz(DLookup("ProductsSimpleQ.[Marketing Label]", "ProductsSimpleQ", "ProductsSimpleQ.Product_ID = " & EntityIDint), "-") & "," & _
   vbCrLf & "����� : " & Nz(DLookup("ProductsSimpleQ.Type_Description", "ProductsSimpleQ", "ProductsSimpleQ.Product_ID = " & EntityIDint), "-") & "," & _
   vbCrLf & "��������� : " & Nz(DLookup("ProductsSimpleQ.Category_Description", "ProductsSimpleQ", "ProductsSimpleQ.Product_ID = " & EntityIDint), "-") & "," & _
   vbCrLf & "������������ : " & Nz(DLookup("ProductsSimpleQ.Subcategory_Description", "ProductsSimpleQ", "ProductsSimpleQ.Product_ID = " & EntityIDint), "-") & "," & _
   vbCrLf & "Base Product ID : " & Nz(DLookup("ProductsSimpleQ.Base_Product_ID", "ProductsSimpleQ", "ProductsSimpleQ.Product_ID = " & EntityIDint), "-") & "," & _
   vbCrLf & "SKU : " & Nz(DLookup("ProductsSimpleQ.SKU", "ProductsSimpleQ", "ProductsSimpleQ.Product_ID = " & EntityIDint), "-")
Case 2 ' transactor
   EntityInfoForSubattributes = "��������� : " & Nz(DLookup("TransactorsWithBasicTransactorsDescriptionQ.Basic_Transactor_Description", "TransactorsWithBasicTransactorsDescriptionQ", "TransactorsWithBasicTransactorsDescriptionQ.Transactor_ID = " & EntityIDint), "-") & "," & _
   vbCrLf & "����� ��������������� : " & Nz(DLookup("TransactorsWithBasicTransactorsDescriptionQ.Transactor_Type_Desription", "TransactorsWithBasicTransactorsDescriptionQ", "TransactorsWithBasicTransactorsDescriptionQ.Transactor_ID = " & EntityIDint), "-")
Case 3 'document issued
   EntityInfoForSubattributes = "���������: " & Nz(DLookup("IssuedDocumentSimpleQ.Issuable_Document_Description", "IssuedDocumentSimpleQ", "IssuedDocumentSimpleQ.Issued_Document_ID = " & EntityIDint), "-") & "," & _
   vbCrLf & "��.: " & Nz(DLookup("IssuedDocumentSimpleQ.Document_Official_Number_Sequence", "IssuedDocumentSimpleQ", "IssuedDocumentSimpleQ.Issued_Document_ID = " & EntityIDint), "-") & "," & _
   " ����� : " & Nz(DLookup("IssuedDocumentSimpleQ.Distinction_Additive", "IssuedDocumentSimpleQ", "IssuedDocumentSimpleQ.Issued_Document_ID = " & EntityIDint), "-") & "," & _
   " ��/��� : " & Nz(DLookup("IssuedDocumentSimpleQ.Issued_Date", "IssuedDocumentSimpleQ", "IssuedDocumentSimpleQ.Issued_Document_ID = " & EntityIDint), "-") & "," & _
   vbCrLf & "������� : " & Nz(DLookup("IssuedDocumentSimpleQ.Document_Issuer_Description", "IssuedDocumentSimpleQ", "IssuedDocumentSimpleQ.Issued_Document_ID = " & EntityIDint), "-") & "," & _
   vbCrLf & "������ : " & Nz(DLookup("IssuedDocumentSimpleQ.Intention_Description", "IssuedDocumentSimpleQ", "IssuedDocumentSimpleQ.Issued_Document_ID = " & EntityIDint), "-")
Case 4 ' document financial details
   If IsNull(DLookup("IssuedDocumentFinancialDetailsSimpleQ.Debit", "IssuedDocumentFinancialDetailsSimpleQ", "IssuedDocumentFinancialDetailsSimpleQ.IssuedDocumentFinancialDetailsT.Issued_Document_Financial_Details_ID = " & EntityIDint)) Then
      valuetobring = "�������: " & Format(DLookup("IssuedDocumentFinancialDetailsSimpleQ.Debit", "IssuedDocumentFinancialDetailsSimpleQ", "IssuedDocumentFinancialDetailsSimpleQ.[IssuedDocumentFinancialDetailsT.Issued_Document_Financial_Details_ID] = " & EntityIDint), "currency")
   Else
      valuetobring = "��������: " & Format(DLookup("IssuedDocumentFinancialDetailsSimpleQ.Credit", "IssuedDocumentFinancialDetailsSimpleQ", "IssuedDocumentFinancialDetailsSimpleQ.[IssuedDocumentFinancialDetailsT.Issued_Document_Financial_Details_ID] = " & EntityIDint), "currency")
   End If
   EntityInfoForSubattributes = "���������: " & Nz(DLookup("IssuedDocumentFinancialDetailsSimpleQ.Issuable_Document_Description", "IssuedDocumentFinancialDetailsSimpleQ", "IssuedDocumentFinancialDetailsSimpleQ.[IssuedDocumentFinancialDetailsT.Issued_Document_Financial_Details_ID] = " & EntityIDint), "-") & "," & _
   vbCrLf & "��.: " & Nz(DLookup("IssuedDocumentFinancialDetailsSimpleQ.Document_Official_Number_Sequence", "IssuedDocumentFinancialDetailsSimpleQ", "IssuedDocumentFinancialDetailsSimpleQ.[IssuedDocumentFinancialDetailsT.Issued_Document_Financial_Details_ID] = " & EntityIDint), "-") & "," & _
   " ����� : " & Nz(DLookup("IssuedDocumentFinancialDetailsSimpleQ.Distinction_Additive", "IssuedDocumentFinancialDetailsSimpleQ", "IssuedDocumentFinancialDetailsSimpleQ.[IssuedDocumentFinancialDetailsT.Issued_Document_Financial_Details_ID] = " & EntityIDint), "-") & "," & _
   " ��/��� : " & Nz(DLookup("IssuedDocumentFinancialDetailsSimpleQ.Issued_Date", "IssuedDocumentFinancialDetailsSimpleQ", "IssuedDocumentFinancialDetailsSimpleQ.[IssuedDocumentFinancialDetailsT.Issued_Document_Financial_Details_ID] = " & EntityIDint), "-") & "," & _
   vbCrLf & "������� : " & Nz(DLookup("IssuedDocumentFinancialDetailsSimpleQ.Document_Issuer_Description", "IssuedDocumentFinancialDetailsSimpleQ", "IssuedDocumentFinancialDetailsSimpleQ.[IssuedDocumentFinancialDetailsT.Issued_Document_Financial_Details_ID] = " & EntityIDint), "-") & "," & _
   vbCrLf & "������ : " & Nz(DLookup("IssuedDocumentFinancialDetailsSimpleQ.Intention_Description", "IssuedDocumentFinancialDetailsSimpleQ", "IssuedDocumentFinancialDetailsSimpleQ.[IssuedDocumentFinancialDetailsT.Issued_Document_Financial_Details_ID] = " & EntityIDint), "-") & "," & _
   vbCrLf & "���/������� : " & Nz(DLookup("IssuedDocumentFinancialDetailsSimpleQ.Basic_Transactor_Description", "IssuedDocumentFinancialDetailsSimpleQ", "IssuedDocumentFinancialDetailsSimpleQ.[IssuedDocumentFinancialDetailsT.Issued_Document_Financial_Details_ID] = " & EntityIDint), "-") & "," & _
   " ID : " & Nz(DLookup("IssuedDocumentFinancialDetailsSimpleQ.Transactor_ID", "IssuedDocumentFinancialDetailsSimpleQ", "IssuedDocumentFinancialDetailsSimpleQ.[IssuedDocumentFinancialDetailsT.Issued_Document_Financial_Details_ID] = " & EntityIDint), "-") & "," & _
   " ����� : " & Nz(DLookup("IssuedDocumentFinancialDetailsSimpleQ.Transactor_Type_Desription", "IssuedDocumentFinancialDetailsSimpleQ", "IssuedDocumentFinancialDetailsSimpleQ.[IssuedDocumentFinancialDetailsT.Issued_Document_Financial_Details_ID] = " & EntityIDint), "-") & "," & _
   vbCrLf & "���� " & valuetobring
Case 5 'document product details
   EntityInfoForSubattributes = "Product Description : " & Nz(DLookup("IssuedDocumentProductDetailsSimpleQ.Product_Description", "IssuedDocumentProductDetailsSimpleQ", "IssuedDocumentProductDetailsSimpleQ.[IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID] = " & EntityIDint), "-") & "," & _
   " ID : " & Nz(DLookup("IssuedDocumentProductDetailsSimpleQ.IssuedDocumentProductDetailsT.Product_ID", "IssuedDocumentProductDetailsSimpleQ", "IssuedDocumentProductDetailsSimpleQ.[IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID] = " & EntityIDint), "-") & "," & _
   vbCrLf & "Document Description: " & Nz(DLookup("IssuedDocumentProductDetailsSimpleQ.Issuable_Document_Description", "IssuedDocumentProductDetailsSimpleQ", "IssuedDocumentProductDetailsSimpleQ.[IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID] = " & EntityIDint), "-") & "," & _
   vbCrLf & "��.: " & Nz(DLookup("IssuedDocumentProductDetailsSimpleQ.Document_Official_Number_Sequence", "IssuedDocumentProductDetailsSimpleQ", "IssuedDocumentProductDetailsSimpleQ.[IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID] = " & EntityIDint), "-") & "," & _
   vbCrLf & " Distinctive Additive : " & Nz(DLookup("IssuedDocumentProductDetailsSimpleQ.Distinction_Additive", "IssuedDocumentProductDetailsSimpleQ", "IssuedDocumentProductDetailsSimpleQ.[IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID] = " & EntityIDint), "-") & "," & _
   vbCrLf & " Date : " & Nz(DLookup("IssuedDocumentProductDetailsSimpleQ.Issued_Date", "IssuedDocumentProductDetailsSimpleQ", "IssuedDocumentProductDetailsSimpleQ.[IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID] = " & EntityIDint), "-") & "," & _
   vbCrLf & "Issuer : " & Nz(DLookup("IssuedDocumentProductDetailsSimpleQ.Document_Issuer_Description", "IssuedDocumentProductDetailsSimpleQ", "IssuedDocumentProductDetailsSimpleQ.[IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID] = " & EntityIDint), "-") & "," & _
   vbCrLf & "Inention : " & Nz(DLookup("IssuedDocumentProductDetailsSimpleQ.Intention_Description", "IssuedDocumentProductDetailsSimpleQ", "IssuedDocumentProductDetailsSimpleQ.[IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID] = " & EntityIDint), "-") & "," & _
   vbCrLf & "Quantity : " & Nz(DLookup("IssuedDocumentProductDetailsSimpleQ.IssuedDocumentProductDetailsT.Quantity", "IssuedDocumentProductDetailsSimpleQ", "IssuedDocumentProductDetailsSimpleQ.[IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID] = " & EntityIDint), "-") & "," & _
   vbCrLf & "Value Type : " & IIf(DLookup("IssuedDocumentProductDetailsSimpleQ.Gross_Or_Net_Values", "IssuedDocumentProductDetailsSimpleQ", "IssuedDocumentProductDetailsSimpleQ.[IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID] = " & EntityIDint) = True, "Gross", "Net") & "," & _
   vbCrLf & "Unit Value After Discount : " & Nz(Format(DLookup("IssuedDocumentProductDetailsSimpleQ.Unit_Price_After_Discount", "IssuedDocumentProductDetailsSimpleQ", "IssuedDocumentProductDetailsSimpleQ.[IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID] = " & EntityIDint), "currency"), 0) & "," & _
   vbCrLf & "Total Value After Discount : " & Nz(Format(DLookup("IssuedDocumentProductDetailsSimpleQ.Total_Value_After_Discount", "IssuedDocumentProductDetailsSimpleQ", "IssuedDocumentProductDetailsSimpleQ.[IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID] = " & EntityIDint), "currency"), 0) & "," & _
   vbCrLf & "Transaction Point : " & Nz(DLookup("IssuedDocumentProductDetailsSimpleQ.[Transaction_Point_Description]", "IssuedDocumentProductDetailsSimpleQ", "IssuedDocumentProductDetailsSimpleQ.[IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID] = " & EntityIDint), "-") & "," & _
   vbCrLf & "Product Transactor (Main) : " & Nz(DLookup("IssuedDocumentProductDetailsSimpleQ.[Basic_Transactor_Description]", "IssuedDocumentProductDetailsSimpleQ", "IssuedDocumentProductDetailsSimpleQ.[IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID] = " & EntityIDint), "-") & "," & _
   " Product Transactor Type : " & Nz(DLookup("IssuedDocumentProductDetailsSimpleQ.[Transactor_Type_Desription]", "IssuedDocumentProductDetailsSimpleQ", "IssuedDocumentProductDetailsSimpleQ.[IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID] = " & EntityIDint), "-")
Case 6 'transaction
   EntityInfoForSubattributes = "����� ����������: " & Nz(DLookup("TransactionsSimpleQ.Transaction_Type_Description", "TransactionsSimpleQ", "TransactionsSimpleQ.Transaction_ID = " & EntityIDint), "-") & "," & _
   vbCrLf & "�� ���������� : " & EntityIDint
Case 8 ' Action
Dim ActionTypeIDVar As Integer
ActionTypeIDVar = Nz(DLookup("ActionsT.Action_Type_ID", "ActionsT", "ActionsT.Action_ID = " & EntityIDint), "0")
    EntityInfoForSubattributes = "Action Description : " & Nz(DLookup("ActionTypesT.Action_Type_Description", "ActionTypesT", "ActionTypesT.Action_Type_ID = " & ActionTypeIDVar), "-") & "," & _
   vbCrLf & "Protocol Type : " & Nz(DLookup("ActionsSimpleQ.Protocol_Type_Description", "ActionsSimpleQ", "ActionsSimpleQ.Action_ID = " & EntityIDint), "-") & "," & _
   vbCrLf & "Protocol ID : " & Nz(DLookup("ActionsSimpleQ.Protocol_ID", "ActionsSimpleQ", "ActionsSimpleQ.Action_ID = " & EntityIDint), "-") & "," & _
   vbCrLf & "Timestamp_Assigned : " & Nz(DLookup("ActionsSimpleQ.Timestamp_Assigned", "ActionsSimpleQ", "ActionsSimpleQ.Action_ID = " & EntityIDint), "-")
Case 9 ' Protocol
     EntityInfoForSubattributes = "Protocol_Type_Description : " & Nz(DLookup("ProtocolsSimpleQ.Protocol_Type_Description", "ProtocolsSimpleQ", "ProtocolsSimpleQ.Protocol_ID = " & EntityIDint), "-") & "," & _
   vbCrLf & "Protocol ID : " & Nz(DLookup("ProtocolsSimpleQ.Protocol_ID", "ProtocolsSimpleQ", "ProtocolsSimpleQ.Protocol_ID = " & EntityIDint), "-") & "," & _
   vbCrLf & "Timestamp_Assigned : " & Nz(DLookup("ProtocolsSimpleQ.Assignment_Date", "ProtocolsSimpleQ", "ProtocolsSimpleQ.Protocol_ID = " & EntityIDint), "-")
Case 10 ' Installation
Case Else
   EntityInfoForSubattributes = ""
  
  End Select
  
End Function

Public Sub OpenAttributeValuesToEntityMainF(EntityIDint As Integer, EntityTypeIDint As Integer, AddOREditDatamode As Integer) ' AddOREditDatamode 1 = Add, 2 = Edit
Debug.Print "Exec Priority - " & "Module Public Functions - " & "OpenAttributeValuesToEntityMainF " & Time()

'Debug.Print "EntityIDint = " & EntityIDint
'Debug.Print "EntityTypeIDint = " & EntityTypeIDint
Dim FormRecordSource As String
If IsNull(EntityIDint) Or IsNull(EntityTypeIDint) Then
  MsgBox "Entity Type info has not been transferred. Attributes Form will not open.", vbOKOnly, "ATTENTION"
  DoCmd.Close acForm, "AttributeValuesToEntityMainF", acSaveNo
  Exit Sub
  Else
  'Debug.Print TempVars!EntityTypeIDForAttributeValuesToEntityMainF & " - " & TempVars!EntityIDForAttributeValuesToEntityMainF
 Select Case EntityTypeIDint
   
   Case 1  ' products
   
    FormRecordSource = "SELECT ProductsSimpleQ1.Product_ID AS Entity_ID, ProductsSimpleQ1.EntityTypesToHaveAttributesID AS Entity_Type_ID, EntitiesTypesToHaveAttributesT.Entities_To_Have_Attributes_Description " & _
" from " & _
"(SELECT ProductsSimpleQ.*, 1 AS EntityTypesToHaveAttributesID " & _
"FROM ProductsSimpleQ where ProductsSimpleQ.Product_ID = " & EntityIDint & ") as ProductsSimpleQ1 " & _
"INNER JOIN EntitiesTypesToHaveAttributesT on ProductsSimpleQ1.EntityTypesToHaveAttributesID = EntitiesTypesToHaveAttributesT.Entities_To_Have_Attributes_ID;"

    Case 2  ' transactors
    
   FormRecordSource = "SELECT TransactorsSimpleQ1.Transactor_ID AS Entity_ID, TransactorsSimpleQ1.EntityTypesToHaveAttributesID AS Entity_Type_ID, EntitiesTypesToHaveAttributesT.Entities_To_Have_Attributes_Description " & _
" from " & _
"(SELECT TransactorsWithBasicTransactorsDescriptionQ.*, 2 AS EntityTypesToHaveAttributesID " & _
"FROM TransactorsWithBasicTransactorsDescriptionQ where TransactorsWithBasicTransactorsDescriptionQ.Transactor_ID = " & EntityIDint & ") as TransactorsSimpleQ1 " & _
"INNER JOIN EntitiesTypesToHaveAttributesT on TransactorsSimpleQ1.EntityTypesToHaveAttributesID = EntitiesTypesToHaveAttributesT.Entities_To_Have_Attributes_ID;"

    Case 3   ' Issued Document
    
      FormRecordSource = "SELECT IssuedDocumentSimpleQ1.Issued_Document_ID AS Entity_ID,  " & _
    " IssuedDocumentSimpleQ1.EntityTypesToHaveAttributesID AS Entity_Type_ID, " & _
    " EntitiesTypesToHaveAttributesT.Entities_To_Have_Attributes_Description " & _
    " from " & _
    "(SELECT IssuedDocumentSimpleQ.*, 3 AS EntityTypesToHaveAttributesID " & _
    "FROM IssuedDocumentSimpleQ where IssuedDocumentSimpleQ.Issued_Document_ID = " & EntityIDint & ") as IssuedDocumentSimpleQ1 " & _
    "INNER JOIN EntitiesTypesToHaveAttributesT on IssuedDocumentSimpleQ1.EntityTypesToHaveAttributesID = EntitiesTypesToHaveAttributesT.Entities_To_Have_Attributes_ID;"

    Case 4   ' Document Financial Details
   

      FormRecordSource = "SELECT IssuedDocumentFinancialDetailsSimpleQ1.IssuedDocumentFinancialDetailsT.Issued_Document_Financial_Details_ID AS Entity_ID, " & _
    "IssuedDocumentFinancialDetailsSimpleQ1.EntityTypesToHaveAttributesID AS Entity_Type_ID, " & _
    " EntitiesTypesToHaveAttributesT.Entities_To_Have_Attributes_Description " & _
    " from " & _
    "(SELECT IssuedDocumentFinancialDetailsSimpleQ.*, 4 AS EntityTypesToHaveAttributesID " & _
    "FROM IssuedDocumentFinancialDetailsSimpleQ where IssuedDocumentFinancialDetailsSimpleQ.IssuedDocumentFinancialDetailsT.Issued_Document_Financial_Details_ID = " & EntityIDint & ") as IssuedDocumentFinancialDetailsSimpleQ1 " & _
    "INNER JOIN EntitiesTypesToHaveAttributesT on IssuedDocumentFinancialDetailsSimpleQ1.EntityTypesToHaveAttributesID = EntitiesTypesToHaveAttributesT.Entities_To_Have_Attributes_ID;"

    
    Case 5   ' Document Product Details
    
      FormRecordSource = "SELECT IssuedDocumentProductDetailsSimpleQ1.IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID AS Entity_ID, " & _
    "IssuedDocumentProductDetailsSimpleQ1.EntityTypesToHaveAttributesID AS Entity_Type_ID, " & _
    " EntitiesTypesToHaveAttributesT.Entities_To_Have_Attributes_Description " & _
    " from " & _
    "(SELECT IssuedDocumentProductDetailsSimpleQ.*, 5 AS EntityTypesToHaveAttributesID " & _
    "FROM IssuedDocumentProductDetailsSimpleQ where IssuedDocumentProductDetailsSimpleQ.IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID = " & EntityIDint & ") as IssuedDocumentProductDetailsSimpleQ1 " & _
    "INNER JOIN EntitiesTypesToHaveAttributesT on IssuedDocumentProductDetailsSimpleQ1.EntityTypesToHaveAttributesID = EntitiesTypesToHaveAttributesT.Entities_To_Have_Attributes_ID;"
    
    Case 6   ' Transaction Details
    
   FormRecordSource = "SELECT TransactionsQ1.Transaction_ID AS Entity_ID, TransactionsQ1.EntityTypesToHaveAttributesID AS Entity_Type_ID, EntitiesTypesToHaveAttributesT.Entities_To_Have_Attributes_Description " & _
" from " & _
"(SELECT TransactionsT.*, 6 AS EntityTypesToHaveAttributesID " & _
"FROM TransactionsT where TransactionsT.Transaction_ID = " & EntityIDint & ") as TransactionsQ1 " & _
"INNER JOIN EntitiesTypesToHaveAttributesT on TransactionsQ1.EntityTypesToHaveAttributesID = EntitiesTypesToHaveAttributesT.Entities_To_Have_Attributes_ID;"
    
      Case 8   ' Actions
    
  FormRecordSource = "SELECT ActionsSimpleQ1.ActionsT.Action_ID AS Entity_ID, " & _
"ActionsSimpleQ1.EntityTypesToHaveAttributesID AS Entity_Type_ID, EntitiesTypesToHaveAttributesT.Entities_To_Have_Attributes_Description " & _
" from " & _
"(SELECT  ActionsT.*, 8 AS EntityTypesToHaveAttributesID " & _
"FROM ActionsT where ActionsT.Action_ID = " & EntityIDint & ") as ActionsSimpleQ1 " & _
"INNER JOIN EntitiesTypesToHaveAttributesT on ActionsSimpleQ1.EntityTypesToHaveAttributesID = EntitiesTypesToHaveAttributesT.Entities_To_Have_Attributes_ID;"

     Case 9   ' Protocols
    
   FormRecordSource = "SELECT ProtocolsSimpleQ1.ProtocolsT.Protocol_ID AS Entity_ID, " & _
"ProtocolsSimpleQ1.EntityTypesToHaveAttributesID AS Entity_Type_ID, EntitiesTypesToHaveAttributesT.Entities_To_Have_Attributes_Description " & _
" from " & _
"(SELECT ProtocolsT.*, 9 AS EntityTypesToHaveAttributesID " & _
"FROM ProtocolsT where ProtocolsT.Protocol_ID = " & EntityIDint & ") as ProtocolsSimpleQ1 " & _
"INNER JOIN EntitiesTypesToHaveAttributesT on ProtocolsSimpleQ1.EntityTypesToHaveAttributesID = EntitiesTypesToHaveAttributesT.Entities_To_Have_Attributes_ID;"

     Case 10   ' Installations
    
   FormRecordSource = "SELECT InstallationsSimpleQ1.Installation_ID AS Entity_ID, " & _
"InstallationsSimpleQ1.EntityTypesToHaveAttributesID AS Entity_Type_ID, EntitiesTypesToHaveAttributesT.Entities_To_Have_Attributes_Description " & _
" from " & _
"(SELECT InstallationsT.*, 10 AS EntityTypesToHaveAttributesID " & _
"FROM InstallationsT where InstallationsT.Installation_ID = " & EntityIDint & ") as InstallationsSimpleQ1 " & _
"INNER JOIN EntitiesTypesToHaveAttributesT on InstallationsSimpleQ1.EntityTypesToHaveAttributesID = EntitiesTypesToHaveAttributesT.Entities_To_Have_Attributes_ID;"


  End Select
  
 If AddOREditDatamode = 1 Then
 DoCmd.OpenForm "AttributeValuesToEntityMainF", , , , acFormAdd, acWindowNormal, FormRecordSource
 Else
 If AddOREditDatamode = 2 Then
 DoCmd.OpenForm "AttributeValuesToEntityMainF", , , , acFormEdit, acWindowNormal, FormRecordSource
 End If
 End If


End If
   

End Sub

Public Function CheckForUnlinkedRecords()
Debug.Print "Module Public Functions - " & "CheckForUnlinkedRecords " & Time()
On Error GoTo ErrorHandler

CheckForUnlinkedRecords = 0
If Not CurrentProject.AllForms("CheckAllDataLinksMainF").IsLoaded Then
DoCmd.OpenForm "CheckAllDataLinksMainF", acNormal, , , acFormReadOnly, acHidden
Forms!CheckAllDataLinksMainF.Requery
Forms!CheckAllDataLinksMainF.Refresh
CheckForUnlinkedRecords = Forms!CheckAllDataLinksMainF!TotalUnlinkedRecordsTbox
DoCmd.Close acForm, "CheckForTransactionsWithoutDocumentsF", acSaveNo
Else
Forms!CheckAllDataLinksMainF.Requery
Forms!CheckAllDataLinksMainF.Refresh
CheckForUnlinkedRecords = Forms!CheckAllDataLinksMainF!TotalUnlinkedRecordsTbox
End If

ExitProcedure:
Exit Function
   
ErrorHandler:
Select Case Err.Number
        
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: CheckForUnlinkedRecords" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
   
End Function


Public Sub CopyFinancialDetailsToBackupTable(IssuedDocumentIDArg As Integer)
Debug.Print "Module Public Functions - " & "CopyFinancialDetailsToBackupTable " & Time()
'On Error GoTo ErrorHandler

Dim db As DAO.Database
Dim rstFinancialDetailsRecordsToInsertTobackupTable As DAO.Recordset
Dim rstForLastRecordInsertedToDocumentFinancialDetailBackupTable As DAO.Recordset
Dim VarlastDocumentFinancialDetailsBackupID As Long
Dim AppendToFinancialDetailsBackUpTableQ As QueryDef
Dim VarCurrentUser As Integer
Dim VarCurrentTimestamp As Date

VarCurrentUser = DLookup("[CurrentUserT]![Current_User_ID]", "[CurrentUserT]", "[Current_User_ID]>0")
VarCurrentTimestamp = Now()

'We insert into back up table ALL the records (financial details)
'So, first we make a recordset with all the records that we are about to insert to the backup table
 Set db = CurrentDb
 Set rstFinancialDetailsRecordsToInsertTobackupTable = db.OpenRecordset("Select * from IssuedDocumentFinancialDetailsT WHERE Issued_Document_ID = " & IssuedDocumentIDArg & " AND Is_New = false and IssuedDocumentFinancialDetails_Backup_ID is null")
   If Not rstFinancialDetailsRecordsToInsertTobackupTable.EOF Then
      rstFinancialDetailsRecordsToInsertTobackupTable.MoveLast
      rstFinancialDetailsRecordsToInsertTobackupTable.MoveFirst
'Then we itterate the recordset and insert its records one by one to the backup table
    Do Until rstFinancialDetailsRecordsToInsertTobackupTable.EOF
    Call CopyLinkAttributeValueToEntitiesToBackupTable(4, rstFinancialDetailsRecordsToInsertTobackupTable(0))
      Set AppendToFinancialDetailsBackUpTableQ = CurrentDb.QueryDefs("DocumentFinancialDetailsSaveBeforeEditQ")
      AppendToFinancialDetailsBackUpTableQ.Parameters(0) = VarCurrentTimestamp
      AppendToFinancialDetailsBackUpTableQ.Parameters(1) = VarCurrentUser
      AppendToFinancialDetailsBackUpTableQ.Parameters(2) = rstFinancialDetailsRecordsToInsertTobackupTable(0)
      AppendToFinancialDetailsBackUpTableQ.Execute dbFailOnError
      
   'After each one insertion to the backup table, we take the ID of this insert to the back up table and we update the rstFinancialDetailsRecordsToInsertTobackupTable recordset (the basic table IssuedDocumentFinancialDetailsT)

    Set rstForLastRecordInsertedToDocumentFinancialDetailBackupTable = db.OpenRecordset("SELECT @@IDENTITY")
    VarlastDocumentFinancialDetailsBackupID = rstForLastRecordInsertedToDocumentFinancialDetailBackupTable(0)
    rstFinancialDetailsRecordsToInsertTobackupTable.Edit
    rstFinancialDetailsRecordsToInsertTobackupTable!IssuedDocumentFinancialDetails_Backup_ID = VarlastDocumentFinancialDetailsBackupID
    rstFinancialDetailsRecordsToInsertTobackupTable.Update
    rstFinancialDetailsRecordsToInsertTobackupTable.MoveNext
  Loop
  
  ' We close all objects
    AppendToFinancialDetailsBackUpTableQ.Close
  End If
   
ExitProcedure:
If Not rstFinancialDetailsRecordsToInsertTobackupTable Is Nothing Then
    rstFinancialDetailsRecordsToInsertTobackupTable.Close
    Set rstFinancialDetailsRecordsToInsertTobackupTable = Nothing
End If
    
If Not rstForLastRecordInsertedToDocumentFinancialDetailBackupTable Is Nothing Then
    rstForLastRecordInsertedToDocumentFinancialDetailBackupTable.Close
    Set rstForLastRecordInsertedToDocumentFinancialDetailBackupTable = Nothing
End If

If Not AppendToFinancialDetailsBackUpTableQ Is Nothing Then
    Set AppendToFinancialDetailsBackUpTableQ = Nothing
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
            "Error Source: CopyFinancialDetailsToBackupTable" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
   
End Sub

Public Sub CopyIssuedDocumentToBackupTable(TransactionIDArg As Integer, Optional IssuedDocumentIDArg As Integer)
Debug.Print "Module Public Functions - " & "CopyIssuedDocumentToBackupTable " & Time()
'On Error GoTo ErrorHandler

Dim db As DAO.Database
Dim rstIssuedDocumentRecordsToInsertTobackupTable As DAO.Recordset
Dim rstForLastRecordInsertedToIssuedDocumentBackupTable As DAO.Recordset
Dim VarlastIssuedDocumentBackupID As Long
Dim AppendToIssuedDocumentBackUpTableQ As QueryDef
Dim VarCurrentUser As Integer
Dim VarCurrentTimestamp As Date

VarCurrentUser = DLookup("[CurrentUserT]![Current_User_ID]", "[CurrentUserT]", "[Current_User_ID]>0")
VarCurrentTimestamp = Now()

'We insert into back up table ALL the records (issued documents)
'So, first we make a recordset with the records that we are about to insert to the backup table
 Set db = CurrentDb
 
 If IssuedDocumentIDArg <> 0 Then
 Set rstIssuedDocumentRecordsToInsertTobackupTable = db.OpenRecordset("Select * from IssuedDocumentT WHERE Issued_Document_ID = " & IssuedDocumentIDArg & " AND Is_New = false and Issued_Document_Backup_ID is null")
 Else
 Set rstIssuedDocumentRecordsToInsertTobackupTable = db.OpenRecordset("Select * from IssuedDocumentT WHERE Transaction_ID = " & TransactionIDArg & " AND Is_New = false and Issued_Document_Backup_ID is null")
 End If
 
   If Not rstIssuedDocumentRecordsToInsertTobackupTable.EOF Then
      rstIssuedDocumentRecordsToInsertTobackupTable.MoveLast
      rstIssuedDocumentRecordsToInsertTobackupTable.MoveFirst
'Then we itterate the recordset and insert its records one by one to the backup table
    Do Until rstIssuedDocumentRecordsToInsertTobackupTable.EOF
    
    Call CopyLinkAttributeValueToEntitiesToBackupTable(3, rstIssuedDocumentRecordsToInsertTobackupTable(0))
    Call CopyFinancialDetailsToBackupTable(rstIssuedDocumentRecordsToInsertTobackupTable(0))
    Call CopyProductDetailsToBackupTable(rstIssuedDocumentRecordsToInsertTobackupTable(0))
    
      Set AppendToIssuedDocumentBackUpTableQ = CurrentDb.QueryDefs("IssuedDocumentSaveBeforeEditQ")
      AppendToIssuedDocumentBackUpTableQ.Parameters(0) = VarCurrentTimestamp
      AppendToIssuedDocumentBackUpTableQ.Parameters(1) = VarCurrentUser
      AppendToIssuedDocumentBackUpTableQ.Parameters(2) = rstIssuedDocumentRecordsToInsertTobackupTable(0)
      AppendToIssuedDocumentBackUpTableQ.Execute dbFailOnError
      
   'After each one insertion to the backup table, we take the ID of this insert to the back up table and we update the rstIssuedDocumentRecordsToInsertTobackupTable recordset (the basic table IssuedDocumentFinancialDetailsT)

    Set rstForLastRecordInsertedToIssuedDocumentBackupTable = db.OpenRecordset("SELECT @@IDENTITY")
    VarlastIssuedDocumentBackupID = rstForLastRecordInsertedToIssuedDocumentBackupTable(0)
    rstIssuedDocumentRecordsToInsertTobackupTable.Edit
    rstIssuedDocumentRecordsToInsertTobackupTable!Issued_Document_Backup_ID = VarlastIssuedDocumentBackupID
    rstIssuedDocumentRecordsToInsertTobackupTable.Update
    rstIssuedDocumentRecordsToInsertTobackupTable.MoveNext
  Loop
    
  End If
   

ExitProcedure:
If Not AppendToIssuedDocumentBackUpTableQ Is Nothing Then
    AppendToIssuedDocumentBackUpTableQ.Close
    Set AppendToIssuedDocumentBackUpTableQ = Nothing
End If
    
If Not rstForLastRecordInsertedToIssuedDocumentBackupTable Is Nothing Then
    rstForLastRecordInsertedToIssuedDocumentBackupTable.Close
    Set rstForLastRecordInsertedToIssuedDocumentBackupTable = Nothing
End If

If Not rstIssuedDocumentRecordsToInsertTobackupTable Is Nothing Then
    rstIssuedDocumentRecordsToInsertTobackupTable.Close
    Set rstIssuedDocumentRecordsToInsertTobackupTable = Nothing
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
            "Error Source: CopyIssuedDocumentToBackupTable" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
   
   
End Sub


Public Sub CopyDiscountLogsToBackupTable(IssuedDocumentIDArg As Long)
Debug.Print "Module Public Functions - " & "CopyDiscountLogsToBackupTable " & Time()
'On Error GoTo ErrorHandler

Dim db As DAO.Database
Dim rstDiscountLogsRecordsToInsertTobackupTable As DAO.Recordset
Dim rstForLastRecordInsertedToDiscountLogsBackupTable As DAO.Recordset
Dim VarlastDicountLogBackupID As Long
Dim AppendToDicountLogsBackUpTableQ As QueryDef
Dim VarCurrentUser As Integer
Dim VarCurrentTimestamp As Date

VarCurrentUser = DLookup("[CurrentUserT]![Current_User_ID]", "[CurrentUserT]", "[Current_User_ID]>0")
VarCurrentTimestamp = Now()

'We insert into back up table ALL the records (DiscountLogs)
'So, first we make a recordset with all the records that we are about to insert to the backup table
 Set db = CurrentDb
 Set rstDiscountLogsRecordsToInsertTobackupTable = db.OpenRecordset("Select * from DiscountLogsT WHERE Issued_Document_ID = " & IssuedDocumentIDArg & " AND IS_DELETED = FALSE and is_new = false and DiscountLogs_Backup_ID is null")
   If Not rstDiscountLogsRecordsToInsertTobackupTable.EOF Then
      rstDiscountLogsRecordsToInsertTobackupTable.MoveLast
      rstDiscountLogsRecordsToInsertTobackupTable.MoveFirst
'Then we itterate the recordset and insert its records one by one to the backup table
    Do Until rstDiscountLogsRecordsToInsertTobackupTable.EOF
    CopyDiscountLogsDetailsToBackupTable (rstDiscountLogsRecordsToInsertTobackupTable(0))
      Set AppendToDicountLogsBackUpTableQ = CurrentDb.QueryDefs("DiscountLogsSaveBeforeEditQ")
      AppendToDicountLogsBackUpTableQ.Parameters(0) = VarCurrentTimestamp
      AppendToDicountLogsBackUpTableQ.Parameters(1) = VarCurrentUser
      AppendToDicountLogsBackUpTableQ.Parameters(2) = rstDiscountLogsRecordsToInsertTobackupTable(0)
      AppendToDicountLogsBackUpTableQ.Execute dbFailOnError
      
   'After each one insertion to the backup table, we take the ID of this insert to the back up table and we update the rstDiscountLogsRecordsToInsertTobackupTable recordset

    Set rstForLastRecordInsertedToDiscountLogsBackupTable = db.OpenRecordset("SELECT @@IDENTITY")
    VarlastDicountLogBackupID = rstForLastRecordInsertedToDiscountLogsBackupTable(0)
    rstDiscountLogsRecordsToInsertTobackupTable.Edit
    rstDiscountLogsRecordsToInsertTobackupTable!DiscountLogs_Backup_ID = VarlastDicountLogBackupID
    rstDiscountLogsRecordsToInsertTobackupTable.Update
    rstDiscountLogsRecordsToInsertTobackupTable.MoveNext
    Loop
  
  ' We close all objects
    AppendToDicountLogsBackUpTableQ.Close
    rstForLastRecordInsertedToDiscountLogsBackupTable.Close
    Set AppendToDicountLogsBackUpTableQ = Nothing
    Set rstForLastRecordInsertedToDiscountLogsBackupTable = Nothing
  
    rstDiscountLogsRecordsToInsertTobackupTable.Close
    db.Close
    Set rstDiscountLogsRecordsToInsertTobackupTable = Nothing
    Set db = Nothing

  End If
   
ExitProcedure:
Exit Sub
   
ErrorHandler:
Select Case Err.Number
        
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: CopyDiscountLogsToBackupTable" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
            
End Sub


Public Sub CopyProductDetailsToBackupTable(IssuedDocumentIDArg As Long)
Debug.Print "Module Public Functions - " & "CopyProductDetailsToBackupTable " & Time()
'On Error GoTo ErrorHandler

Dim db As DAO.Database
Dim rstProductDetailsRecordsToInsertTobackupTable As DAO.Recordset
Dim rstForLastRecordInsertedToDocumentProductDetailBackupTable As DAO.Recordset
Dim VarlastDocumentProductDetailsBackupID As Long
Dim AppendToProductDetailsBackUpTableQ As QueryDef
Dim VarCurrentUser As Integer
Dim VarCurrentTimestamp As Date

VarCurrentUser = DLookup("[CurrentUserT]![Current_User_ID]", "[CurrentUserT]", "[Current_User_ID]>0")
VarCurrentTimestamp = Now()

'We insert into back up table ALL the records (financial details)
'So, first we make a recordset with all the records that we are about to insert to the backup table
 Set db = CurrentDb
 Set rstProductDetailsRecordsToInsertTobackupTable = db.OpenRecordset("Select * from IssuedDocumentProductDetailsT WHERE Issued_Document_ID = " & IssuedDocumentIDArg & " And Is_New = False And IssuedDocumentProductDetails_Backup_ID Is Null")
   If Not rstProductDetailsRecordsToInsertTobackupTable.EOF Then
      rstProductDetailsRecordsToInsertTobackupTable.MoveLast
      rstProductDetailsRecordsToInsertTobackupTable.MoveFirst
'Then we itterate the recordset and insert its records one by one to the backup table
    Do Until rstProductDetailsRecordsToInsertTobackupTable.EOF
      Call CopyLinkAttributeValueToEntitiesToBackupTable(5, rstProductDetailsRecordsToInsertTobackupTable(0))
      Call CopyDiscountLogsToBackupTable(IssuedDocumentIDArg)
      Set AppendToProductDetailsBackUpTableQ = CurrentDb.QueryDefs("DocumentProductDetailsSaveBeforeEditQ")
      AppendToProductDetailsBackUpTableQ.Parameters(0) = VarCurrentTimestamp
      AppendToProductDetailsBackUpTableQ.Parameters(1) = VarCurrentUser
      AppendToProductDetailsBackUpTableQ.Parameters(2) = rstProductDetailsRecordsToInsertTobackupTable(0)
      'Debug.Print AppendToProductDetailsBackUpTableQ.Parameters(0)
      'Debug.Print AppendToProductDetailsBackUpTableQ.Parameters(1)
      'Debug.Print AppendToProductDetailsBackUpTableQ.Parameters(2)
      
      AppendToProductDetailsBackUpTableQ.Execute dbFailOnError
      
   'After each one insertion to the backup table, we take the ID of this insert to the back up table and we update the rstProductDetailsRecordsToInsertTobackupTable recordset (the basic table IssuedDocumentProductDetailsT)

    Set rstForLastRecordInsertedToDocumentProductDetailBackupTable = db.OpenRecordset("SELECT @@IDENTITY")
    VarlastDocumentProductDetailsBackupID = rstForLastRecordInsertedToDocumentProductDetailBackupTable(0)
    rstProductDetailsRecordsToInsertTobackupTable.Edit
    rstProductDetailsRecordsToInsertTobackupTable!IssuedDocumentProductDetails_Backup_ID = VarlastDocumentProductDetailsBackupID
    rstProductDetailsRecordsToInsertTobackupTable.Update
    rstProductDetailsRecordsToInsertTobackupTable.MoveNext
  Loop
  ' We close all objects
    AppendToProductDetailsBackUpTableQ.Close
  End If
   

ExitProcedure:
If Not rstProductDetailsRecordsToInsertTobackupTable Is Nothing Then
    rstProductDetailsRecordsToInsertTobackupTable.Close
    Set rstProductDetailsRecordsToInsertTobackupTable = Nothing
End If

If Not AppendToProductDetailsBackUpTableQ Is Nothing Then
    Set AppendToProductDetailsBackUpTableQ = Nothing
End If

If Not rstForLastRecordInsertedToDocumentProductDetailBackupTable Is Nothing Then
    rstForLastRecordInsertedToDocumentProductDetailBackupTable.Close
    Set rstForLastRecordInsertedToDocumentProductDetailBackupTable = Nothing
End If

If Not db Is Nothing Then
    db.Close
    Set db = Nothing
End If

Exit Sub
   
ErrorHandler:
Select Case Err.Number
        Case 3314
        MsgBox "You have left empty fields which must be filled", vbInformation, "������ ���������"
        Response = acDataErrContinue
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: CopyProductDetailsToBackupTable" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
End Sub

Public Sub CopyDiscountLogsDetailsToBackupTable(DiscountLogIDArg As Integer)
Debug.Print "Module Public Functions - " & "CopyDiscountLogsDetailsToBackupTable " & Time()
'On Error GoTo ErrorHandler

Dim db As DAO.Database
Dim rstDiscountLogsDetailsRecordsToInsertTobackupTable As DAO.Recordset
Dim rstForLastRecordInsertedToDiscountLogsDetailsBackupTable As DAO.Recordset
Dim VarlastDiscountLogsDetailsBackupID As Long
Dim AppendToDiscountLogsDetailsBackUpTableQ As QueryDef
Dim VarCurrentUser As Integer
Dim VarCurrentTimestamp As Date

VarCurrentUser = DLookup("[CurrentUserT]![Current_User_ID]", "[CurrentUserT]", "[Current_User_ID]>0")
VarCurrentTimestamp = Now()

'We insert into back up table ALL the records (financial details)
'So, first we make a recordset with all the records that we are about to insert to the backup table
 Set db = CurrentDb
 Set rstDiscountLogsDetailsRecordsToInsertTobackupTable = db.OpenRecordset("Select * from DiscountLogsDetailsT WHERE Discount_Logs_ID = " & DiscountLogIDArg & " AND is_new = false and DiscountLogsDetails_Backup_ID is null")
   If Not rstDiscountLogsDetailsRecordsToInsertTobackupTable.EOF Then
      rstDiscountLogsDetailsRecordsToInsertTobackupTable.MoveLast
      rstDiscountLogsDetailsRecordsToInsertTobackupTable.MoveFirst
'Then we itterate the recordset and insert its records one by one to the backup table
    Do Until rstDiscountLogsDetailsRecordsToInsertTobackupTable.EOF
      Set AppendToDiscountLogsDetailsBackUpTableQ = CurrentDb.QueryDefs("DiscountLogsDetailsSaveBeforeEditQ")
      AppendToDiscountLogsDetailsBackUpTableQ.Parameters(0) = VarCurrentTimestamp
      AppendToDiscountLogsDetailsBackUpTableQ.Parameters(1) = VarCurrentUser
      AppendToDiscountLogsDetailsBackUpTableQ.Parameters(2) = rstDiscountLogsDetailsRecordsToInsertTobackupTable(0)
      AppendToDiscountLogsDetailsBackUpTableQ.Execute dbFailOnError
      
   'After each one insertion to the backup table, we take the ID of this insert to the back up table and we update the rstDiscountLogsDetailsRecordsToInsertTobackupTable recordset (the basic table IssuedDocumentFinancialDetailsT)

    Set rstForLastRecordInsertedToDiscountLogsDetailsBackupTable = db.OpenRecordset("SELECT @@IDENTITY")
    VarlastDiscountLogsDetailsBackupID = rstForLastRecordInsertedToDiscountLogsDetailsBackupTable(0)
    rstDiscountLogsDetailsRecordsToInsertTobackupTable.Edit
    rstDiscountLogsDetailsRecordsToInsertTobackupTable!DiscountLogsDetails_Backup_ID = VarlastDiscountLogsDetailsBackupID
    rstDiscountLogsDetailsRecordsToInsertTobackupTable.Update
    rstDiscountLogsDetailsRecordsToInsertTobackupTable.MoveNext
  Loop
  ' We close all objects
    AppendToDiscountLogsDetailsBackUpTableQ.Close
    rstForLastRecordInsertedToDiscountLogsDetailsBackupTable.Close
    Set AppendToDiscountLogsDetailsBackUpTableQ = Nothing
    Set rstForLastRecordInsertedToDiscountLogsDetailsBackupTable = Nothing
  
    rstDiscountLogsDetailsRecordsToInsertTobackupTable.Close
    db.Close
    Set rstDiscountLogsDetailsRecordsToInsertTobackupTable = Nothing
    Set db = Nothing

  End If
  
ExitProcedure:
If Not rstDiscountLogsDetailsRecordsToInsertTobackupTable Is Nothing Then
    rstDiscountLogsDetailsRecordsToInsertTobackupTable.Close
    Set rstDiscountLogsDetailsRecordsToInsertTobackupTable = Nothing
End If
    
If Not rstForLastRecordInsertedToDiscountLogsDetailsBackupTable Is Nothing Then
    rstForLastRecordInsertedToDiscountLogsDetailsBackupTable.Close
    Set rstForLastRecordInsertedToDiscountLogsDetailsBackupTable = Nothing
End If

If Not AppendToDiscountLogsDetailsBackUpTableQ Is Nothing Then
    Set AppendToDiscountLogsDetailsBackUpTableQ = Nothing
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
            "Error Source: CopyDiscountLogsDetailsToBackupTable" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
  
End Sub

Public Function CheckIfControlHasFocus(thisControl As Control, frm As Access.Form) As Boolean
 ' On Error GoTo Err_Handler
Debug.Print "Module Public Functions - " & "CheckIfControlHasFocus " & Time()
        If Not frm.ActiveControl Is Nothing Then
            If (frm.ActiveControl Is thisControl) Then
                CheckIfControlHasFocus = True
            Else
                CheckIfControlHasFocus = False
            End If
        Else
            GoTo Err_Handler
        End If
  Debug.Print thisControl.Name & " = " & CheckIfControlHasFocus
close_function:
    On Error GoTo 0
    Exit Function
        
Err_Handler:
        CheckIfControlHasFocus = False
        Debug.Print thisControl.Name & " = " & CheckIfControlHasFocus
        Resume close_function

End Function


Public Function FindRelatedEntityIDDetails(EntityTypeID As Integer, EntityID As Integer, SpecificField As String)
Debug.Print "Module Public Functions - " & "FindRelatedEntityIDDetails " & Time()
'On Error GoTo ErrorHandler
'this function brings the value from the lookup field (given from "SpecificField" argument) that corresponds to the values _
in the fields "EntityTypeID_For_RelevantTablePKField" as EntityTypeID argument and "Attribute_Value_Number" as EntityID argument that are stored in LinkAttributeValueToEntitiesT.

Select Case EntityTypeID
Case 1 'product
FindRelatedEntityIDDetails = DLookup(SpecificField, "ProductsT", "ProductsT.Product_ID = " & EntityID)
Case 2 'transactor
FindRelatedEntityIDDetails = DLookup(SpecificField, "TransactorsWithBasicTransactorsDescriptionQ", "TransactorsWithBasicTransactorsDescriptionQ.Transactor_ID = " & EntityID)
Case 3 'issued document
FindRelatedEntityIDDetails = DLookup(SpecificField, "IssuedDocumentT", "IssuedDocumentT.Issued_Document_ID = " & EntityID)
Case 4 'document financial details
FindRelatedEntityIDDetails = DLookup(SpecificField, "IssuedDocumentFinancialDetailsT", "IssuedDocumentFinancialDetailsT.Issued_Document_Financial_Details_ID = " & EntityID)
Case 5 'document product details
FindRelatedEntityIDDetails = DLookup(SpecificField, "IssuedDocumentProductDetailsT", "IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID = " & EntityID)
Case 6 'transaction
FindRelatedEntityIDDetails = DLookup(SpecificField, "TransactionsT", "TransactionsT.Transaction_ID = " & EntityID)
Case 8 'action
FindRelatedEntityIDDetails = DLookup(SpecificField, "ActionsT", "ActionsT.Action_ID = " & EntityID)
Case 9 'protocol
FindRelatedEntityIDDetails = DLookup(SpecificField, "ProtocolsT", "ProtocolsT.Protocol_ID = " & EntityID)
Case 10 'installation
FindRelatedEntityIDDetails = DLookup(SpecificField, "InstallationsT", "InstallationsT.Installation_ID = " & EntityID)
End Select


ExitProcedure:

Exit Function
   
ErrorHandler:
Select Case Err.Number
        
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: FindRelatedEntityIDDetails" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
  
  
End Function

Public Function Orphaned_LinkAttributeValuesToEntityID_Records(EntityTypeIDArg As Integer, ActionArg As Integer) 'Action 1 = delete orphans, Action 2 = count Orphans
Debug.Print "Module Public Functions - " & "Orphaned_LinkAttributeValuesToEntityID_Records " & Time()
'On Error GoTo ErrorHandler

Dim db As DAO.Database
Dim strrstOrphaned_LinkAttributeValuesToEntityID_Records As String
Dim rstOrphaned_LinkAttributeValuesToEntityID_Records As DAO.Recordset

Set db = CurrentDb

Select Case EntityTypeIDArg
Case 1 'product
strrstOrphaned_LinkAttributeValuesToEntityID_Records = "SELECT DISTINCT LinkAttributeValueToEntitiesT.* " & _
"FROM ProductsT RIGHT JOIN LinkAttributeValueToEntitiesT ON ProductsT.Product_ID = LinkAttributeValueToEntitiesT.Entity_ID " & _
"WHERE (((LinkAttributeValueToEntitiesT.Entity_Type_ID)=1) AND ((ProductsT.Product_ID) Is Null));"
Case 2 'transactor
strrstOrphaned_LinkAttributeValuesToEntityID_Records = "SELECT DISTINCT LinkAttributeValueToEntitiesT.* " & _
"FROM TransactorsT RIGHT JOIN LinkAttributeValueToEntitiesT ON TransactorsT.Transactor_ID = LinkAttributeValueToEntitiesT.Entity_ID " & _
"WHERE (((LinkAttributeValueToEntitiesT.Entity_Type_ID)=2) AND ((TransactorsT.Transactor_ID) Is Null));"
Case 3 'issued document
strrstOrphaned_LinkAttributeValuesToEntityID_Records = "SELECT DISTINCT LinkAttributeValueToEntitiesT.* " & _
"FROM IssuedDocumentT RIGHT JOIN LinkAttributeValueToEntitiesT ON IssuedDocumentT.Issued_Document_ID = LinkAttributeValueToEntitiesT.Entity_ID " & _
"WHERE (((LinkAttributeValueToEntitiesT.Entity_Type_ID)=3) AND ((IssuedDocumentT.Issued_Document_ID) Is Null));"
Case 4 'document financial details
strrstOrphaned_LinkAttributeValuesToEntityID_Records = "SELECT DISTINCT LinkAttributeValueToEntitiesT.* " & _
"FROM IssuedDocumentFinancialDetailsT RIGHT JOIN LinkAttributeValueToEntitiesT ON IssuedDocumentFinancialDetailsT.Issued_Document_Financial_Details_ID = LinkAttributeValueToEntitiesT.Entity_ID " & _
"WHERE (((LinkAttributeValueToEntitiesT.Entity_Type_ID)=4) AND ((IssuedDocumentFinancialDetailsT.Issued_Document_Financial_Details_ID) Is Null));"
Case 5 'document product details
strrstOrphaned_LinkAttributeValuesToEntityID_Records = "SELECT DISTINCT LinkAttributeValueToEntitiesT.* " & _
"FROM IssuedDocumentProductDetailsT RIGHT JOIN LinkAttributeValueToEntitiesT ON IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID = LinkAttributeValueToEntitiesT.Entity_ID " & _
"WHERE (((LinkAttributeValueToEntitiesT.Entity_Type_ID)=5) AND ((IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID) Is Null));"
Case 6 'transaction
strrstOrphaned_LinkAttributeValuesToEntityID_Records = "SELECT DISTINCT LinkAttributeValueToEntitiesT.* " & _
"FROM TransactionsT RIGHT JOIN LinkAttributeValueToEntitiesT ON TransactionsT.Transaction_ID = LinkAttributeValueToEntitiesT.Entity_ID " & _
"WHERE (((LinkAttributeValueToEntitiesT.Entity_Type_ID)=7) AND ((TransactionsT.Transaction_ID) Is Null));"
Case 7 'Subattributes
strrstOrphaned_LinkAttributeValuesToEntityID_Records = "SELECT LinkSubAttributeValueToAttributeValuesQ.Entity_ID " & _
"FROM (SELECT LinkAttributeValueToEntitiesT.*, LinkAttributeValueToEntitiesT.Entity_Type_ID " & _
"FROM LinkAttributeValueToEntitiesT " & _
"WHERE (((LinkAttributeValueToEntitiesT.Entity_Type_ID)=7)))  AS LinkSubAttributeValueToAttributeValuesQ LEFT JOIN LinkAttributeValueToEntitiesT " & _
"ON LinkSubAttributeValueToAttributeValuesQ.Entity_ID = LinkAttributeValueToEntitiesT.Link_Attribute_Value_To_Entity_ID " & _
"WHERE (((LinkAttributeValueToEntitiesT.Entity_ID) Is Null));"
Case 8 'action
strrstOrphaned_LinkAttributeValuesToEntityID_Records = "SELECT DISTINCT LinkAttributeValueToEntitiesT.* " & _
"FROM ActionsT RIGHT JOIN LinkAttributeValueToEntitiesT ON ActionsT.Action_ID = LinkAttributeValueToEntitiesT.Entity_ID " & _
"WHERE (((LinkAttributeValueToEntitiesT.Entity_Type_ID)=8) AND ((ActionsT.Action_ID) Is Null));"
Case 9 'protocol
strrstOrphaned_LinkAttributeValuesToEntityID_Records = "SELECT DISTINCT LinkAttributeValueToEntitiesT.* " & _
"FROM ProtocolsT RIGHT JOIN LinkAttributeValueToEntitiesT ON ProtocolsT.Protocol_ID = LinkAttributeValueToEntitiesT.Entity_ID " & _
"WHERE (((LinkAttributeValueToEntitiesT.Entity_Type_ID)=9) AND ((ProtocolsT.Protocol_ID) Is Null));"
Case 10 'installation
strrstOrphaned_LinkAttributeValuesToEntityID_Records = "SELECT DISTINCT LinkAttributeValueToEntitiesT.* " & _
"FROM InstallationsT RIGHT JOIN LinkAttributeValueToEntitiesT ON InstallationsT.Installation_ID = LinkAttributeValueToEntitiesT.Entity_ID " & _
"WHERE (((LinkAttributeValueToEntitiesT.Entity_Type_ID)=10) AND ((InstallationsT.Installation_ID) Is Null));"
End Select

Set rstOrphaned_LinkAttributeValuesToEntityID_Records = db.OpenRecordset(strrstOrphaned_LinkAttributeValuesToEntityID_Records, , dbFailOnError)

rstOrphaned_LinkAttributeValuesToEntityID_Records.MoveLast
rstOrphaned_LinkAttributeValuesToEntityID_Records.MoveFirst

If Not rstOrphaned_LinkAttributeValuesToEntityID_Records.EOF Then

  Select Case ActionArg
  Case 1
  Orphaned_LinkAttributeValuesToEntityID_Records = rstOrphaned_LinkAttributeValuesToEntityID_Records.recordcount
  GoTo ExitProcedure

  Case 2
  Do Until rstOrphaned_LinkAttributeValuesToEntityID_Records.EOF
  Call Delete_One_LinkAttributeValuesToEntities(rstOrphaned_LinkAttributeValuesToEntityID_Records(1), EntityTypeIDArg, False)
  rstOrphaned_LinkAttributeValuesToEntityID_Records.MoveNext
  Loop
  
  End Select
End If


ExitProcedure:
If Not rstOrphaned_LinkAttributeValuesToEntityID_Records Is Nothing Then
    rstOrphaned_LinkAttributeValuesToEntityID_Records.Close
    Set rstOrphaned_LinkAttributeValuesToEntityID_Records = Nothing
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
            "Error Source: Orphaned_LinkAttributeValuesToEntityID_Records" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
  

End Function

Public Function RelatedEntityDescription(AttributeIDArg As Integer, EntityIDArg As Integer, EntityTypeIDArg As Integer) As String
Debug.Print "Module Public Functions - " & "RelatedEntityDescription " & Time()
'On Error GoTo ErrorHandler

Dim query_string As String
If IsNull(EntityTypeIDArg) Or IsNull(EntityIDArg) Then
RelatedEntityDescription = ""
Debug.Print "NULL ARGUMENTS"
Exit Function
End If

Dim db As DAO.Database
Dim rst As DAO.Recordset

Set db = CurrentDb
query_string = "Select EntityTypeID_For_RelevantTablePKField, Attribute_Value_number from LinkAttrValToEntWithAttrDescAndEntityTypeIDFKQ " & _
"WHERE LinkAttrValToEntWithAttrDescAndEntityTypeIDFKQ.Attribute_ID = " & AttributeIDArg & " AND LinkAttrValToEntWithAttrDescAndEntityTypeIDFKQ.Entity_ID = " & EntityIDArg & _
" AND LinkAttrValToEntWithAttrDescAndEntityTypeIDFKQ.Entity_Type_ID = " & EntityTypeIDArg

Debug.Print query_string

Set rst = db.OpenRecordset(query_string)

If rst.EOF Then
RelatedEntityDescription = ""
Exit Function
Else
rst.MoveLast
rst.MoveFirst

RelatedEntityDescription = "" ' initialization of the value

Do Until rst.EOF
'Debug.Print "rst(0) = " & rst(0) & " - rst(1) = " & rst(1)
RelatedEntityDescription = RelatedEntityDescription & IIf(RelatedEntityDescription = "", "", "," & vbCrLf) & CStr(SingleRelatedEntityDescription(rst(0), rst(1)))
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

Exit Function
   
ErrorHandler:
Select Case Err.Number
        
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: RelatedEntityDescription" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
  
End Function

Public Function SingleRelatedEntityDescription(RelevantEntityTypeIDarg As Integer, EntityIDArg As Integer) As String
Debug.Print "Module Public Functions - " & "SingleRelatedEntityDescription " & Time()

If IsNull(RelevantEntityTypeIDarg) Or IsNull(EntityIDArg) Or EntityIDArg = 0 Then
SingleRelatedEntityDescription = ""
Exit Function
End If

SingleRelatedEntityDescription = "" ' initialization of the value

Select Case RelevantEntityTypeIDarg

Case 1 'Products
SingleRelatedEntityDescription = Nz(DLookup("ProductsT.Product_Description", "ProductsT", "ProductsT.Product_ID = " & EntityIDArg), "")

Case 2 'transactors
SingleRelatedEntityDescription = Nz(DLookup("TransactorsWithBasicTransactorsDescriptionQ.Basic_Transactor_Description", "TransactorsWithBasicTransactorsDescriptionQ", "TransactorsWithBasicTransactorsDescriptionQ.Transactor_ID = " & EntityIDArg), "")

Case 7 'Countries
SingleRelatedEntityDescription = Nz(DLookup("CountriesT.Country_Description", "CountriesT", "CountriesT.Country_ID = " & EntityIDArg), "")

End Select
End Function

Public Function SingleRelatedEntityTypeDescription(RelevantEntityTypeIDarg As Integer, EntityIDArg As Integer) As String
Debug.Print "Module Public Functions - " & "SingleRelatedEntityTypeDescription " & Time()
If IsNull(RelevantEntityTypeIDarg) Or IsNull(EntityIDArg) Or EntityIDArg = 0 Then
SingleRelatedEntityTypeDescription = ""
Exit Function
End If

Dim TypeIDForRelatedEntityType As Long

SingleRelatedEntityTypeDescription = "" ' initialization of the value

Select Case RelevantEntityTypeIDarg

Case 1 'Products
TypeIDForRelatedEntityType = DLookup("ProductsT.Product_Type_ID", "ProductsT", "ProductsT.Product_ID = " & EntityIDArg)
SingleRelatedEntityTypeDescription = DLookup("ProductTypeT.Product_Type_Description", "Product_Type", "ProductTypeT.Product_Type_ID = " & TypeIDForRelatedEntityType)

Case 2 'transactors
TypeIDForRelatedEntityType = DLookup("TransactorsT.Transactor_Type_ID", "TransactorsT", "TransactorsT.Transactor_ID = " & EntityIDArg)
SingleRelatedEntityTypeDescription = DLookup("TransactorsWithBasicTransactorsDescriptionQ.Transactor_Type_Desription", "TransactorsWithBasicTransactorsDescriptionQ", "TransactorsWithBasicTransactorsDescriptionQ.Transactor_Type_ID = " & TypeIDForRelatedEntityType)
End Select

End Function

Public Sub ClearListbox(lstbox As ListBox)
Debug.Print "Module Public Functions - " & "ClearListbox " & Time()
    Dim varlistbox As ListBox
    Dim varItm As Variant
    
             With lstbox
        For Each varItm In .ItemsSelected
            .Selected(varItm) = False
        Next varItm

    End With
End Sub

Public Function RowsourceForCboTosearchForTransactorsWithSimilarInfo(AttributeIDArg As Long, AttributeValueArg As Variant) As String
Debug.Print "Module Public Functions - " & "FindTransactorIDWithSimilarInfo " & Time()

FindTransactorWithSimilarInfo = 0

Select Case AttributeIDArg
    Case 306
    RowsourceForCboTosearchForTransactorsWithSimilarInfo = "Select TransactorsWithBasicTransactorsDescriptionQ.Basic_Transactor_Description, TransactorsWithBasicTransactorsDescriptionQ.Transactor_ID, " & _
    "TransactorsWithBasicTransactorsDescriptionQ.Transactor_Type_Desription " & _
    "from TransactorsWithBasicTransactorsDescriptionQ inner join TransactorContactDetailsT on TransactorsWithBasicTransactorsDescriptionQ.Basic_Transactor_ID = TransactorContactDetailsT.Basic_Tsansactor_ID_FK " & _
    "Where (Transactor_Contact_Type_ID = 1 OR Transactor_Contact_Type_ID = 2) " & _
    "AND Transactor_Contact_Decription Like ""*" & AttributeValueArg & "*"" order by TransactorsWithBasicTransactorsDescriptionQ.Basic_Transactor_Description "

    Case 307
    RowsourceForCboTosearchForTransactorsWithSimilarInfo = "Select TransactorsWithBasicTransactorsDescriptionQ.Basic_Transactor_Description, TransactorsWithBasicTransactorsDescriptionQ.Transactor_ID, " & _
    "TransactorsWithBasicTransactorsDescriptionQ.Transactor_Type_Desription " & _
    "from TransactorsWithBasicTransactorsDescriptionQ inner join TransactorContactDetailsT on TransactorsWithBasicTransactorsDescriptionQ.Basic_Transactor_ID = TransactorContactDetailsT.Basic_Tsansactor_ID_FK " & _
    "Where TransactorContactDetailsT.Transactor_Contact_Type_ID = 2 " & _
    "AND Transactor_Contact_Decription Like ""*" & AttributeValueArg & "*"" order by TransactorsWithBasicTransactorsDescriptionQ.Basic_Transactor_Description "

End Select

End Function

Public Function BringAnyAttributeValue(AttributeIDArg As Long, EntityTypeIDArg As Integer, EntityIDArg As Long) As Variant

Debug.Print "Module Public Functions - " & "BringAnyAttributeValue " & Time()

Dim db As DAO.Database
Dim rst As DAO.Recordset

Set db = CurrentDb
Set rst = db.OpenRecordset("Select CStr(Nz([Attribute_Value_String],"""")) & CStr(Nz([Attribute_Value_Number],"""")) & CStr(IIf([Attribute_Value_Boolean]=0,""No"",IIf([Attribute_Value_Boolean]=-1,""Yes"",""""))) & CStr(Nz([Attribute_Value_Date],"""")) & CStr(Nz([Attribute_Value_Time],"""")) & CStr(Nz([Attribute_Value_TImestamp],"""")) AS AttrVALUE " & _
"from LinkAttributeValueToEntitiesT where Attribute_ID = " & AttributeIDArg & " AND Entity_Type_ID = " & EntityTypeIDArg & " AND Entity_ID = " & EntityIDArg)

If Not rst.EOF Then

rst.MoveLast
rst.MoveFirst

Do Until rst.EOF
BringAnyAttributeValue = ", " & rst(0).Value
rst.MoveNext
Loop

If Len(BringAnyAttributeValue) > 2 Then
        BringAnyAttributeValue = Mid(BringAnyAttributeValue, 3)
End If
Else
BringAnyAttributeValue = ""
End If

rst.Close
Set rst = Nothing
db.Close
Set db = Nothing

End Function

Public Sub TotalUnitDiscountAmount(DocumentIDArg As Long) 'it affects ProductDetailsT. It updates all documentProductDetailsIDs with the correct "Total # Unit Discount" field value
Debug.Print "Module Public Functions - " & "TotalUnitDiscountAmount " & Time(); ""
'On Error GoTo ErrorHandler

Dim db As DAO.Database
Dim rst As DAO.Recordset
Dim HasBeenBackedUpVar As Boolean
Dim NewTotalUnitDiscountVar As Double
Dim NewUnitPriceAfterDiscount As Double

HasBackedUpVar = False

Set db = CurrentDb
Set rst = db.OpenRecordset("Select * from IssuedDocumentProductDetailsT where Is_Deleted = false AND Issued_Document_ID = " & DocumentIDArg)

If Not rst.EOF Then

rst.MoveFirst

Do Until rst.EOF
NewUnitPriceAfterDiscount = CalculateUnitPriceAfterDiscount(rst("Issued_Document_Product_Details_ID"), rst("Unit_Price_Before_Discount"))
NewTotalUnitDiscountVar = rst("Unit_Price_Before_Discount") - NewUnitPriceAfterDiscount
If (rst("Total # Unit Discount") = NewTotalUnitDiscountVar) And (rst("Unit_Price_After_Discount") = NewUnitPriceAfterDiscount) Then
 rst.MoveNext
Else
   If HasBeenBackedUpVar = False Then
   Call BackupFullTransaction(, DocumentIDArg)
   HasBeenBackedUpVar = True
   End If
 rst.Edit
 rst("Total # Unit Discount") = NewTotalUnitDiscountVar
 rst("Unit_Price_After_Discount") = NewUnitPriceAfterDiscount
 
 rst.Update
 rst.MoveNext
End If
Loop
End If

rst.Close
db.Close

Set rst = Nothing
Set db = Nothing

ExitProcedure:
Exit Sub
   
ErrorHandler:
Select Case Err.Number
       
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: TotalUnitDiscountAmount" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
End Sub
Public Function CalculateUnitPriceAfterDiscount(ByVal ProductDetailsIDArg As Long, ByVal InitialUnitPriceBeforeDiscountArg As Double) As Currency  'it affects DiscountLogsDetailsT
Debug.Print "Module Public Functions - " & "CalculateUnitPriceAfterDiscount " & Time()
'On Error GoTo ErrorHandler

Dim db As DAO.Database
Dim rstDiscountsForProduct As DAO.Recordset

Set db = CurrentDb
Set rstDiscountsForProduct = db.OpenRecordset("Select * from DiscountLogsDetailsT inner join DiscountLogsT on DiscountLogsDetailsT.Discount_Logs_ID = DiscountLogsT.Discount_Logs_ID " & _
"where DiscountLogsDetailsT.Product_Details_ID = " & ProductDetailsIDArg & " And DiscountLogsT.Is_Deleted = False " & _
" AND DiscountLogsDetailsT.Is_Deleted = false order by Discounts_Logs_Details_ID desc") '

If Not rstDiscountsForProduct.EOF Then
rstDiscountsForProduct.MoveLast
rstDiscountsForProduct.MoveFirst
CalculateUnitPriceAfterDiscount = rstDiscountsForProduct("Unit_Price_After_This_Discount")
Else
CalculateUnitPriceAfterDiscount = InitialUnitPriceBeforeDiscountArg
End If

rstDiscountsForProduct.Close
db.Close

Set rstDiscountsForProduct = Nothing
Set db = Nothing

ExitProcedure:
Exit Function
   
ErrorHandler:
Select Case Err.Number
       Case 7878
       Resume Next
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: CalculateUnitPriceAfterDiscount" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select

End Function
Public Function NewOrEditedDiscountRecords(FormArg As Form, IsNewArg As Integer, BackUpIDArg As Integer) As Boolean
Debug.Print "Module Public Functions - " & "NewOrEditedDiscountRecords " & Time(); ""
'On Error GoTo ErrorHandler

Dim db As DAO.Database
Dim rst As DAO.Recordset

Set db = CurrentDb
Set rst = FormArg.RecordsetClone
If Not rst.EOF Then
rst.MoveLast
rst.MoveFirst


NewOrEditedDiscountRecords = False
'Dim i As Long

Do Until rst.EOF

    'For i = 0 To rst.Fields.Count - 1
       ' Debug.Print rst.Fields(i).Name,
   'Next
'Debug.Print "IsNewArg = " & IsNewArg
'Debug.Print "BackUpIDArg = " & BackUpIDArg
'Debug.Print "rst(IsNewArg) = " & rst(IsNewArg)
'Debug.Print "rst(BackUpIDArg) = " & rst(BackUpIDArg)
If rst(IsNewArg) = True Or Not IsNull(rst(BackUpIDArg)) Then
  NewOrEditedDiscountRecords = True
  Exit Do
  Else
  rst.MoveNext
End If
Loop

End If

rst.Close
db.Close
Set rst = Nothing
Set db = Nothing

ExitProcedure:
Exit Function
   
ErrorHandler:
Select Case Err.Number
     Case 3167
       Resume Next
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: NewOrEditedDiscountRecords" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
End Function

Public Function UnitPriceBeforeThisDiscount(ThisDiscountIDArg As Long, ProductDetailsIDArg As Long)
Debug.Print "Module Public Functions - " & "UnitPriceBeforeThisDiscount " & Time(); ""
'On Error GoTo ErrorHandler

Dim db As DAO.Database
Dim rst As DAO.Recordset

Set db = CurrentDb

Set rst = db.OpenRecordset("Select DiscountLogsDetailsT.Unit_Price_After_This_Discount from DiscountLogsDetailsT  inner join DiscountLogsT on DiscountLogsDetailsT.Discount_Logs_ID = DiscountLogsT.Discount_Logs_ID  " & _
"where DiscountLogsDetailsT.Discount_Logs_ID < " & ThisDiscountIDArg & " AND Product_Details_ID = " & ProductDetailsIDArg & " And DiscountLogsT.Is_Deleted = False And DiscountLogsDetailsT.Is_Deleted = False")

If Not rst.EOF Then
rst.Move first
rst.MoveLast
UnitPriceBeforeThisDiscount = rst(0)
'Debug.Print "UnitPriceBeforeThisDiscount = " & UnitPriceBeforeThisDiscount
Else
UnitPriceBeforeThisDiscount = DLookup("IssuedDocumentProductDetailsT.Unit_Price_Before_Discount", "IssuedDocumentProductDetailsT", "IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID = " & ProductDetailsIDArg)
End If

ExitProcedure:
Exit Function
   
ErrorHandler:
Select Case Err.Number
       
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: UnitPriceBeforeThisDiscount" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select

End Function

Public Sub RecalculateDiscountsNew(ArgDocumentID As Long)
Debug.Print "Module Public Functions - " & "RecalculateDiscountsNew " & Time()
'On Error GoTo Error_Handler

Dim dbRecalculateDiscounts As DAO.Database
Dim rstDiscountLogsIDs As DAO.Recordset
Dim rstProductsAffected As DAO.Recordset
Dim Response As Boolean
Dim UnitPriceBeforeThisDiscountVar As Double
Dim varTotalPriceOfAffectedProductsBeforeThisDiscount  As Double
Dim DiscountTypeIsPercentageVar As Boolean
Dim DiscountPercentageOrAmountVar As Double
Dim varDiscountOrOfferOnSingleProductOnly As Boolean
Dim DiscountAmountPerUnitVar As Double
Dim DiscountLogIDVar As Long

Set dbRecalculateDiscounts = CurrentDb

Set rstDiscountLogsIDs = dbRecalculateDiscounts.OpenRecordset("Select *  " & _
"from DiscountLogsT " & _
"where Issued_Document_ID = " & ArgDocumentID & " AND Is_Deleted = false " & _
"order by Discount_Logs_ID")

If Not (rstDiscountLogsIDs.BOF And rstDiscountLogsIDs.EOF) Then
rstDiscountLogsIDs.MoveLast
rstDiscountLogsIDs.MoveFirst
Else
Call TotalUnitDiscountAmount(ArgDocumentID)
MsgBox "No discounts found for recaclulation."
GoTo ExitProcedure
End If

' ReIterationForDiscountUpdate(DiscountLogsIDArg As Long, IsPercentageDiscountArg As Boolean, DiscountValueArg As Double, DiscountOnSingleProductOnlyArg As Boolean)

Do Until rstDiscountLogsIDs.EOF
'UnitPriceBeforeThisDiscountVar = UnitPriceBeforeThisDiscount(rstDiscountLogsIDs("Discount_Logs_ID"), rstDiscountLogsIDs("Product_Details_ID"))
'varDiscountOrOfferOnSingleProductOnly = rstDiscountLogsIDs("Discount_For_Single_ProductDetailsID_Only")
If Nz(rstDiscountLogsIDs("[%Discount_Percentage]"), 0) = 0 Then
DiscountTypeIsPercentageVar = False
Else
DiscountTypeIsPercentageVar = True
End If

DiscountLogIDVar = rstDiscountLogsIDs("Discount_Logs_ID")

DiscountPercentageOrAmountVar = Round(Nz(rstDiscountLogsIDs("[%Discount_Percentage]"), 0) + Nz(rstDiscountLogsIDs("[#Discount_Value]"), 0), 2)

Call Form_DiscountLogsF.ReIterationForDiscountUpdate(rstDiscountLogsIDs("Discount_Logs_ID"), DiscountTypeIsPercentageVar, DiscountPercentageOrAmountVar, rstDiscountLogsIDs("Discount_For_Single_ProductDetailsID_Only"))
                                                                       
rstDiscountLogsIDs.MoveNext
Loop

rstDiscountLogsIDs.Close
Set rstDiscountLogsIDs = Nothing

If Not rstProductsAffected Is Nothing Then
rstProductsAffected.Close
Set rstProductsAffected = Nothing
End If

dbRecalculateDiscounts.Close
Set dbRecalculateDiscounts = Nothing

Call TotalUnitDiscountAmount(ArgDocumentID)

ExitProcedure:
Exit Sub
   
Error_handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: RecalculateDiscountsNew" & vbCrLf & _
           "Error Description: " & Err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    Resume ExitProcedure

End Sub


Public Function CheckIfFormIsOpen(FormNameArg As String) As Boolean
    Dim frm As Object
    For Each frm In Forms
        If frm.Name = FormNameArg Then
            CheckIfFormIsOpen = True
            Exit Function
        End If
    Next frm
    CheckIfFormIsOpen = False
End Function

Public Sub InitiallizeAllFormsCollections()
Debug.Print "Module Public Functions - " & "InitiallizeAllFormsCollections " & Time()
On Error GoTo ErrorHandler

 Set TransactorsFormCollection = New Collection


ExitProcedure:
Exit Sub
   
ErrorHandler:
Select Case Err.Number
       
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: InitiallizeAllFormsCollections" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
End Sub

Public Function NewTransactorADD(Optional TransactorTypeIDArg As Long, Optional BasicTransactorIDArg As Long, Optional TextForNewTransactorDescriptionArg As String)
Debug.Print "Module Public Functions - " & "NewTransactorADD " & Time()
On Error GoTo ErrorHandler

DoCmd.OpenForm "TransactorsAddF", acDialog, , , acFormAdd
TransactorsAddF.SetFocus

If Not IsNull(TransactorTypeIDArg) Then
TransactorTypeIDCbo.SetFocus
TransactorTypeIDCbo = TransactorTypeIDArg
End If

If Not IsNull(BasicTransactorIDArg) Then
BasicTransactorIDTbox.SetFocus
BasicTransactorIDTbox = BasicTransactorIDArg
End If

If Not IsNull(TextForNewTransactorDescriptionArg) Then
BasicTransactorIDCbo1.SetFocus
BasicTransactorIDCbo1.Text = TextForNewTransactorDescriptionArg
End If


ExitProcedure:
Exit Function
   
ErrorHandler:
Select Case Err.Number
       
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: NewTransactorADD " & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select

End Function

Function FetchLastInsertedTransactionID() As Long
Debug.Print "Module Public Functions - " & "FetchLastInsertedTransactionID " & Time()
On Error GoTo ErrorHandler

Dim db As DAO.Database
Dim rst As DAO.Recordset

Set db = CurrentDb
Set rst = db.OpenRecordset("Select max(Transaction_ID) from TransactionsT")

If rst.recordcount = 1 Then
FetchLastInsertedTransactionID = rst(0)
End If

ExitProcedure:
Exit Function
   
ErrorHandler:
Select Case Err.Number
       
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: FetchLastInsertedTransactionID " & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select

End Function
Public Function UpdateTransactorsTotalDebitAndTotalCreditByDocumentFinancialDetailsRecordset(DocumentFinancialDetailsRecordsetArg As Recordset, IntentionIDArg As Integer, Optional DocIsDeletedArg As Boolean) As Boolean
Debug.Print "Module Public Functions - " & "UpdateTransactorsTotalDebitAndTotalCreditByDocumentFinancialDetailsRecordset " & Time()
'On Error GoTo ErrorHandler

Dim VarFinTransactorDebitDifferenceToAddToTransactorDebitTotal As Double
Dim VarFinTransactorCreditDifferenceToAddToTransactorCreditTotal As Double
Dim VarNewFinancialTransactorID As Long
Dim VarOldFinancialTransactorID As Long
Dim VarNewDebit As Double
Dim VarOldDebit As Double
Dim VarNewCredit As Double
Dim VarOldCredit As Double
Dim VarIntentionAffectsFinancial As Boolean
Dim db As DAO.Database
Dim Vartest As Double
Set db = CurrentDb

VarIntentionAffectsFinancial = DLookup("IntentionsT.Affects_Financial", "IntentionsT", "Intention_ID = " & IntentionIDArg)

UpdateTransactorsTotalDebitAndTotalCreditByDocumentFinancialDetailsRecordset = False

If DocumentFinancialDetailsRecordsetArg.recordcount > 0 Then
   DocumentFinancialDetailsRecordsetArg.MoveLast
   DocumentFinancialDetailsRecordsetArg.MoveFirst
   IterateRecordsets DocumentFinancialDetailsRecordsetArg
   DocumentFinancialDetailsRecordsetArg.MoveFirst
'Debug.Print "DocumentFinancialDetailsRecordsetArg.recordcount = " & DocumentFinancialDetailsRecordsetArg.recordCount
'Debug.Print "DocumentFinancialDetailsRecordsetArg(""DebitToAdd"") = " & DocumentFinancialDetailsRecordsetArg("DebitToAdd")
'Vartest = DocumentFinancialDetailsRecordsetArg("DebitToAdd")

   Do Until DocumentFinancialDetailsRecordsetArg.EOF
    If VarIntentionAffectsFinancial = True Then
     '___We take care updating the Financial transactor and Vat transactor
'     IterateRecordsets DocumentFinancialDetailsRecordsetArg
     VarFinTransactorDebitDifferenceToAddToTransactorDebitTotal = DocumentFinancialDetailsRecordsetArg("DebitToAdd")
     VarFinTransactorCreditDifferenceToAddToTransactorCreditTotal = DocumentFinancialDetailsRecordsetArg("CreditToAdd")
     VarNewFinancialTransactorID = DocumentFinancialDetailsRecordsetArg("NewFinancialTransactorID")
     VarOldFinancialTransactorID = Nz(DocumentFinancialDetailsRecordsetArg("OldFinancialTransactorID"), VarNewFinancialTransactorID)
     VarNewDebit = DocumentFinancialDetailsRecordsetArg("NewDebit")
     VarOldDebit = Nz(DocumentFinancialDetailsRecordsetArg("OldDebit"), 0)
     VarNewCredit = DocumentFinancialDetailsRecordsetArg("NewCredit")
     VarOldCredit = Nz(DocumentFinancialDetailsRecordsetArg("OldCredit"), 0)

' We proceed updating Financial Transactors TotalDebit and TotalCredit fields in the TransactorsT
   'We check if there was any change in TransactorID of the record
    'if no, then we simply update total debit and total credit of the transactor
     If VarNewFinancialTransactorID = VarOldFinancialTransactorID Then
       
       If DocIsDeletedArg = True Then
       db.Execute "Update TransactorsT SET TransactorsT.Total_Debit = TransactorsT.Total_Debit - """ & VarFinTransactorDebitDifferenceToAddToTransactorDebitTotal & _
       """ , TransactorsT.Total_Credit = TransactorsT.Total_Credit - """ & VarFinTransactorCreditDifferenceToAddToTransactorCreditTotal & _
       """ WHERE Transactor_ID = " & VarNewFinancialTransactorID, dbFailOnError
       Else
       db.Execute "Update TransactorsT SET TransactorsT.Total_Debit = TransactorsT.Total_Debit + """ & VarFinTransactorDebitDifferenceToAddToTransactorDebitTotal & _
       """ , TransactorsT.Total_Credit = TransactorsT.Total_Credit + """ & VarFinTransactorCreditDifferenceToAddToTransactorCreditTotal & _
       """ WHERE Transactor_ID = " & VarNewFinancialTransactorID, dbFailOnError
       End If
     Else
     'if yes, then we proceed updating Total Debit and Total Credit for both transactors respectively (for the new transactor we update using NewDebit and NewCredit, while for the old transactor (backedup) we use OldDebit and OldCredit

       db.Execute "Update TransactorsT SET TransactorsT.Total_Debit = TransactorsT.Total_Debit + """ & VarNewDebit & _
       """ , TransactorsT.Total_Credit = TransactorsT.Total_Credit + """ & VarNewCredit & _
       """ WHERE Transactor_ID = " & VarNewFinancialTransactorID, dbFailOnError
      
       db.Execute "Update TransactorsT SET TransactorsT.Total_Debit = TransactorsT.Total_Debit - """ & VarOldDebit & _
       """ , TransactorsT.Total_Credit = TransactorsT.Total_Credit - """ & VarOldCredit & _
       """ WHERE Transactor_ID = " & VarOldFinancialTransactorID, dbFailOnError
     
     End If
    End If
    DocumentFinancialDetailsRecordsetArg.MoveNext
   Loop

End If


UpdateTransactorsTotalDebitAndTotalCreditByDocumentFinancialDetailsRecordset = True


ExitProcedure:

If Not db Is Nothing Then
db.Close
Set db = Nothing
End If

Exit Function
   
ErrorHandler:

Select Case Err.Number
       Case 3167
        Resume Next
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: UpdateTransactorsTotalDebitAndTotalCreditByDocumentFinancialDetailsRecordset " & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            UpdateTransactorsTotalDebitAndTotalCreditByDocumentFinancialDetailsRecordset = False
            Resume ExitProcedure
            End Select

End Function

Public Sub CheckForFalslyFlaggedRecordsAsNew(TransactionIDArg As Long)
Debug.Print "Module Public Functions - " & "CheckForFalslyFlaggedRecordsAsNew " & Time()
'On Error GoTo Errorhandler

Dim db As DAO.Database
Dim rstTransaction
Dim rstDocuments As DAO.Recordset


Set db = CurrentDb

Set rstTransaction = db.OpenRecordset("select Transaction_ID, Is_new from TransactionsT where Transaction_ID = " & TransactionIDArg)
If Not rstTransaction.EOF Then
rstTransaction.MoveLast
rstTransaction.MoveFirst

Do Until rstTransaction.EOF
If rstTransaction("Is_New") = True Then
rstTransaction.Edit
rstTransaction("Is_New") = False
rstTransaction.Update
End If
rstTransaction.MoveNext
Loop
End If

Set rstDocuments = db.OpenRecordset("select Issued_Document_ID, Is_new from IssuedDocumentT where Transaction_ID = " & TransactionIDArg)
If Not rstDocuments.EOF Then
rstDocuments.MoveLast
rstDocuments.MoveFirst

Do Until rstDocuments.EOF
If rstDocuments("Is_New") = True Then
rstDocuments.Edit
rstDocuments("Is_New") = False
rstDocuments.Update
CheckForFalslyFlaggedFinancialOrProductDetailsRecordsAsNew (rstDocuments("Issued_Document_ID"))
End If
rstDocuments.MoveNext
Loop
End If

ExitProcedure:
Exit Sub
   
ErrorHandler:
Select Case Err.Number
       
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: CheckForFalslyFlaggedRecordsAsNew " & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select


End Sub


Public Sub CheckForFalslyFlaggedFinancialOrProductDetailsRecordsAsNew(IssuedDocumentIDArg As Long)
Debug.Print "Module Public Functions - " & "CheckForFalslyFlaggedFinancialOrProductDetailsRecordsAsNew " & Time()
'On Error GoTo Errorhandler

Dim db As DAO.Database
Dim rstDocumentFinancialDetails As DAO.Recordset
Dim rstDocumentProductDetails As DAO.Recordset

Set db = CurrentDb

Set rstDocumentFinancialDetails = db.OpenRecordset("select Issued_Document_Financial_Details_ID, Is_New from IssuedDocumentFinancialDetailsT where Issued_Document_ID = " & IssuedDocumentIDArg)
If Not rstDocumentFinancialDetails.EOF Then
rstDocumentFinancialDetails.MoveLast
rstDocumentFinancialDetails.MoveFirst

Do Until rstDocumentFinancialDetails.EOF
If rstDocumentFinancialDetails("Is_New") = True Then
rstDocumentFinancialDetails.Edit
rstDocumentFinancialDetails("Is_New") = False
rstDocumentFinancialDetails.Update
End If
rstDocumentFinancialDetails.MoveNext
Loop
End If

Set rstDocumentProductDetails = db.OpenRecordset("select Issued_Document_Product_Details_ID, Is_new from IssuedDocumentProductDetailsT where Issued_Document_ID = " & IssuedDocumentIDArg)
If Not rstDocumentProductDetails.EOF Then
rstDocumentProductDetails.MoveLast
rstDocumentProductDetails.MoveFirst

Do Until rstDocumentProductDetails.EOF
If rstDocumentProductDetails("Is_New") = True Then
rstDocumentProductDetails.Edit
rstDocumentProductDetails("Is_New") = False
rstDocumentProductDetails.Update
End If
Loop
End If

ExitProcedure:
Exit Sub
   
ErrorHandler:
Select Case Err.Number
       
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: CheckForFalslyFlaggedFinancialOrProductDetailsRecordsAsNew " & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select

End Sub

Public Sub LinkProductDetailsRecordsWithFinancialTransactors(IntentionTypeIDArg As Integer, IssuedDoumentIDArg As Long)
Debug.Print "Module Public Functions - " & "LinkProductDetailsRecordsWithFinancialTransactors " & Time()
'On Error GoTo Errorhandler

Dim db As DAO.Database
'Dim RstFormNewRecords As DAO.Recordset
'Dim RstFormEditRecords As DAO.Recordset
Dim rstDocumentProductDetails  As DAO.Recordset
Dim Response As Integer
Dim VarCurrentUser As Integer
Dim VarCurrentTimestamp As Date
Dim VarAttributeValueNumber As Long
Dim VarFinancialTransactorID As Long
Dim VarLinkAttributeValueToEntitiesID As Long
Dim VarVatTransactorID As Long
Dim VarAttributeID As Long

VarCurrentUser = FetchUserID
VarCurrentTimestamp = Now()

Set db = CurrentDb

'Set RstFormNewRecords = db.OpenRecordset("SELECT IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID, " & _
"LinkProductsToCompanyFinancialTransactorsT.Link_Products_To_Company_Financial_Transactors_ID, LinkProductsToCompanyFinancialTransactorsT.Transactor_Financial_ID, " & _
"[LinkFinTransactorsToVatTransactorsWithVat%Q].Vat_Transactor_ID " & _
"FROM (IssuedDocumentProductDetailsT INNER JOIN LinkProductsToCompanyFinancialTransactorsT ON " & _
"(IssuedDocumentProductDetailsT.Product_ID = LinkProductsToCompanyFinancialTransactorsT.Product_ID) " & _
"AND (IssuedDocumentProductDetailsT.Accounting_Behavior_ID = LinkProductsToCompanyFinancialTransactorsT.Accounting_Behavior_ID)) " & _
"INNER JOIN [LinkFinTransactorsToVatTransactorsWithVat%Q] ON (IssuedDocumentProductDetailsT.[VAT%] = [LinkFinTransactorsToVatTransactorsWithVat%Q].[Vat%]) " & _
"AND (LinkProductsToCompanyFinancialTransactorsT.Transactor_Financial_ID = [LinkFinTransactorsToVatTransactorsWithVat%Q].Financial_Transactor_ID) " & _
"WHERE LinkProductsToCompanyFinancialTransactorsT.Intention_Type_ID = " & IntentionTypeIDArg & " AND IssuedDocumentProductDetailsT.Is_Deleted = False " & _
"AND IssuedDocumentProductDetailsT.Issued_Document_ID = " & IssuedDoumentIDArg)


'Set RstFormEditRecords = db.OpenRecordset("SELECT IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID, " & _
"LinkProductsToCompanyFinancialTransactorsT.Transactor_Financial_ID, [LinkFinTransactorsToVatTransactorsWithVat%Q].Vat_Transactor_ID, " & _
"LinkAttributeValueToEntitiesT.Link_Attribute_Value_To_Entity_ID, LinkAttributeValueToEntitiesT.Attribute_ID, LinkAttributeValueToEntitiesT.Attribute_Value_Number " & _
"FROM ((IssuedDocumentProductDetailsT INNER JOIN LinkProductsToCompanyFinancialTransactorsT " & _
"ON (IssuedDocumentProductDetailsT.Accounting_Behavior_ID = LinkProductsToCompanyFinancialTransactorsT.Accounting_Behavior_ID) " & _
"AND (IssuedDocumentProductDetailsT.Product_ID = LinkProductsToCompanyFinancialTransactorsT.Product_ID)) " & _
"INNER JOIN [LinkFinTransactorsToVatTransactorsWithVat%Q] " & _
"ON (LinkProductsToCompanyFinancialTransactorsT.Transactor_Financial_ID = [LinkFinTransactorsToVatTransactorsWithVat%Q].Financial_Transactor_ID) " & _
"AND (IssuedDocumentProductDetailsT.[VAT%] = [LinkFinTransactorsToVatTransactorsWithVat%Q].[Vat%])) " & _
"INNER JOIN LinkAttributeValueToEntitiesT ON IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID = LinkAttributeValueToEntitiesT.Entity_ID " & _
"WHERE (((LinkAttributeValueToEntitiesT.Attribute_ID) = 352) AND ((LinkProductsToCompanyFinancialTransactorsT.Intention_Type_ID) = " & IntentionTypeIDArg & ") " & _
"AND ((IssuedDocumentProductDetailsT.Is_Deleted)=False) AND ((IssuedDocumentProductDetailsT.Issued_Document_ID)= " & IssuedDoumentIDArg & ") AND ((LinkAttributeValueToEntitiesT.Entity_Type_ID)=5))")

Set rstDocumentProductDetails = db.OpenRecordset("SELECT DISTINCTROW LinkAttributeValueToEntitiesT.Link_Attribute_Value_To_Entity_ID, " & _
"IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID AS Entity_ID, LinkAttributeValueToEntitiesT.Attribute_ID, LinkAttributeValueToEntitiesT.Attribute_Value_Number " & _
"FROM ((IssuedDocumentProductDetailsT LEFT JOIN LinkAttributeValueToEntitiesT ON IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID = LinkAttributeValueToEntitiesT.Entity_ID) " & _
"INNER JOIN LinkProductsToCompanyFinancialTransactorsT ON (IssuedDocumentProductDetailsT.Accounting_Behavior_ID = LinkProductsToCompanyFinancialTransactorsT.Accounting_Behavior_ID) " & _
"AND (IssuedDocumentProductDetailsT.Product_ID = LinkProductsToCompanyFinancialTransactorsT.Product_ID)) " & _
"INNER JOIN [LinkFinTransactorsToVatTransactorsWithVat%Q] ON (LinkProductsToCompanyFinancialTransactorsT.Transactor_Financial_ID = [LinkFinTransactorsToVatTransactorsWithVat%Q].Financial_Transactor_ID) " & _
"AND (IssuedDocumentProductDetailsT.[VAT%] = [LinkFinTransactorsToVatTransactorsWithVat%Q].[Vat%]) " & _
"WHERE (((LinkAttributeValueToEntitiesT.Attribute_ID)=352 Or (LinkAttributeValueToEntitiesT.Attribute_ID) Is Null) " & _
"AND ((IssuedDocumentProductDetailsT.Is_Deleted)=False) AND ((IssuedDocumentProductDetailsT.Issued_Document_ID)=1584) " & _
"AND ((LinkAttributeValueToEntitiesT.Entity_Type_ID) = 5 Or (LinkAttributeValueToEntitiesT.Entity_Type_ID) Is Null))")


'in case we have edited records
'first we check if recordset has records
If Not RstFormEditRecords.EOF Then
    RstFormEditRecords.MoveLast
    RstFormEditRecords.MoveFirst
    'Debug.Print "RstFormEditRecords.recordCount = " & RstFormEditRecords.recordCount
    'after we itterate each record of the recordset
    Do Until RstFormEditRecords.EOF
    'Then we capture to variables IssuedDocumentProductDetailsID, FinancialTransactorID, VatTransactorID, AttributeID and Attribute_Value_Number
    varIssuedDocumentProductDetailsID = RstFormEditRecords(0)
    VarFinancialTransactorID = RstFormEditRecords(1)
    VarVatTransactorID = RstFormEditRecords(2)
    VarAttributeID = RstFormEditRecords(4)
    VarAttributeValueNumber = RstFormEditRecords(5)
    VarLinkAttributeValueToEntitiesID = RstFormEditRecords(3)
     'we check if the new value (of financial transactor as AttributeValueNumber) is different of the one that we want to update to.
      If VarFinancialTransactorID <> VarAttributeValueNumber Then
      'If the value is different then we warn the user with a message about the forthcoming change and he, by pressing "YES" the change continues and by pressing "NO" change is aborted.
        Response = MsgBox("The linked Financial transactor for the record " & RstFormEditRecords(0) & " was ID " & RstFormEditRecords(3) & " and it is about to change to ID " & VarFinancialTransactorID & ". Click ""YES"" to proceed to change or ""NO"" to abort this change.", vbYesNo + vbInformation)
        If Response = vbYes Then
           db.Execute "Update LinkAttributeValueToEntitiesT SET LinkAttributeValueToEntitiesT.Attribute_Value_Number = " & VarFinancialTransactorID & " WHERE LinkAttributeValueToEntitiesT.Link_Attribute_Value_To_Entity_ID = " & VarLinkAttributeValueToEntitiesID, dbFailOnError
        End If
      End If
    RstFormEditRecords.MoveNext
    Loop
   GoTo ExitProcedure
End If

'in case this is new data entry (new records)
'first we check if recordset has records
If Not RstFormNewRecords.EOF Then
   RstFormNewRecords.MoveLast
   RstFormNewRecords.MoveFirst
      
   'after we itterate each record of the recordset
   Do Until RstFormNewRecords.EOF
   'Then we capture to variables IssuedDocumentProductDetailsID, FinancialTransactorID and VatTransactorID
   varIssuedDocumentProductDetailsID = RstFormNewRecords(0)
   VarFinancialTransactorID = RstFormNewRecords(2)
   VarVatTransactorID = RstFormNewRecords(3)
   'after, we insert to the table LinkAttributeValueToEntitiesT a new record linking the IssuedDocumentProductDetailsID (as Entity_ID) with the Attribute FINANCIAL TRANSACTOR LINKED TO PRODUCT DETAILS ID _
   which takes the value we just captured with the variable VarFinancialTransactorID
   db.Execute "Insert Into LinkAttributeValueToEntitiesT (Entity_Type_ID, Entity_ID, Attribute_ID, Attribute_Value_Number, Is_Included_To_Suggested_Attributes, " & _
   "AttLinkToEntityInsert_UserID, AttLinkToEntityInsert_Timestamp) Values(5, " & varIssuedDocumentProductDetailsID & ", 352, " & VarFinancialTransactorID & ", false, " & _
   VarCurrentUser & ", #" & Format(VarCurrentTimestamp, "yyyy-mm-dd hh:nn:ss") & "# );", dbFailOnError
   RstFormNewRecords.MoveNext
   Loop
  GoTo ExitProcedure
End If

ExitProcedure:
If Not (RstFormNewRecords) Is Nothing Then
RstFormNewRecords.Close
Set RstFormNewRecords = Nothing
End If

If Not (RstFormEditRecords) Is Nothing Then
RstFormEditRecords.Close
Set RstFormEditRecords = Nothing
End If

If Not (db) Is Nothing Then
db.Close
Set db = Nothing
End If

Exit Sub
   
ErrorHandler:
Select Case Err.Number
       
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: LinkProductDetailsRecordsWithFinancialTransactors " & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
            
End Sub

Public Sub LinkProductDetailsRecordsWithVatTransactors(IntentionTypeIDArg As Integer, IssuedDoumentIDArg As Long)  ', IsNewTransactionArg as Boolean)
Debug.Print "Module Public Functions - " & "LinkProductDetailsRecordsWithVatTransactors " & Time()
'On Error GoTo Errorhandler

Dim db As DAO.Database
Dim RstFormNewRecords As DAO.Recordset
Dim RstFormEditRecords As DAO.Recordset
Dim Response As Integer
Dim VarCurrentUser As Integer
Dim VarCurrentTimestamp As Date
Dim VarAttributeValueNumber As Long
Dim VarFinancialTransactorID As Long
Dim VarLinkAttributeValueToEntitiesID As Long
Dim VarVatTransactorID As Long
Dim VarAttributeID As Long

VarCurrentUser = FetchUserID
VarCurrentTimestamp = Now()

Set db = CurrentDb

Set RstFormNewRecords = db.OpenRecordset("SELECT IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID, " & _
"LinkProductsToCompanyFinancialTransactorsT.Link_Products_To_Company_Financial_Transactors_ID, LinkProductsToCompanyFinancialTransactorsT.Transactor_Financial_ID, " & _
"[LinkFinTransactorsToVatTransactorsWithVat%Q].Vat_Transactor_ID " & _
"FROM (IssuedDocumentProductDetailsT INNER JOIN LinkProductsToCompanyFinancialTransactorsT ON " & _
"(IssuedDocumentProductDetailsT.Product_ID = LinkProductsToCompanyFinancialTransactorsT.Product_ID) " & _
"AND (IssuedDocumentProductDetailsT.Accounting_Behavior_ID = LinkProductsToCompanyFinancialTransactorsT.Accounting_Behavior_ID)) " & _
"INNER JOIN [LinkFinTransactorsToVatTransactorsWithVat%Q] ON (IssuedDocumentProductDetailsT.[VAT%] = [LinkFinTransactorsToVatTransactorsWithVat%Q].[Vat%]) " & _
"AND (LinkProductsToCompanyFinancialTransactorsT.Transactor_Financial_ID = [LinkFinTransactorsToVatTransactorsWithVat%Q].Financial_Transactor_ID) " & _
"WHERE LinkProductsToCompanyFinancialTransactorsT.Intention_Type_ID = " & IntentionTypeIDArg & " AND IssuedDocumentProductDetailsT.Is_Deleted = False " & _
"AND IssuedDocumentProductDetailsT.Issued_Document_ID = " & IssuedDoumentIDArg)


Set RstFormEditRecords = db.OpenRecordset("SELECT IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID, " & _
"LinkProductsToCompanyFinancialTransactorsT.Transactor_Financial_ID, [LinkFinTransactorsToVatTransactorsWithVat%Q].Vat_Transactor_ID, " & _
"LinkAttributeValueToEntitiesT.Link_Attribute_Value_To_Entity_ID, LinkAttributeValueToEntitiesT.Attribute_ID, LinkAttributeValueToEntitiesT.Attribute_Value_Number " & _
"FROM ((IssuedDocumentProductDetailsT INNER JOIN LinkProductsToCompanyFinancialTransactorsT " & _
"ON (IssuedDocumentProductDetailsT.Accounting_Behavior_ID = LinkProductsToCompanyFinancialTransactorsT.Accounting_Behavior_ID) " & _
"AND (IssuedDocumentProductDetailsT.Product_ID = LinkProductsToCompanyFinancialTransactorsT.Product_ID)) " & _
"INNER JOIN [LinkFinTransactorsToVatTransactorsWithVat%Q] " & _
"ON (LinkProductsToCompanyFinancialTransactorsT.Transactor_Financial_ID = [LinkFinTransactorsToVatTransactorsWithVat%Q].Financial_Transactor_ID) " & _
"AND (IssuedDocumentProductDetailsT.[VAT%] = [LinkFinTransactorsToVatTransactorsWithVat%Q].[Vat%])) " & _
"INNER JOIN LinkAttributeValueToEntitiesT ON IssuedDocumentProductDetailsT.Issued_Document_Product_Details_ID = LinkAttributeValueToEntitiesT.Entity_ID " & _
"WHERE (((LinkAttributeValueToEntitiesT.Attribute_ID) = 351) AND ((LinkProductsToCompanyFinancialTransactorsT.Intention_Type_ID) = " & IntentionTypeIDArg & ") " & _
"AND ((IssuedDocumentProductDetailsT.Is_Deleted)=False) AND ((IssuedDocumentProductDetailsT.Issued_Document_ID)= " & IssuedDoumentIDArg & ") AND ((LinkAttributeValueToEntitiesT.Entity_Type_ID)=5))")



'in case we have edited records
'first we check if recordset has records
If Not RstFormEditRecords.EOF Then
    RstFormEditRecords.MoveLast
    RstFormEditRecords.MoveFirst
    Debug.Print "RstFormEditRecords.recordCount = " & RstFormEditRecords.recordcount
    'after we itterate each record of the recordset
    Do Until RstFormEditRecords.EOF
    'Then we capture to variables IssuedDocumentProductDetailsID, FinancialTransactorID, VatTransactorID, AttributeID and Attribute_Value_Number
    varIssuedDocumentProductDetailsID = RstFormEditRecords(0)
    VarFinancialTransactorID = RstFormEditRecords(1)
    VarVatTransactorID = RstFormEditRecords(2)
    VarAttributeID = RstFormEditRecords(4)
    VarAttributeValueNumber = RstFormEditRecords(5)
    VarLinkAttributeValueToEntitiesID = RstFormEditRecords(3)
    ' we check if the new value (of VAT transactor as AttributeValueNumber) is different of the one that we want to update to.
      If VarVatTransactorID <> VarAttributeValueNumber Then
       'If the value is different then we warn the user with a message about the forthcoming change and he, by pressing "YES" the change continues and by pressing "NO" change is aborted.
         Response = MsgBox("The linked VAT transactor for the record " & RstFormEditRecords(0) & " was ID " & RstFormEditRecords(5) & " and it is about to change to ID " & VarVatTransactorID & ". Click ""YES"" to proceed to change or ""NO"" to abort this change.", vbYesNo + vbInformation)
         If Response = vbYes Then
           db.Execute "Update LinkAttributeValueToEntitiesT SET LinkAttributeValueToEntitiesT.Attribute_Value_Number = " & VarVatTransactorID & " WHERE LinkAttributeValueToEntitiesT.Link_Attribute_Value_To_Entity_ID = " & VarLinkAttributeValueToEntitiesID, dbFailOnError
         End If
      End If
    
    RstFormEditRecords.MoveNext
    Loop
    GoTo ExitProcedure
End If

'in case this is new data entry (new records)
'first we check if recordset has records
If Not RstFormNewRecords.EOF Then
   RstFormNewRecords.MoveLast
   RstFormNewRecords.MoveFirst
   Debug.Print "RstFormNewRecords.recordCount = " & RstFormNewRecords.recordcount
   
   'after we itterate each record of the recordset
   Do Until RstFormNewRecords.EOF
   'Then we capture to variables IssuedDocumentProductDetailsID, FinancialTransactorID and VatTransactorID
   varIssuedDocumentProductDetailsID = RstFormNewRecords(0)
   VarFinancialTransactorID = RstFormNewRecords(2)
   VarVatTransactorID = RstFormNewRecords(3)
   'after, we insert to the table LinkAttributeValueToEntitiesT a new record linking the IssuedDocumentProductDetailsID (as Entity_ID) with the Attribute VAT TRANSACTOR ID LINKED TO PRODUCT DETAILS ID _
   which takes the value we just captured with the variable VarVatTransactorID
   db.Execute "Insert Into LinkAttributeValueToEntitiesT (Entity_Type_ID, Entity_ID, Attribute_ID, Attribute_Value_Number, Is_Included_To_Suggested_Attributes, " & _
   "AttLinkToEntityInsert_UserID, AttLinkToEntityInsert_Timestamp) Values(5, " & varIssuedDocumentProductDetailsID & ", 351, " & VarVatTransactorID & ", false, " & _
   VarCurrentUser & ", #" & Format(VarCurrentTimestamp, "yyyy-mm-dd hh:nn:ss") & "# );", dbFailOnError
   
   RstFormNewRecords.MoveNext
   Loop
   GoTo ExitProcedure
End If

ExitProcedure:
If Not (RstFormNewRecords) Is Nothing Then
RstFormNewRecords.Close
Set RstFormNewRecords = Nothing
End If

If Not (RstFormEditRecords) Is Nothing Then
RstFormEditRecords.Close
Set RstFormEditRecords = Nothing
End If

If Not (db) Is Nothing Then
db.Close
Set db = Nothing
End If

Exit Sub
   
ErrorHandler:
Select Case Err.Number
       
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: LinkProductDetailsRecordsWithVatTransactors " & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
            
End Sub

Public Function FetchTransactorAllAddressDetailsAsOneString(TransactorIDArg As Long) As String
Debug.Print "Module Public Functions - " & "FetchTransactorAllAddressDetailsAsOneString " & Time()
'On Error GoTo Errorhandler

Dim db As DAO.Database
Dim RstAddressDetails As DAO.Recordset
Dim ConcatenatedString As String

Set db = CurrentDb

Set RstAddressDetails = db.OpenRecordset("select Street_Name, Street_Number, Postal_Code, City_Description, County_Description, Country_Description from TransactorAddressDetailsWithDescriptionsQ " & _
"where Transactor_ID = " & TransactorIDArg)

If Not RstAddressDetails.EOF Then
RstAddressDetails.MoveLast
RstAddressDetails.MoveFirst

Do Until RstAddressDetails.EOF
ConcatenatedString = ConcatenatedString & RstAddressDetails("Street_Name") & ", " & RstAddressDetails("Street_Number") & ", " & _
RstAddressDetails("Postal_Code") & ", " & RstAddressDetails("City_Description") & ", " & RstAddressDetails("County_Description") & ", " & RstAddressDetails("Country_Description") & " | "
RstAddressDetails.MoveNext
Loop
End If

' Remove the trailing comma and space if the string is not empty
    If Len(ConcatenatedString) > 0 Then
        ConcatenatedString = Left(ConcatenatedString, Len(ConcatenatedString) - 2)
    End If
    
FetchTransactorAllAddressDetailsAsOneString = ConcatenatedString

ExitProcedure:
If Not (RstAddressDetails) Is Nothing Then
RstAddressDetails.Close
Set RstAddressDetails = Nothing
End If

If Not (db) Is Nothing Then
db.Close
Set db = Nothing
End If

Exit Function
   
ErrorHandler:
Select Case Err.Number
       
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: FetchTransactorAllAddressDetailsAsOneString " & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
            

End Function

Public Function FetchAllEmailAddressesWithDelimiterFromBasicTransactorsRecordset(RecordsetArg As Recordset) As String
Debug.Print "Module Public Functions - " & "FetchAllEmailAddressesWithDelimiterFromBasicTransactorsRecordset " & Time()
'On Error GoTo Errorhandler

Dim db As DAO.Database
Dim RstBasicTransactors As DAO.Recordset
Dim RstBasicTransactorsWithEmails As DAO.Recordset
Dim ConcatenatedString As String
Dim VarBasicTransactorID As Long
Dim SqlStringForRstBasicTransactorsWithEmails As String

FetchAllEmailAddressesWithDelimiterFromBasicTransactorsRecordset = ""

Set db = CurrentDb
Set RstBasicTransactors = RecordsetArg

' Create the temporary table
If TableExists("TempTransactorsT") Then
CurrentDb.Execute ("DROP TABLE TempTransactorsT")
End If

    db.Execute "CREATE TABLE TempTransactorsT (ID AUTOINCREMENT PRIMARY KEY, Basic_Transactor_ID Long);"
    
If Not RstBasicTransactors.EOF Then
RstBasicTransactors.MoveLast
RstBasicTransactors.MoveFirst
Do Until RstBasicTransactors.EOF
VarBasicTransactorID = RstBasicTransactors(0)
db.Execute "INSERT INTO TempTransactorsT (Basic_Transactor_ID) VALUES (" & VarBasicTransactorID & ");"
RstBasicTransactors.MoveNext
Loop
End If

SqlStringForRstBasicTransactorsWithEmails = "Select TransactorsEmailsQ.Transactor_Contact_Decription from " & _
"TransactorsEmailsQ inner join BasicTransactorsFormTempTransactorsTQ on TransactorsEmailsQ.Basic_Tsansactor_ID_FK = BasicTransactorsFormTempTransactorsTQ.Basic_Transactor_ID"
Debug.Print "SqlStringForRstBasicTransactorsWithEmails = " & SqlStringForRstBasicTransactorsWithEmails
Set RstBasicTransactorsWithEmails = db.OpenRecordset(SqlStringForRstBasicTransactorsWithEmails, dbOpenSnapshot, dbReadOnly)
Debug.Print "RstBasicTransactorsWithEmails.recordcount  = " & RstBasicTransactorsWithEmails.recordcount
If Not RstBasicTransactorsWithEmails.EOF Then
RstBasicTransactorsWithEmails.MoveLast
RstBasicTransactorsWithEmails.MoveFirst
Do Until RstBasicTransactorsWithEmails.EOF
ConcatenatedString = ConcatenatedString & RstBasicTransactorsWithEmails(0) & ";"
RstBasicTransactorsWithEmails.MoveNext
Loop
Debug.Print "ConcatenatedString = " & ConcatenatedString
' Remove the trailing comma and space if the string is not empty
    If Len(ConcatenatedString) > 0 Then
        ConcatenatedString = Left(ConcatenatedString, Len(ConcatenatedString) - 1)
    End If
    
FetchAllEmailAddressesWithDelimiterFromBasicTransactorsRecordset = ConcatenatedString
End If


ExitProcedure:

If Not RstBasicTransactors Is Nothing Then
RstBasicTransactors.Close
Set RstBasicTransactors = Nothing
End If

If Not RstBasicTransactorsWithEmails Is Nothing Then
RstBasicTransactorsWithEmails.Close
Set RstBasicTransactorsWithEmails = Nothing
End If

db.Execute ("Drop table TempTransactorsT"), dbFailOnError

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
            "Error Source: FetchAllEmailAddressesWithDelimiterFromBasicTransactorsRecordset " & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
            

End Function

Public Function FetchTransactorAllContactDetailsAsOneString(BasicTransactorIDArg As Long) As String
Debug.Print "Module Public Functions - " & "FetchTransactorAllContactDetailsAsOneString " & Time()
'On Error GoTo Errorhandler

Dim db As DAO.Database
Dim RstContactDetails As DAO.Recordset
Dim ConcatenatedString As String
'Exit Function
Set db = CurrentDb
Set RstContactDetails = db.OpenRecordset("select Transactor_Contact_Decription from TransactorContactDetailsT " & _
"where Basic_Tsansactor_ID_FK = " & BasicTransactorIDArg)

If Not RstContactDetails.EOF Then
RstContactDetails.MoveLast
RstContactDetails.MoveFirst

Do Until RstContactDetails.EOF
ConcatenatedString = ConcatenatedString & RstContactDetails("Transactor_Contact_Decription") & ", "
RstContactDetails.MoveNext
Loop
End If

' Remove the trailing comma and space if the string is not empty
    If Len(ConcatenatedString) > 0 Then
        ConcatenatedString = Left(ConcatenatedString, Len(ConcatenatedString) - 2)
    End If
    
FetchTransactorAllContactDetailsAsOneString = ConcatenatedString

ExitProcedure:
If Not (RstContactDetails) Is Nothing Then
RstContactDetails.Close
Set RstContactDetails = Nothing
End If

If Not (db) Is Nothing Then
db.Close
Set db = Nothing
End If

Exit Function
   
ErrorHandler:
Select Case Err.Number
       
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: FetchTransactorAllContactDetailsAsOneString " & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
            

End Function

Public Function FetchEntityMultipleAttributeValuesAsOneString(AttributeIDArg As Integer, EntityIDArg As Long, EntityTypeIDArg As Integer) As String
Debug.Print "Module Public Functions - " & "FetchEntityMultipleAttributeValuesAsOneString " & Time()
'On Error GoTo Errorhandler

Dim db As DAO.Database
Dim RstAttributeValues As DAO.Recordset
Dim ConcatenatedString As String

Set db = CurrentDb
Set RstAttributeValues = db.OpenRecordset("select Attribute_Value_String from LinkAttributeValueToEntitiesT " & _
"where Entity_ID = " & EntityIDArg & " AND Entity_Type_ID = " & EntityTypeIDArg & " AND Attribute_ID = " & AttributeIDArg)

If Not RstAttributeValues.EOF Then
RstAttributeValues.MoveLast
RstAttributeValues.MoveFirst

Do Until RstAttributeValues.EOF
ConcatenatedString = ConcatenatedString & RstAttributeValues("Attribute_Value_String") & ", "
RstAttributeValues.MoveNext
Loop
End If

' Remove the trailing comma and space if the string is not empty
    If Len(ConcatenatedString) > 0 Then
        ConcatenatedString = Left(ConcatenatedString, Len(ConcatenatedString) - 2)
    End If
    
FetchEntityMultipleAttributeValuesAsOneString = ConcatenatedString

ExitProcedure:
If Not (RstAttributeValues) Is Nothing Then
RstAttributeValues.Close
Set RstAttributeValues = Nothing
End If

If Not (db) Is Nothing Then
db.Close
Set db = Nothing
End If

Exit Function
   
ErrorHandler:
Select Case Err.Number
       
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: FetchEntityMultipleAttributeValuesAsOneString " & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
            

End Function


Public Sub LinkFinancialTransactorRecordsWithVatTransactors(IssuedDocumentFinancialDetailsArg As Long, TransactorIDArg As Long, VatPercentageArg As Double, IsNewRecord As Boolean)
Debug.Print "Module Public Functions - " & "LinkFinancialTransactorRecordsWithVatTransactors " & Time()
'On Error GoTo Errorhandler

Dim VarQueryForInsertVatTransactorStr As String
Dim VarQueryForInsertVatPercentageStr As String
Dim db As DAO.Database
Dim VarCurrentUser As Integer
Dim VarCurrentTimestamp As Date

Set db = CurrentDb
 
VarCurrentUser = FetchUserID
VarCurrentTimestamp = Now()

If CheckIfEntityMayHaveMultipleAttributeValues(351, IssuedDocumentFinancialDetailsArg, 4) = 0 Then
    VarQueryForInsertVatTransactorStr = "Insert Into LinkAttributeValueToEntitiesT (Entity_Type_ID, Entity_ID, Attribute_ID, Attribute_Value_Number, Is_Included_To_Suggested_Attributes, " & _
   "AttLinkToEntityInsert_UserID, AttLinkToEntityInsert_Timestamp) Values(4, " & IssuedDocumentFinancialDetailsArg & ", 351, " & TransactorIDArg & ", false, " & _
    VarCurrentUser & ", #" & Format(VarCurrentTimestamp, "yyyy-mm-dd hh:nn:ss") & "# );"
    db.Execute (VarQueryForInsertVatTransactorStr), dbFailOnError
Else
    VarQueryForInsertVatTransactorStr = "Update LinkAttributeValueToEntitiesT set Attribute_Value_Number = " & TransactorIDArg & _
   " where Entity_Type_ID = 4 AND Entity_ID = " & IssuedDocumentFinancialDetailsArg & " AND Attribute_ID = 351 "
    db.Execute (VarQueryForInsertVatTransactorStr), dbFailOnError
End If


If CheckIfEntityMayHaveMultipleAttributeValues(16, IssuedDocumentFinancialDetailsArg, 4) = 0 Then
VarQueryForInsertVatPercentageStr = "Insert Into LinkAttributeValueToEntitiesT (Entity_Type_ID, Entity_ID, Attribute_ID, Attribute_Value_Number, Is_Included_To_Suggested_Attributes, " & _
   "AttLinkToEntityInsert_UserID, AttLinkToEntityInsert_Timestamp) Values(4, " & IssuedDocumentFinancialDetailsArg & ", 16, """ & CStr(VatPercentageArg) & """, false, " & _
    VarCurrentUser & ", #" & Format(VarCurrentTimestamp, "yyyy-mm-dd hh:nn:ss") & "# );"
    db.Execute (VarQueryForInsertVatPercentageStr), dbFailOnError
Else
VarQueryForInsertVatPercentageStr = "Update LinkAttributeValueToEntitiesT set Attribute_Value_Number = """ & CStr(VatPercentageArg) & _
   """ where Entity_Type_ID = 4 AND Entity_ID = " & IssuedDocumentFinancialDetailsArg & " AND Attribute_ID = 16 "
    db.Execute (VarQueryForInsertVatPercentageStr), dbFailOnError
End If

ExitProcedure:

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
            "Error Source: LinkFinancialTransactorRecordsWithVatTransactors " & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
            
End Sub

Public Sub InsertOrUpdateHelpTableForTransactorSearchByContactAndAddressDetails(TransactorIDArg As Long, BasicTransactorIDArg As Long)
Debug.Print "Module Public Functions - " & "InsertOrUpdateHelpTableForTransactorSearchByContactAndAddressDetails " & Time()
'On Error GoTo Errorhandler

Dim db As DAO.Database
Dim rstHelpTable As DAO.Recordset
Dim AddressDetailsString As String
Dim ContactDetailsString As String

AddressDetailsString = FetchTransactorAllAddressDetailsAsOneString(TransactorIDArg)
ContactDetailsString = FetchTransactorAllContactDetailsAsOneString(BasicTransactorIDArg)

Debug.Print "Insert into HelpTableForTransactorSearchByContactAndAddressDetailsT Transactor_ID, Basic_Transactor_ID, Address_String, Contact_String " & _
" values(" & TransactorIDArg & " , " & BasicTransactorIDArg & " , """ & AddressDetailsString & """ , """ & ContactDetailsString & """"

Set db = CurrentDb
Set rstHelpTable = db.OpenRecordset("Select * from HelpTableForTransactorSearchByContactAndAddressDetailsT where Transactor_ID = " & TransactorIDArg)

If rstHelpTable.EOF Then

db.Execute "Insert into HelpTableForTransactorSearchByContactAndAddressDetailsT (Transactor_ID, Basic_Transactor_ID, Address_String, Contact_String) " & _
" values(" & TransactorIDArg & " , " & BasicTransactorIDArg & " , """ & AddressDetailsString & """ , """ & ContactDetailsString & """)", dbFailOnError

Else

db.Execute "Update HelpTableForTransactorSearchByContactAndAddressDetailsT set Address_String = """ & AddressDetailsString & _
""", Contact_String = """ & ContactDetailsString & """ where Transactor_ID = " & TransactorIDArg, dbFailOnError

End If

ExitProcedure:

If Not rstHelpTable Is Nothing Then
rstHelpTable.Close
Set rstHelpTable = Nothing
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
            "Error Source: InsertOrUpdateHelpTableForTransactorSearchByContactAndAddressDetails " & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
End Sub


Public Function UpdateTransactorsBalanceAndInventoryBalanceByProductDocumentDetailsRecordset(ProductDocumentDetailsRecordsetArg As Recordset, IntentionIDArg As Integer, Optional DocIsDeletedArg As Boolean) As Boolean
Debug.Print "Module Public Functions - " & "UpdateTransactorsBalanceAndInventoryBalanceByProductDocumentDetailsRecordset " & Time()
'On Error GoTo Errorhandler

Dim VarFinTransactorDebitDifferenceToAddToTransactorDebitTotal As Double
Dim VarFinTransactorCreditDifferenceToAddToTransactorCreditTotal As Double
Dim VarVatTransactorDebitDifferenceToAddToTransactorDebitTotal As Double
Dim VarVatTransactorCreditDifferenceToAddToTransactorCreditTotal As Double
Dim VarQuantityDifferenceToAddToInvntoryDebitTotal As Double
Dim VarQuantityDifferenceToAddToInvntoryCreditTotal As Double
Dim VarNewFinancialTransactorID As Long
Dim VarOldFinancialTransactorID As Long
Dim VarNewVatTransactorID As Long
Dim VarOldVatTransactorID As Long
Dim VarNewProductID As Long
Dim VarOldProductID As Long
Dim VarNewMainEntityID As Long
Dim VarOldMainEntityID As Long
Dim VarNewOtherEntityID As Long
Dim VarOldOtherEntityID As Long
Dim VarNewDebit As Double
Dim VarOldDebit As Double
Dim VarNewCredit As Double
Dim VarOldCredit As Double
Dim VarNewVatDebit As Double
Dim VarOldVatDebit As Double
Dim VarNewVatCredit As Double
Dim VarOldVatCredit As Double
Dim VarNewQuantityDebit As Double
Dim VarOldQuantityDebit As Double
Dim VarNewQuantityCredit As Double
Dim VarOldQuantityCredit As Double
Dim VarAffectsInventoy As Integer
Dim VarAffectsFinancial As Boolean
Dim VarProductDetailsProduceFinancialResults As Boolean
Dim db As DAO.Database

Set db = CurrentDb

VarAffectsInventoy = DLookup("IntentionsT.Affects_Inventory", "IntentionsT", "Intention_ID = " & IntentionIDArg)
VarProductDetailsProduceFinancialResults = DLookup("IntentionsT.Product_Details_Produce_Financial_Results", "IntentionsT", "Intention_ID = " & IntentionIDArg)
VarAffectsFinancial = DLookup("IntentionsT.Affects_Financial", "IntentionsT", "Intention_ID = " & IntentionIDArg)
UpdateTransactorsBalanceByProductDocumentRecordset = False
'________________
'1. updating Financial Transactors TotalDebit and TotalCredit fields in the TransactorsT for the link financial transactors with the product.
'2. updating Financial Transactors TotalDebit and TotalCredit fields in the TransactorsT for the link Vat transactors with the product.
'3. updating ProductInventoryBalanceT TotalDebit and TotalCredit fields for the specific product at the specific warehouse (main entity).
'4. updating ProductInventoryBalanceT TotalDebit and TotalCredit fields for the specific product at the specific warehouse (other entity).

 Call IterateRecordsets(ProductDocumentDetailsRecordsetArg, "Following is iteration of ProductDetailsRecordset with fields for update related Financial Transactors and Inventory Balance")
 'We proceed to  execute the above 4 actions one by one:
If Not ProductDocumentDetailsRecordsetArg.EOF Then
   ProductDocumentDetailsRecordsetArg.MoveLast
   ProductDocumentDetailsRecordsetArg.MoveFirst

   Do Until ProductDocumentDetailsRecordsetArg.EOF
  If VarAffectsFinancial = True Then
     '___We take care updating the Financial transactor and Vat transactor
     VarFinTransactorDebitDifferenceToAddToTransactorDebitTotal = ProductDocumentDetailsRecordsetArg("DebitToAdd")
     VarFinTransactorCreditDifferenceToAddToTransactorCreditTotal = ProductDocumentDetailsRecordsetArg("CreditToAdd")
     VarVatTransactorDebitDifferenceToAddToTransactorDebitTotal = ProductDocumentDetailsRecordsetArg("VatDebitToAdd")
     VarVatTransactorCreditDifferenceToAddToTransactorCreditTotal = ProductDocumentDetailsRecordsetArg("VatCreditToAdd")
     VarNewFinancialTransactorID = ProductDocumentDetailsRecordsetArg("NewFinancialTransactorID")
     VarOldFinancialTransactorID = Nz(ProductDocumentDetailsRecordsetArg("OldFinancialTransactorID"), VarNewFinancialTransactorID)
     VarNewVatTransactorID = ProductDocumentDetailsRecordsetArg("NewVatTransactorID")
     VarOldVatTransactorID = Nz(ProductDocumentDetailsRecordsetArg("OldVatTransactorID"), VarNewVatTransactorID)
     VarNewDebit = ProductDocumentDetailsRecordsetArg("NewDebit")
     VarOldDebit = Nz(ProductDocumentDetailsRecordsetArg("OldDebit"), 0)
     VarNewCredit = ProductDocumentDetailsRecordsetArg("NewCredit")
     VarOldCredit = Nz(ProductDocumentDetailsRecordsetArg("OldCredit"), 0)
     VarNewVatDebit = ProductDocumentDetailsRecordsetArg("NewVatDebit")
     VarOldVatDebit = Nz(ProductDocumentDetailsRecordsetArg("OldVatDebit"), 0)
     VarNewVatCredit = ProductDocumentDetailsRecordsetArg("NewVatCredit")
     VarOldVatCredit = Nz(ProductDocumentDetailsRecordsetArg("OldVatCredit"), 0)
     


'1. updating Financial Transactors TotalDebit and TotalCredit fields in the TransactorsT for the link financial transactors with the product
   'We check if there was any change in TransactorID of the record
    'if no, then we simply update total debit and total credit of the transactor
     If VarNewFinancialTransactorID = VarOldFinancialTransactorID Then
     
       If DocIsDeletedArg = True Then
       db.Execute "Update TransactorsT SET TransactorsT.Total_Debit = TransactorsT.Total_Debit - """ & VarFinTransactorDebitDifferenceToAddToTransactorDebitTotal & _
       """ , TransactorsT.Total_Credit = TransactorsT.Total_Credit - """ & VarFinTransactorCreditDifferenceToAddToTransactorCreditTotal & _
       """ WHERE Transactor_ID = " & VarNewFinancialTransactorID, dbFailOnError
       Else
       db.Execute "Update TransactorsT SET TransactorsT.Total_Debit = TransactorsT.Total_Debit + """ & VarFinTransactorDebitDifferenceToAddToTransactorDebitTotal & _
       """ , TransactorsT.Total_Credit = TransactorsT.Total_Credit + """ & VarFinTransactorCreditDifferenceToAddToTransactorCreditTotal & _
       """ WHERE Transactor_ID = " & VarNewFinancialTransactorID, dbFailOnError
       End If
     Else
     'if yes, then we proceed updating Total Debit and Total Credit for both transactors respectively (for the new transactor we update using NewDebit and NewCredit,
     'while for the old transactor (backedup) we use OldDebit and OldCredit
              
       db.Execute "Update TransactorsT SET TransactorsT.Total_Debit = TransactorsT.Total_Debit + """ & VarNewDebit & _
       """ , TransactorsT.Total_Credit = TransactorsT.Total_Credit + """ & VarNewCredit & """ WHERE Transactor_ID = " & VarNewFinancialTransactorID, dbFailOnError
      
       db.Execute "Update TransactorsT SET TransactorsT.Total_Debit = TransactorsT.Total_Debit - """ & VarOldDebit & _
       """ , TransactorsT.Total_Credit = TransactorsT.Total_Credit - """ & VarOldCredit & _
       """ WHERE Transactor_ID = " & VarOldFinancialTransactorID, dbFailOnError

      End If
      
'2. updating Financial Transactors TotalDebit and TotalCredit fields in the TransactorsT for the link Vat transactors with the product.
   'We check if there was any change in VatTransactorID of the record
    'if no, then we simply update total debit and total credit of the transactor
      If VarNewVatTransactorID = Nz(VarOldVatTransactorID, VarNewVatTransactorID) Then
       
       If DocIsDeletedArg = True Then
       db.Execute "Update TransactorsT SET TransactorsT.Total_Debit = TransactorsT.Total_Debit - """ & VarVatTransactorDebitDifferenceToAddToTransactorDebitTotal & _
       """ , TransactorsT.Total_Credit = TransactorsT.Total_Credit - """ & VarVatTransactorCreditDifferenceToAddToTransactorCreditTotal & _
       """ WHERE Transactor_ID = " & VarNewVatTransactorID, dbFailOnError
       Else
       db.Execute "Update TransactorsT SET TransactorsT.Total_Debit = TransactorsT.Total_Debit + """ & VarVatTransactorDebitDifferenceToAddToTransactorDebitTotal & _
       """ , TransactorsT.Total_Credit = TransactorsT.Total_Credit + """ & VarVatTransactorCreditDifferenceToAddToTransactorCreditTotal & _
       """ WHERE Transactor_ID = " & VarNewVatTransactorID, dbFailOnError
       End If
     Else
     'if yes, then we proceed updating Total Debit and Total Credit for both transactors respectively (for the new transactor we update using NewDebit and NewCredit,
     'while for the old transactor (backedup) we use OldDebit and OldCredit
              
       db.Execute "Update TransactorsT SET TransactorsT.Total_Debit = TransactorsT.Total_Debit + """ & VarNewVatDebit & _
       """ , TransactorsT.Total_Credit = TransactorsT.Total_Credit + """ & VarNewVatCredit & _
       """ WHERE Transactor_ID = " & VarNewVatTransactorID, dbFailOnError
      
       CurrentDb.Execute "Update TransactorsT SET TransactorsT.Total_Debit = TransactorsT.Total_Debit - """ & VarOldVatDebit & _
       """ , TransactorsT.Total_Credit = TransactorsT.Total_Credit - """ & VarOldVatCredit & _
       """ WHERE Transactor_ID = " & VarOldVatTransactorID, dbFailOnError

      End If
    End If
 '___We take care updating the inventory
   If VarAffectsInventoy = 1 Then
     VarQuantityDifferenceToAddToInventoryDebitTotal = ProductDocumentDetailsRecordsetArg("QuantityDebitToAdd")
     VarQuantityDifferenceToAddToInventoryCreditTotal = ProductDocumentDetailsRecordsetArg("QuantityCreditToAdd")
     VarNewProductID = ProductDocumentDetailsRecordsetArg("NewProductID")
     VarOldProductID = Nz(ProductDocumentDetailsRecordsetArg("OldProductID"), VarNewProductID)
     VarNewMainEntityID = ProductDocumentDetailsRecordsetArg("NewMainTransactorID")
     VarOldMainEntityID = Nz(ProductDocumentDetailsRecordsetArg("OldMainTransactorID"), VarNewMainEntityID)
     VarNewOtherEntityID = ProductDocumentDetailsRecordsetArg("NewOtherTransactorID")
     VarOldOtherEntityID = Nz(ProductDocumentDetailsRecordsetArg("OldOtherTransactorID"), VarNewOtherEntityID)
     VarNewQuantityDebit = ProductDocumentDetailsRecordsetArg("NewQuantityDebit")
     VarOldQuantityDebit = ProductDocumentDetailsRecordsetArg("OldQuantityDebit")
     VarNewQuantityCredit = ProductDocumentDetailsRecordsetArg("NewQuantityCredit")
     VarOldQuantityCredit = Nz(ProductDocumentDetailsRecordsetArg("OldQuantityCredit"), 0) 'for an unknown reason, the DebitAndCreditForTransactorsFromProductDetailsRecordsQ when executed by itself brings 0 in field "OldQuantityCredit" but when you just iterate it, it brings null!!! I have checked it quite deeply and nothing is wrong. It just disbehaves! That is the reason I use the NZ function, in order to bypass the null error it raises.
     
'3. updating ProductInventoryBalanceT TotalDebit and TotalCredit fields for the specific product with the specific transactor_ID (main entity).
   'first we check if there is a record in the table ProductInventoryBalanceT which corresponds to the specific ProductID and TransactorID.
   'if there is not any records, then we insert one in order to be present for next rutines to update totaldebit and totalcredit fields of this record.
       
      If IsNull(DLookup("ProductInventoryBalanceT.Product_Inventory_Balance_ID", "ProductInventoryBalanceT", "Product_ID = " & VarNewProductID & " AND Transactor_ID = " & VarNewMainEntityID)) Then
      db.Execute "Insert Into ProductInventoryBalanceT (Product_ID, Transactor_ID) Values (" & VarNewProductID & " , " & VarNewMainEntityID & ")", dbFailOnError
      End If
    
    'We check if there was any change in ProductID or MainTransactorID of the record
    'if no, then we simply update total debit and total credit of the relative record of ProductInventoryBalanceT
      If VarNewProductID = VarOldProductID And VarNewMainEntityID = VarOldMainEntityID Then
      
       If DocIsDeletedArg = True Then
       db.Execute "Update ProductInventoryBalanceT SET ProductInventoryBalanceT.Total_Debit = ProductInventoryBalanceT.Total_Debit - """ & VarQuantityDifferenceToAddToInventoryDebitTotal & _
       """ , ProductInventoryBalanceT.Total_Credit = ProductInventoryBalanceT.Total_Credit - """ & VarQuantityDifferenceToAddToInventoryCreditTotal & _
       """ WHERE Product_ID = " & VarNewProductID & "and Transactor_ID = " & VarNewMainEntityID, dbFailOnError
       Else
       db.Execute "Update ProductInventoryBalanceT SET ProductInventoryBalanceT.Total_Debit = ProductInventoryBalanceT.Total_Debit + """ & VarQuantityDifferenceToAddToInventoryDebitTotal & _
       """ , ProductInventoryBalanceT.Total_Credit = ProductInventoryBalanceT.Total_Credit + """ & VarQuantityDifferenceToAddToInventoryCreditTotal & _
       """ WHERE Product_ID = " & VarNewProductID & "and Transactor_ID = " & VarNewMainEntityID, dbFailOnError
       End If
      Else
     'if yes, then we proceed updating Total Debit and Total Credit for both records of ProductInventoryBalanceT respectively, for the new record (pair of Product_ID and Transactor_ID) we update using NewDebit and NewCredit, while for the old record we use OldDebit and OldCredit
              
       db.Execute "Update ProductInventoryBalanceT SET ProductInventoryBalanceT.Total_Debit = ProductInventoryBalanceT.Total_Debit + """ & VarNewQuantityDebit & _
       """ , ProductInventoryBalanceT.Total_Credit = ProductInventoryBalanceT.Total_Credit + """ & VarNewQuantityCredit & _
       """ WHERE Product_ID = " & VarNewProductID & "and Transactor_ID = " & VarNewMainEntityID, dbFailOnError
      
       db.Execute "Update ProductInventoryBalanceT SET ProductInventoryBalanceT.Total_Debit = ProductInventoryBalanceT.Total_Debit - """ & VarOldQuantityDebit & _
       """ , ProductInventoryBalanceT.Total_Credit = ProductInventoryBalanceT.Total_Credit - """ & VarOldQuantityCredit & _
       """ WHERE Product_ID = " & VarOldProductID & "and Transactor_ID = " & VarOldMainEntityID, dbFailOnError
      End If
      
'4. updating ProductInventoryBalanceT TotalDebit and TotalCredit fields for the specific product with the specific transactor_ID (Other entity).
   'ATTENTION!!! We use VarQuantityDifferenceToAddToInventoryDebitTotal,VarNewQuantityDebit and VarOldQuantityDebit in opposite way to case 3 (above),
   'as it affects other entity, which has to be updated opposite main entity (updated in case 3)
   'We check if there was any change in ProductID or OtherTransactorID of the record
    'if no, then we simply update total debit and total credit of the relative record of ProductInventoryBalanceT
    
      If IsNull(DLookup("ProductInventoryBalanceT.Product_Inventory_Balance_ID", "ProductInventoryBalanceT", "Product_ID = " & VarNewProductID & " AND Transactor_ID = " & VarNewOtherEntityID)) Then
      db.Execute "Insert Into ProductInventoryBalanceT (Product_ID, Transactor_ID) Values (" & VarNewProductID & " , " & VarNewOtherEntityID & ")", dbFailOnError
      End If
      
      If VarNewProductID = VarOldProductID And VarNewOtherEntityID = VarOldOtherEntityID Then
      
       If DocIsDeletedArg = True Then
       db.Execute "Update ProductInventoryBalanceT SET ProductInventoryBalanceT.Total_Debit = ProductInventoryBalanceT.Total_Debit - """ & VarQuantityDifferenceToAddToInventoryCreditTotal & _
       """ , ProductInventoryBalanceT.Total_Credit = ProductInventoryBalanceT.Total_Credit - """ & VarQuantityDifferenceToAddToInventoryDebitTotal & _
       """ WHERE Product_ID = " & VarNewProductID & "and Transactor_ID = " & VarNewOtherEntityID, dbFailOnError
       Else
       db.Execute "Update ProductInventoryBalanceT SET ProductInventoryBalanceT.Total_Debit = ProductInventoryBalanceT.Total_Debit + """ & VarQuantityDifferenceToAddToInventoryCreditTotal & _
       """ , ProductInventoryBalanceT.Total_Credit = ProductInventoryBalanceT.Total_Credit + """ & VarQuantityDifferenceToAddToInventoryDebitTotal & _
       """ WHERE Product_ID = " & VarNewProductID & "and Transactor_ID = " & VarNewOtherEntityID, dbFailOnError
       End If
      Else
     'if yes, then we proceed updating Total Debit and Total Credit for both records of ProductInventoryBalanceT respectively, for the new record (pair of Product_ID and Transactor_ID) we update using NewDebit and NewCredit, while for the old record we use OldDebit and OldCredit
              
       db.Execute "Update ProductInventoryBalanceT SET ProductInventoryBalanceT.Total_Debit = ProductInventoryBalanceT.Total_Debit + """ & VarNewQuantityCredit & _
       """ , ProductInventoryBalanceT.Total_Credit = ProductInventoryBalanceT.Total_Credit + """ & VarNewQuantityDebit & _
       """ WHERE Product_ID = " & VarNewProductID & "and Transactor_ID = " & VarNewOtherEntityID, dbFailOnError
      
       db.Execute "Update ProductInventoryBalanceT SET ProductInventoryBalanceT.Total_Debit = ProductInventoryBalanceT.Total_Debit - """ & VarOldQuantityCredit & _
       """ , ProductInventoryBalanceT.Total_Credit = ProductInventoryBalanceT.Total_Credit - """ & VarOldQuantityDebit & _
       """ WHERE Product_ID = " & VarOldProductID & "and Transactor_ID = " & VarOldOtherEntityID, dbFailOnError
      End If
    End If
    
      

   
       ProductDocumentDetailsRecordsetArg.MoveNext
   Loop
   
End If
 
UpdateTransactorsBalanceAndInventoryBalanceByProductDocumentDetailsRecordset = True


ExitProcedure:

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
            "Error Source: UpdateTransactorsBalanceAndInventoryBalanceByProductDocumentDetailsRecordset " & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            UpdateTransactorsBalanceAndInventoryBalanceByProductDocumentDetailsRecordset = False
            Resume ExitProcedure
            End Select
End Function

Public Sub UpdateFinancialAndVatTransactorFieldsForAllProductDocumentDetails(DocumentIDArg, IntentionIDArg)
Debug.Print "Module Public Functions - " & "UpdateFinancialAndVatTransactorFieldsForAllProductDocumentDetails " & Time()
'On Error GoTo Errorhandler
'This procedure checks if new intention affects financial. If yes then it itterated all product details and fills "Financial_Transactor_ID" and "Vat_Transactor_ID" fields regardless they are filled or not.

Dim db As DAO.Database
Dim rstProductDetails As DAO.Recordset
Dim VarFinancialTransactorID As Long
Dim VarVatTransactorID As Long
Dim vatValue As String
Dim VarIntentionAffectsFinancial As Boolean

Set db = CurrentDb
Set rstProductDetails = db.OpenRecordset("select Issued_Document_Product_Details_ID, Issued_Document_ID, Product_ID, [VAT%], Accounting_Behavior_ID, Financial_Transactor_ID, Vat_Transactor_ID " & _
"from IssuedDocumentProductDetailsT where Issued_Document_ID = " & DocumentIDArg)

VarIntentionAffectsFinancial = DLookup("Affects_Financial", "IntentionsT", "Intention_ID = " & IntentionIDArg)


If Not rstProductDetails.EOF Then
rstProductDetails.MoveLast
rstProductDetails.MoveFirst

Do Until rstProductDetails.EOF

 VarFinancialTransactorID = Nz(DLookup("Transactor_Financial_ID", "LinkProductsToCompanyFinancialTransactorsT", "Product_ID = " & rstProductDetails("Product_ID") & " AND Intention_Type_ID = " & VarIntentionTypeID & _
   " AND Accounting_Behavior_ID = " & Nz(rstProductDetails("Accounting_Behavior_ID"), 0)), 0)

vatValue = Replace(rstProductDetails("VAT%"), ",", ".")

VarVatTransactorID = Nz(DLookup("Vat_Transactor_ID", "LinkFinTransactorsToVatTransactorsWithVat%Q", "Financial_Transactor_ID = " & VarFinancialTransactorID & _
   " AND [Vat%] = " & vatValue), 0)

rstProductDetails.Edit
If VarIntentionAffectsFinancial = True Then
rstProductDetails("Financial_Transactor_ID") = VarFinancialTransactorID
rstProductDetails("Vat_Transactor_ID") = VarVatTransactorID
Else
rstProductDetails("Financial_Transactor_ID") = Null
rstProductDetails("Vat_Transactor_ID") = Null
End If
rstProductDetails.Update
   
rstProductDetails.MoveNext
Loop

End If
ExitProcedure:

If Not rstProductDetails Is Nothing Then
rstProductDetails.Close
Set rstProductDetails = Nothing
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
            "Error Source: UpdateFinancialAndVatTransactorFieldsForAllProductDocumentDetails " & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
End Sub

Public Function ProductDocumentTotalAmount(IssuedDocumentIDArg As Long) As Double
Debug.Print "Module Public Functions - " & "ProductDocumentTotalAmount " & Time()
'On Error GoTo Errorhandler

Dim db As DAO.Database
Dim rstProductDetailsValuesGroupedByDocumentID As DAO.Recordset

ProductDocumentTotalAmount = 0

Set db = CurrentDb
Set rstProductDetailsValuesGroupedByDocumentID = db.OpenRecordset("SELECT Sum(FinancialDetailsFromProductDetailsQ.Total_Gross_Value) AS SumOfTotal_Gross_Value, " & _
"FinancialDetailsFromProductDetailsQ.Issued_Document_ID" & _
"FROM FinancialDetailsFromProductDetailsQ " & _
"GROUP BY FinancialDetailsFromProductDetailsQ.Issued_Document_ID " & _
"HAVING FinancialDetailsFromProductDetailsQ.Issued_Document_ID)= " & IssuedDocumentIDArg)

ProductDocumentTotalAmount = rstProductDetailsValuesGroupedByDocumentID(0)

ExitProcedure:

If Not rstProductDetails Is Nothing Then
rstProductDetails.Close
Set rstProductDetails = Nothing
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
            "Error Source: ProductDocumentTotalAmount " & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
            
End Function

Public Function InventorySumForOneProduct(ProductIDArg As Long, ByRef TotalQuantityDebitedArg As Double, ByRef TotalQuantityCreditedArg As Double) As Double
Debug.Print "Module Public Functions - " & "InventorySumForOneProduct " & Time()
'On Error GoTo Errorhandler

Dim db As DAO.Database
Dim InventorySumForOneProductQdef As QueryDef
Dim rstInventorySumForOneProduct As DAO.Recordset

InventorySumForOneProduct = 0

Set db = CurrentDb

Set InventorySumForOneProductQdef = db.QueryDefs("InventorySumPerProductQ")
 InventorySumForOneProductQdef.Parameters("ParProductID").Value = ProductIDArg

Set rstInventorySumForOneProduct = InventorySumForOneProductQdef.OpenRecordset()

If Not rstInventorySumForOneProduct.EOF Then
 rstInventorySumForOneProduct.MoveLast
 rstInventorySumForOneProduct.MoveFirst

 TotalQuantityDebitedArg = rstInventorySumForOneProduct("SumOfTotal_Debit")
 TotalQuantityCreditedArg = rstInventorySumForOneProduct("SumOfTotal_Credit")
 InventorySumForOneProduct = TotalQuantityDebitedArg - TotalQuantityCreditedArg    'rstInventorySumForOneProduct("SumOfInventory_Per_Product")
End If

ExitProcedure:

If Not InventorySumForOneProductQdef Is Nothing Then
 InventorySumForOneProductQdef.Close
 Set InventorySumForOneProductQdef = Nothing
End If

If Not rstInventorySumForOneProduct Is Nothing Then
 rstInventorySumForOneProduct.Close
 Set rstInventorySumForOneProduct = Nothing
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
            "Error Source: InventorySumForOneProduct " & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
End Function
Public Function CheckPrecalculatedFinancialBalance(ByRef PrecalculatedTotalDebitArg As Double, ByRef PrecalculatedTotalCreditArg As Double, Optional TransactionIDArg As Long, Optional DocumentIDArg As Long, Optional FinancialDetailsIDArg As Long, Optional ProductDetailsIDArg As Long) As Double
Debug.Print "Module Public Functions - " & "CheckPrecalculatedFinancialBalance " & Time()
'On Error GoTo Errorhandler

Dim db As DAO.Database
Dim rstPrecalculatedFinalTotalBalance As DAO.Recordset
Dim rstForLastRecordInFinancialBalanceErrorsTable As DAO.Recordset
Dim SQLstr As String
Dim VarNotesStr As String
Dim VarPrecalculatedTotalDebit As Double
Dim VarPrecalculatedTotalCredit As Double
Dim VarRecordsetRecordCount As String
Dim VarLastRecordIDInFinancialBalanceErrorsTable As Long

Set db = CurrentDb
Set rstPrecalculatedFinalTotalBalance = db.OpenRecordset("select sum(Total_Debit), sum(Total_Credit) from TransactorsT")

If Not rstPrecalculatedFinalTotalBalance.EOF Then
    If rstPrecalculatedFinalTotalBalance.recordcount = 1 Then
      rstPrecalculatedFinalTotalBalance.MoveLast
      rstPrecalculatedFinalTotalBalance.MoveFirst
      VarPrecalculatedTotalDebit = Round(rstPrecalculatedFinalTotalBalance(0), 2)
      VarPrecalculatedTotalCredit = Round(rstPrecalculatedFinalTotalBalance(1), 2)
      CheckPrecalculatedFinancialBalance = VarPrecalculatedTotalDebit - VarPrecalculatedTotalCredit
      PrecalculatedTotalDebitArg = Round(VarPrecalculatedTotalDebit, 2)
      PrecalculatedTotalCreditArg = Round(VarPrecalculatedTotalCredit, 2)
      If CheckPrecalculatedFinancialBalance <> 0 Then
      MsgBox "Financial imbalance detected in saved precalculations! Check if the last document you tried to enter, did update precalculations partialy and caused this imbalance. If this is the case, try to correct it. If problem persists, please contact IT department.", vbExclamation + vbOKOnly
      VarNotesStr = ""
      SQLstr = SQLStringConstructionForInsertIntoFinancialBalanceErrorsTable(CDbl(CheckPrecalculatedFinancialBalance), VarPrecalculatedTotalDebit, VarPrecalculatedTotalCredit, TransactionIDArg, DocumentIDArg, FinancialDetailsIDArg, ProductDetailsIDArg, VarNotesStr)
      db.Execute (SQLstr), dbFailOnError
      GoTo ExitProcedure
      End If
    Else
      VarRecordsetRecordCount = rstPrecalculatedFinalTotalBalance.recordcount
      VarNotesStr = "TestQueryForFinancialBalanceQ brought " & VarRecordsetRecordCount & " records, while it should return only one."
      SQLstr = SQLStringConstructionForInsertIntoFinancialBalanceErrorsTable(, , , TransactionIDArg, DocumentIDArg, FinancialDetailsIDArg, ProductDetailsIDArg, VarNotesStr)
      db.Execute (SQLstr), dbFailOnError
      Set rstForLastRecordInFinancialBalanceErrorsTable = db.OpenRecordset("SELECT @@IDENTITY")
      VarLastRecordIDInFinancialBalanceErrorsTable = rstForLastRecordInFinancialBalanceErrorsTable(0)
      MsgBox "Balance Query did not execute succesfully. Please contact IT and provide following info ""Error Record = " & VarLastRecordIDInFinancialBalanceErrorsTable & " "".", vbOKOnly
      GoTo ExitProcedure
    End If
Else
   VarNotesStr = "TestQueryForFinancialBalanceQ brought no records."
   SQLstr = SQLStringConstructionForInsertIntoFinancialBalanceErrorsTable(, , , TransactionIDArg, DocumentIDArg, FinancialDetailsIDArg, ProductDetailsIDArg, VarNotesStr)
   db.Execute (SQLstr)
   Set rstForLastRecordInFinancialBalanceErrorsTable = db.OpenRecordset("SELECT @@IDENTITY")
   VarLastRecordIDInFinancialBalanceErrorsTable = rstForLastRecordInFinancialBalanceErrorsTable(0)
   MsgBox "A problem occured with the CheckPrecalculatedFinancialBalance Function. Please contact IT and provide following info ""Error Record = " & VarLastRecordIDInFinancialBalanceErrorsTable & " "".", vbOKOnly
   GoTo ExitProcedure
End If

ExitProcedure:

If Not rstForLastRecordInFinancialBalanceErrorsTable Is Nothing Then
rstForLastRecordInFinancialBalanceErrorsTable.Close
Set rstForLastRecordInFinancialBalanceErrorsTable = Nothing
End If

If Not rstPrecalculatedFinalTotalBalance Is Nothing Then
rstPrecalculatedFinalTotalBalance.Close
Set rstPrecalculatedFinalTotalBalance = Nothing
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
            "Error Source: CheckPrecalculatedFinancialBalance " & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select



End Function
Public Function CheckPrecalculatedInventoryBalance(ByRef PrecalculatedTotalDebitArg As Double, ByRef PrecalculatedTotalCreditArg As Double, Optional TransactionIDArg As Long, Optional DocumentIDArg As Long, Optional ProductDetailsIDArg As Long) As Double
Debug.Print "Module Public Functions - " & "CheckPrecalculatedInventoryBalance " & Time()
'On Error GoTo Errorhandler

Dim db As DAO.Database
Dim rstPrecalculatedFinalTotalBalance As DAO.Recordset
Dim rstForLastRecordInInventoryBalanceErrorsTable As DAO.Recordset
Dim SQLstr As String
Dim VarNotesStr As String
Dim VarPrecalculatedTotalDebit As Double
Dim VarPrecalculatedTotalCredit As Double
Dim VarRecordsetRecordCount As String
Dim VarLastRecordIDInInventoryBalanceErrorsTable As Long

Set db = CurrentDb
Set rstPrecalculatedFinalTotalBalance = db.OpenRecordset("select sum(Total_Debit), sum(Total_Credit) from ProductInventoryBalanceT")

If Not rstPrecalculatedFinalTotalBalance.EOF Then
    If rstPrecalculatedFinalTotalBalance.recordcount = 1 Then
      rstPrecalculatedFinalTotalBalance.MoveLast
      rstPrecalculatedFinalTotalBalance.MoveFirst
      VarPrecalculatedTotalDebit = Round(Nz(rstPrecalculatedFinalTotalBalance(0), 0), 3)
      VarPrecalculatedTotalCredit = Round(Nz(rstPrecalculatedFinalTotalBalance(1), 0), 3)
      CheckPrecalculatedInventoryBalance = VarPrecalculatedTotalDebit - VarPrecalculatedTotalCredit
      PrecalculatedTotalDebitArg = VarPrecalculatedTotalDebit
      PrecalculatedTotalCreditArg = VarPrecalculatedTotalCredit
      If CheckPrecalculatedInventoryBalance <> 0 Then
      MsgBox "Inventory imbalance detected in saved precalculations! Check if the last document you tried to enter, did update precalculations partialy and caused this imbalance. If this is the case, try to correct it by deleting the document and enter it again from the beginning. If problem persists, please contact IT department.", vbExclamation + vbOKOnly
      VarNotesStr = ""
      SQLstr = SQLStringConstructionForInsertIntoInventoryBalanceErrorsTable(CDbl(CheckPrecalculatedInventoryBalance), VarPrecalculatedTotalDebit, VarPrecalculatedTotalCredit, TransactionIDArg, DocumentIDArg, ProductDetailsIDArg, VarNotesStr)
      db.Execute (SQLstr), dbFailOnError
      GoTo ExitProcedure
      End If
    Else
      VarRecordsetRecordCount = rstPrecalculatedFinalTotalBalance.recordcount
      VarNotesStr = "TestQueryForInventoryBalanceQ brought " & VarRecordsetRecordCount & " records, while it should return only one."
      SQLstr = SQLStringConstructionForInsertIntoInventoryBalanceErrorsTable(, , , TransactionIDArg, DocumentIDArg, ProductDetailsIDArg, VarNotesStr)
      db.Execute (SQLstr), dbFailOnError
      Set rstForLastRecordInInventoryBalanceErrorsTable = db.OpenRecordset("SELECT @@IDENTITY")
      VarLastRecordIDInInventoryBalanceErrorsTable = rstForLastRecordInInventoryBalanceErrorsTable(0)
      MsgBox "Balance Query did not execute succesfully. Please contact IT and provide following info ""Error Record = " & VarLastRecordIDInInventoryBalanceErrorsTable & " "".", vbOKOnly
      GoTo ExitProcedure
    End If
Else
   VarNotesStr = "TestQueryForInventoryBalanceQ brought no records."
   SQLstr = SQLStringConstructionForInsertIntoInventoryBalanceErrorsTable(, , , TransactionIDArg, DocumentIDArg, ProductDetailsIDArg, VarNotesStr)
   db.Execute (SQLstr)
   Set rstForLastRecordInInventoryBalanceErrorsTable = db.OpenRecordset("SELECT @@IDENTITY")
   VarLastRecordIDInInventoryBalanceErrorsTable = rstForLastRecordInInventoryBalanceErrorsTable(0)
   MsgBox "A problem occured with the CheckPrecalculatedInventoryBalance Function. Please contact IT and provide following info ""Error Record = " & VarLastRecordIDInInventoryBalanceErrorsTable & " "".", vbOKOnly
   GoTo ExitProcedure
End If

ExitProcedure:

If Not rstForLastRecordInInventoryBalanceErrorsTable Is Nothing Then
rstForLastRecordInInventoryBalanceErrorsTable.Close
Set rstForLastRecordInInventoryBalanceErrorsTable = Nothing
End If

If Not rstPrecalculatedFinalTotalBalance Is Nothing Then
rstPrecalculatedFinalTotalBalance.Close
Set rstPrecalculatedFinalTotalBalance = Nothing
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
            "Error Source: CheckPrecalculatedInventoryBalance " & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select

End Function
Public Function CheckInventoryBalance(ByRef TotalInventoryDebitArg As Double, ByRef TotalInventoryCreditArg As Double, Optional ByRef DifferenceArg As Double, Optional TransactionIDArg As Long, Optional DocumentIDArg As Long, Optional ProductDetailsIDArg As Long) As Double
Debug.Print "Module Public Functions - " & "CheckInventoryBalance " & Time()
'On Error GoTo Errorhandler

Dim db As DAO.Database
Dim rstFinalTotalInventoryBalance As DAO.Recordset
Dim rstForLastRecordInInventoryBalanceErrorsTable As DAO.Recordset
Dim SQLstr As String
Dim VarNotesStr As String
Dim VarTotalInventoryDebit As Double
Dim VarTotalInventoryCredit As Double
Dim VarRecordsetRecordCount As String
Dim VarLastRecordIDInInventoryBalanceErrorsTable As Long

Set db = CurrentDb
Set rstFinalTotalInventoryBalance = db.OpenRecordset("select * from TestQueryForInventoryBalanceQ")

If Not rstFinalTotalInventoryBalance.EOF Then
    If rstFinalTotalInventoryBalance.recordcount = 1 Then
      rstFinalTotalInventoryBalance.MoveLast
      rstFinalTotalInventoryBalance.MoveFirst
      VarTotalInventoryDebit = Round(Nz(rstFinalTotalInventoryBalance("SumOfDebit"), 0), 2)
      VarTotalInventoryCredit = Round(Nz(rstFinalTotalInventoryBalance("SumOfCredit"), 0), 2)
      CheckInventoryBalance = VarTotalInventoryDebit - VarTotalInventoryCredit
      TotalInventoryDebitArg = VarTotalInventoryDebit
      TotalInventoryCreditArg = VarTotalInventoryCredit
      DifferenceArg = TotalInventoryDebitArg - TotalInventoryCreditArg
      If DifferenceArg <> 0 Then
      MsgBox "Inventory imbalance detected! Check if the last document you tried to enter was partialy entered and caused the imbalance. If this is the case, try to correct it. If problem persists, please contact IT department.", vbExclamation + vbOKOnly
      VarNotesStr = ""
      SQLstr = SQLStringConstructionForInsertIntoInventoryBalanceErrorsTable(CheckInventoryBalance, VarTotalInventoryDebit, VarTotalInventoryCredit, TransactionIDArg, DocumentIDArg, ProductDetailsIDArg, VarNotesStr)
      db.Execute (SQLstr), dbFailOnError
      GoTo ExitProcedure
      End If
    Else
      VarRecordsetRecordCount = rstFinalTotalInventoryBalance.recordcount
      VarNotesStr = "TestQueryForInventoryBalanceQ brought " & VarRecordsetRecordCount & " records, while it should return only one."
      SQLstr = SQLStringConstructionForInsertIntoInventoryBalanceErrorsTable(, , , TransactionIDArg, DocumentIDArg, ProductDetailsIDArg, VarNotesStr)
      db.Execute (SQLstr), dbFailOnError
      Set rstForLastRecordInInventoryBalanceErrorsTable = db.OpenRecordset("SELECT @@IDENTITY")
      VarLastRecordIDInInventoryBalanceErrorsTable = rstForLastRecordInInventoryBalanceErrorsTable(0)
      MsgBox "Balance Query did not execute succesfully. Please contact IT and provide following info ""Error Record = " & VarLastRecordIDInInventoryBalanceErrorsTable & " "".", vbOKOnly
      GoTo ExitProcedure
    End If
Else
   VarNotesStr = "TestQueryForInventoryBalanceQ brought no records."
   SQLstr = SQLStringConstructionForInsertIntoInventoryBalanceErrorsTable(, , , TransactionIDArg, DocumentIDArg, ProductDetailsIDArg, VarNotesStr)
   db.Execute (SQLstr)
   Set rstForLastRecordInInventoryBalanceErrorsTable = db.OpenRecordset("SELECT @@IDENTITY")
   VarLastRecordIDInInventoryBalanceErrorsTable = rstForLastRecordInInventoryBalanceErrorsTable(0)
   MsgBox "A problem occured with the CheckInventoryBalance Function. Please contact IT and provide following info ""Error Record = " & VarLastRecordIDInInventoryBalanceErrorsTable & " "".", vbOKOnly
   GoTo ExitProcedure
End If

ExitProcedure:

If Not rstForLastRecordInInventoryBalanceErrorsTable Is Nothing Then
rstForLastRecordInInventoryBalanceErrorsTable.Close
Set rstForLastRecordInInventoryBalanceErrorsTable = Nothing
End If

If Not rstFinalTotalInventoryBalance Is Nothing Then
rstFinalTotalInventoryBalance.Close
Set rstFinalTotalInventoryBalance = Nothing
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
            "Error Source: CheckInventoryBalance " & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select


End Function
Public Function CheckFinancialBalance(ByRef TotalDebitArg As Double, ByRef TotalCreditArg As Double, Optional ByRef DifferenceArg As Double, Optional TransactionIDArg As Long, Optional DocumentIDArg As Long, Optional FinancialDetailsIDArg As Long, Optional ProductDetailsIDArg As Long) As Double
Debug.Print "Module Public Functions - " & "CheckFinancialBalance " & Time()
'On Error GoTo Errorhandler

Dim db As DAO.Database
Dim rstFinalTotalBalance As DAO.Recordset
Dim rstForLastRecordInFinancialBalanceErrorsTable As DAO.Recordset
Dim SQLstr As String
Dim VarNotesStr As String
Dim VarTotalDebit As Double
Dim VarTotalCredit As Double
Dim VarRecordsetRecordCount As String
Dim VarLastRecordIDInFinancialBalanceErrorsTable As Long

Set db = CurrentDb
Set rstFinalTotalBalance = db.OpenRecordset("select * from TestQueryForFinancialBalanceQ")

If Not rstFinalTotalBalance.EOF Then
    If rstFinalTotalBalance.recordcount = 1 Then
      rstFinalTotalBalance.MoveLast
      rstFinalTotalBalance.MoveFirst
      VarTotalDebit = Round(Nz(rstFinalTotalBalance("SumOfDebit"), 0), 2)
      VarTotalCredit = Round(Nz(rstFinalTotalBalance("SumOfCredit"), 0), 2)
      TotalDebitArg = VarTotalDebit
      TotalCreditArg = VarTotalCredit
      DifferenceArg = TotalDebitArg - TotalCreditArg
      CheckFinancialBalance = DifferenceArg
      If DifferenceArg <> 0 Then
      MsgBox "Financial imbalance detected! Check if the last document you tried to enter was partialy entered and caused the imbalance. If this is the case, try to correct it. If problem persists, please contact IT department.", vbExclamation + vbOKOnly
      VarNotesStr = ""
      SQLstr = SQLStringConstructionForInsertIntoFinancialBalanceErrorsTable(CheckFinancialBalance, VarTotalDebit, VarTotalCredit, TransactionIDArg, DocumentIDArg, FinancialDetailsIDArg, ProductDetailsIDArg, VarNotesStr)
      db.Execute (SQLstr), dbFailOnError
      GoTo ExitProcedure
      End If
    Else
      VarRecordsetRecordCount = rstFinalTotalBalance.recordcount
      VarNotesStr = "TestQueryForFinancialBalanceQ brought " & VarRecordsetRecordCount & " records, while it should return only one."
      SQLstr = SQLStringConstructionForInsertIntoFinancialBalanceErrorsTable(, , , TransactionIDArg, DocumentIDArg, FinancialDetailsIDArg, ProductDetailsIDArg, VarNotesStr)
      db.Execute (SQLstr), dbFailOnError
      Set rstForLastRecordInFinancialBalanceErrorsTable = db.OpenRecordset("SELECT @@IDENTITY")
      VarLastRecordIDInFinancialBalanceErrorsTable = rstForLastRecordInFinancialBalanceErrorsTable(0)
      MsgBox "Balance Query did not execute succesfully. Please contact IT and provide following info ""Error Record = " & VarLastRecordIDInFinancialBalanceErrorsTable & " "".", vbOKOnly
      GoTo ExitProcedure
    End If
Else
   VarNotesStr = "TestQueryForFinancialBalanceQ brought no records."
   SQLstr = SQLStringConstructionForInsertIntoFinancialBalanceErrorsTable(, , , TransactionIDArg, DocumentIDArg, FinancialDetailsIDArg, ProductDetailsIDArg, VarNotesStr)
   db.Execute (SQLstr)
   Set rstForLastRecordInFinancialBalanceErrorsTable = db.OpenRecordset("SELECT @@IDENTITY")
   VarLastRecordIDInFinancialBalanceErrorsTable = rstForLastRecordInFinancialBalanceErrorsTable(0)
   MsgBox "A problem occured with the CheckFinancialBalance Function. Please contact IT and provide following info ""Error Record = " & VarLastRecordIDInFinancialBalanceErrorsTable & " "".", vbOKOnly
   GoTo ExitProcedure
End If

ExitProcedure:

If Not rstForLastRecordInFinancialBalanceErrorsTable Is Nothing Then
rstForLastRecordInFinancialBalanceErrorsTable.Close
Set rstForLastRecordInFinancialBalanceErrorsTable = Nothing
End If

If Not rstFinalTotalBalance Is Nothing Then
rstFinalTotalBalance.Close
Set rstFinalTotalBalance = Nothing
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
            "Error Source: CheckFinancialBalance " & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select


End Function

Function SQLStringConstructionForInsertIntoFinancialBalanceErrorsTable(Optional DifferenceArg As Double, Optional TotalDebitArg As Double, Optional TotalCreditArg As Double, Optional TransactionIDArg As Long, Optional DocumentIDArg As Long, Optional FinancialDetailsIDArg As Long, Optional ProductDetailsIDArg As Long, Optional NotesArg As String) As String
Debug.Print "Module Public Functions - " & "SQLStringConstructionForInsertIntoFinancialBalanceErrorsTable " & Time()
On Error GoTo ErrorHandler

Dim SQLstr As String

SQLstr = "Insert into FinancialBalanceErrorsT ("

If Not IsNull(TransactionIDArg) Then
    SQLstr = SQLstr & "Transaction_ID, "
End If

If Not IsNull(DocumentIDArg) Then
    SQLstr = SQLstr & "Document_ID, "
End If

If Not IsNull(FinancialDetailsIDArg) Then
    SQLstr = SQLstr & "Financial_Details_ID, "
End If

If Not IsNull(ProductDetailsIDArg) Then
    SQLstr = SQLstr & "Product_Details_ID, "
End If

If Not IsNull(TotalDebitArg) Then
    SQLstr = SQLstr & "Total_Debit, "
End If

If Not IsNull(TotalCreditArg) Then
    SQLstr = SQLstr & "Total_Credit, "
End If

If Not IsNull(DifferenceArg) Then
    SQLstr = SQLstr & "Difference, "
End If

If Not IsNull(NotesArg) Then
    SQLstr = SQLstr & "Notes) Values ("""
Else
    SQLstr = SQLstr & ") Values ("""
End If

If Not IsNull(TransactionIDArg) Then
    SQLstr = SQLstr & TransactionIDArg & """, """
End If

If Not IsNull(DocumentIDArg) Then
    SQLstr = SQLstr & DocumentIDArg & """, """
End If

If Not IsNull(FinancialDetailsIDArg) Then
    SQLstr = SQLstr & FinancialDetailsIDArg & """, """
End If

If Not IsNull(ProductDetailsIDArg) Then
    SQLstr = SQLstr & ProductDetailsIDArg & """, """
End If

If Not IsNull(TotalDebitArg) Then
    SQLstr = SQLstr & TotalDebitArg & """, """
End If

If Not IsNull(TotalCreditArg) Then
    SQLstr = SQLstr & TotalCreditArg & """, """
End If

If Not IsNull(DifferenceArg) Then
    SQLstr = SQLstr & DifferenceArg & """, """
End If

If Not IsNull(NotesArg) Then
    SQLstr = SQLstr & NotesArg & """)"
Else
    SQLstr = SQLstr & """)"
End If


SQLStringConstructionForInsertIntoFinancialBalanceErrorsTable = SQLstr

ExitProcedure:
Exit Function

ErrorHandler:
Select Case Err.Number
       
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: SQLStringConstructionForInsertIntoFinancialBalanceErrorsTable " & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
End Function
Function SQLStringConstructionForInsertIntoInventoryBalanceErrorsTable(Optional DifferenceArg As Double, Optional TotalDebitArg As Double, Optional TotalCreditArg As Double, Optional TransactionIDArg As Long, Optional DocumentIDArg As Long, Optional ProductDetailsIDArg As Long, Optional NotesArg As String) As String
Debug.Print "Module Public Functions - " & "SQLStringConstructionForInsertIntoInventoryBalanceErrorsTable " & Time()
On Error GoTo ErrorHandler

Dim SQLstr As String

SQLstr = "Insert into InventoryBalanceErrorsT ("

If Not IsNull(TransactionIDArg) Then
    SQLstr = SQLstr & "Transaction_ID, "
End If

If Not IsNull(DocumentIDArg) Then
    SQLstr = SQLstr & "Document_ID, "
End If

If Not IsNull(ProductDetailsIDArg) Then
    SQLstr = SQLstr & "Product_Details_ID, "
End If

If Not IsNull(TotalDebitArg) Then
    SQLstr = SQLstr & "Total_Debit, "
End If

If Not IsNull(TotalCreditArg) Then
    SQLstr = SQLstr & "Total_Credit, "
End If

If Not IsNull(DifferenceArg) Then
    SQLstr = SQLstr & "Difference, "
End If

If Not IsNull(NotesArg) Then
    SQLstr = SQLstr & "Notes) Values ("""
Else
    SQLstr = SQLstr & ") Values ("""
End If

If Not IsNull(TransactionIDArg) Then
    SQLstr = SQLstr & TransactionIDArg & """, """
End If

If Not IsNull(DocumentIDArg) Then
    SQLstr = SQLstr & DocumentIDArg & """, """
End If

If Not IsNull(ProductDetailsIDArg) Then
    SQLstr = SQLstr & ProductDetailsIDArg & """, """
End If

If Not IsNull(TotalDebitArg) Then
    SQLstr = SQLstr & TotalDebitArg & """, """
End If

If Not IsNull(TotalCreditArg) Then
    SQLstr = SQLstr & TotalCreditArg & """, """
End If

If Not IsNull(DifferenceArg) Then
    SQLstr = SQLstr & DifferenceArg & """, """
End If

If Not IsNull(NotesArg) Then
    SQLstr = SQLstr & NotesArg & """)"
Else
    SQLstr = SQLstr & """)"
End If

SQLStringConstructionForInsertIntoInventoryBalanceErrorsTable = SQLstr

ExitProcedure:
Exit Function

ErrorHandler:
Select Case Err.Number
       
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: SQLStringConstructionForInsertIntoInventoryBalanceErrorsTable " & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
End Function

Public Sub BackupFullTransaction(Optional TransactionIDArg As Long, Optional IssuedDocumentIDArg As Long)
Debug.Print "Module Public Functions - " & "BackupFullTransaction " & Time()
'On Error GoTo Errorhandler

Dim db As DAO.Database
Dim rstDocuments As DAO.Recordset
Dim VarTransactionID As Long

If TransactionIDArg = 0 And IssuedDocumentIDArg > 0 Then
  VarTransactionID = DLookup("IssuedDocumentT.Transaction_ID", "IssuedDocumentT", "Issued_Document_ID = " & IssuedDocumentIDArg)
  Else
  VarTransactionID = TransactionIDArg
End If

If VarTransactionID > 0 Then
  Set db = CurrentDb
  Set rstDocuments = db.OpenRecordset("select * from issuedDocumentT where Transaction_ID = " & VarTransactionID)

  If Not rstDocuments.EOF Then
    rstDocuments.MoveLast
    rstDocuments.MoveFirst
    Do Until rstDocuments.EOF
   
    Call CopyIssuedDocumentToBackupTable(rstDocuments(1), rstDocuments(0))
   
    rstDocuments.MoveNext
    Loop
  End If
Else
GoTo ExitProcedure
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
            "Error Source: BackupFullTransaction " & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
            
End Sub

Public Sub CopyLinkAttributeValueToEntitiesToBackupTable(EntityTypeIDArg As Integer, EntityIDArg As Long)
Debug.Print "Module Public Functions - " & "CopyLinkAttributeValueToEntitiesToBackupTable " & Time()

Dim db As DAO.Database
Dim rstLinkAttributeValueToEntitiesToInsertTobackupTable As DAO.Recordset
Dim rstForLastRecordInsertedToLinkAttributeValueToEntitiesBackupTable As DAO.Recordset
Dim VarlastLinkAttributeValueToEntityID_Backup_ID As Long
Dim AppendToLinkAttributeValueToEntitiesBackUpTableQ As QueryDef
Dim VarCurrentUser As Integer
Dim VarCurrentTimestamp As Date

VarCurrentUser = DLookup("[CurrentUserT]![Current_User_ID]", "[CurrentUserT]", "[Current_User_ID]>0")
VarCurrentTimestamp = Now()
'We insert into back up table ALL the records (LinkAttributeValueToEntityIDs)
'So, first we make a recordset with all the records that we are about to insert to the backup table
 Set db = CurrentDb
 Set rstLinkAttributeValueToEntitiesToInsertTobackupTable = db.OpenRecordset("Select * from LinkAttributeValueToEntitiesT WHERE Entity_Type_ID = " & EntityTypeIDArg & "  AND Entity_ID = " & EntityIDArg & " AND Is_New = false AND LinkAttributeValueToEntityID_Backup_ID is null")
If Not rstLinkAttributeValueToEntitiesToInsertTobackupTable.EOF Then
      rstLinkAttributeValueToEntitiesToInsertTobackupTable.MoveLast
      rstLinkAttributeValueToEntitiesToInsertTobackupTable.MoveFirst
'Then we itterate the recordset and insert its records one by one to the backup table
   Do Until rstLinkAttributeValueToEntitiesToInsertTobackupTable.EOF
      Set AppendToLinkAttributeValueToEntitiesBackUpTableQ = CurrentDb.QueryDefs("AttributeValuesToEntitiesSaveBeforeEditQ")
      AppendToLinkAttributeValueToEntitiesBackUpTableQ.Parameters(0) = VarCurrentTimestamp
      AppendToLinkAttributeValueToEntitiesBackUpTableQ.Parameters(1) = VarCurrentUser
      AppendToLinkAttributeValueToEntitiesBackUpTableQ.Parameters(2) = rstLinkAttributeValueToEntitiesToInsertTobackupTable(0)
      AppendToLinkAttributeValueToEntitiesBackUpTableQ.Execute dbFailOnError
      
 'After each one insertion to the backup table, we take the ID of this insert to the back up table and we update the rstLinkAttributeValueToEntitiesToInsertTobackupTable recordset (the basic table IssuedDocumentFinancialDetailsT)
    Set rstForLastRecordInsertedToLinkAttributeValueToEntitiesBackupTable = db.OpenRecordset("SELECT @@IDENTITY")
    VarlastLinkAttributeValueToEntityID_Backup_ID = rstForLastRecordInsertedToLinkAttributeValueToEntitiesBackupTable(0)
    rstLinkAttributeValueToEntitiesToInsertTobackupTable.Edit
    rstLinkAttributeValueToEntitiesToInsertTobackupTable!LinkAttributeValueToEntityID_Backup_ID = VarlastLinkAttributeValueToEntityID_Backup_ID
    rstLinkAttributeValueToEntitiesToInsertTobackupTable.Update
       CopyLinkSubAttributeValueToEntitiesToBackupTable (rstLinkAttributeValueToEntitiesToInsertTobackupTable(0))
    rstLinkAttributeValueToEntitiesToInsertTobackupTable.MoveNext
  Loop
    
End If
  
ExitProcedure:

If Not rstLinkAttributeValueToEntitiesToInsertTobackupTable Is Nothing Then
rstLinkAttributeValueToEntitiesToInsertTobackupTable.Close
Set rstLinkAttributeValueToEntitiesToInsertTobackupTable = Nothing
End If

If Not rstForLastRecordInsertedToLinkAttributeValueToEntitiesBackupTable Is Nothing Then
rstForLastRecordInsertedToLinkAttributeValueToEntitiesBackupTable.Close
Set rstForLastRecordInsertedToLinkAttributeValueToEntitiesBackupTable = Nothing
End If

If Not AppendToLinkAttributeValueToEntitiesBackUpTableQ Is Nothing Then
AppendToLinkAttributeValueToEntitiesBackUpTableQ.Close
Set AppendToLinkAttributeValueToEntitiesBackUpTableQ = Nothing
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
            "Error Source: CopyLinkAttributeValueToEntitiesToBackupTable " & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
               
End Sub

Public Sub CopyLinkSubAttributeValueToEntitiesToBackupTable(EntityIDArg As Long) ' we use the LinkAttributeValueToEntitiesBackupT, we just do it with different procedure (we do not use the already existing CopyLinkAttributeValueToEntitiesToBackupTable) because they are subattributes (but they are located also in the LinkAttributeValueToEntitiesT
Debug.Print "Module Public Functions - " & "CopyLinkSubAttributeValueToEntitiesToBackupTable " & Time()

Dim db As DAO.Database
Dim rstLinkSubAttributeValueToEntitiesToInsertTobackupTable As DAO.Recordset
Dim rstForLastRecordInsertedToLinkSubAttributeValueToEntitiesBackupTable As DAO.Recordset
Dim VarlastLinkSubAttributeValueToEntityID_Backup_ID As Long
Dim AppendToLinkSubAttributeValueToEntitiesBackUpTableQ As QueryDef
Dim VarCurrentUser As Integer
Dim VarCurrentTimestamp As Date

VarCurrentUser = DLookup("[CurrentUserT]![Current_User_ID]", "[CurrentUserT]", "[Current_User_ID]>0")
VarCurrentTimestamp = Now()

'We insert into back up table ALL the records (LinkAttributeValueToEntityIDs)
'So, first we make a recordset with all the records that we are about to insert to the backup table
 Set db = CurrentDb
 Set rstLinkSubAttributeValueToEntitiesToInsertTobackupTable = db.OpenRecordset("Select * from LinkAttributeValueToEntitiesT WHERE Entity_Type_ID = 7 AND Entity_ID = " & EntityIDArg & " AND Is_New = false AND LinkAttributeValueToEntityID_Backup_ID is null")
If Not rstLinkSubAttributeValueToEntitiesToInsertTobackupTable.EOF Then
      rstLinkSubAttributeValueToEntitiesToInsertTobackupTable.MoveLast
      rstLinkSubAttributeValueToEntitiesToInsertTobackupTable.MoveFirst
'Then we itterate the recordset and insert its records one by one to the backup table
   Do Until rstLinkSubAttributeValueToEntitiesToInsertTobackupTable.EOF
      Set AppendToLinkSubAttributeValueToEntitiesBackUpTableQ = CurrentDb.QueryDefs("AttributeValuesToEntitiesSaveBeforeEditQ")
      AppendToLinkSubAttributeValueToEntitiesBackUpTableQ.Parameters(0) = VarCurrentTimestamp
      AppendToLinkSubAttributeValueToEntitiesBackUpTableQ.Parameters(1) = VarCurrentUser
      AppendToLinkSubAttributeValueToEntitiesBackUpTableQ.Parameters(2) = rstLinkSubAttributeValueToEntitiesToInsertTobackupTable(0)
      AppendToLinkSubAttributeValueToEntitiesBackUpTableQ.Execute dbFailOnError
      
   'After each one insertion to the backup table, we take the ID of this insert to the back up table and we update the rstLinkSubAttributeValueToEntitiesToInsertTobackupTable recordset (the basic table IssuedDocumentFinancialDetailsT)

    Set rstForLastRecordInsertedToLinkSubAttributeValueToEntitiesBackupTable = db.OpenRecordset("SELECT @@IDENTITY")
    VarlastLinkSubAttributeValueToEntityID_Backup_ID = rstForLastRecordInsertedToLinkSubAttributeValueToEntitiesBackupTable(0)
    rstLinkSubAttributeValueToEntitiesToInsertTobackupTable.Edit
    rstLinkSubAttributeValueToEntitiesToInsertTobackupTable!LinkAttributeValueToEntityID_Backup_ID = VarlastLinkSubAttributeValueToEntityID_Backup_ID
    rstLinkSubAttributeValueToEntitiesToInsertTobackupTable.Update
    rstLinkSubAttributeValueToEntitiesToInsertTobackupTable.MoveNext
  Loop
    
End If
  
ExitProcedure:

If Not rstLinkSubAttributeValueToEntitiesToInsertTobackupTable Is Nothing Then
rstLinkSubAttributeValueToEntitiesToInsertTobackupTable.Close
Set rstLinkSubAttributeValueToEntitiesToInsertTobackupTable = Nothing
End If

If Not rstForLastRecordInsertedToLinkSubAttributeValueToEntitiesBackupTable Is Nothing Then
rstForLastRecordInsertedToLinkSubAttributeValueToEntitiesBackupTable.Close
Set rstForLastRecordInsertedToLinkSubAttributeValueToEntitiesBackupTable = Nothing
End If

If Not AppendToLinkSubAttributeValueToEntitiesBackUpTableQ Is Nothing Then
AppendToLinkSubAttributeValueToEntitiesBackUpTableQ.Close
Set AppendToLinkSubAttributeValueToEntitiesBackUpTableQ = Nothing
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
            "Error Source: CopyLinkSubAttributeValueToEntitiesToBackupTable " & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
               
End Sub

Public Sub EmptyIssuedDocumentsTBackUpIDFieldByTransactionID(TransactionIDArg As Long) 'it also trigers empty backupID fields for IssuedDocumentFinancialDetails, IssuedDocumentProductDetails and LinkAttributeValuesToEntities
Debug.Print "Module Public Functions - " & "EmptyIssuedDocumentsTBackUpIDFieldByTransactionID " & Time()
'On Error GoTo ErrorHandler

Dim db As DAO.Database
Dim rstIssuedDocuments As DAO.Recordset

Set db = CurrentDb
Set rstIssuedDocuments = db.OpenRecordset("Select * from IssuedDocumentT where Transaction_ID = " & TransactionIDArg)

If Not rstIssuedDocuments.EOF Then
  rstIssuedDocuments.MoveLast
  rstIssuedDocuments.MoveFirst

  Do Until rstIssuedDocuments.EOF
   rstIssuedDocuments.Edit
   rstIssuedDocuments("Issued_Document_Backup_ID") = Null
   Debug.Print rstIssuedDocuments("Issued_Document_ID")
   rstIssuedDocuments.Update
   Call EmptyIssuedDocumentFinancialDetailsTBackUpIDFieldByDocumentID(rstIssuedDocuments("Issued_Document_ID"))
   Call EmptyIssuedDocumentProductDetailsTBackUpIDFieldByDocumentID(rstIssuedDocuments("Issued_Document_ID"))
   Call EmptyLinkAttributeValuesToEntitiesTBackUpIDField(rstIssuedDocuments("Issued_Document_ID"), 3)
   Call EmptyDiscountLogsTBackUpIDField(rstIssuedDocuments("Issued_Document_ID"))
   rstIssuedDocuments.MoveNext
  Loop
End If
  
ExitProcedure:
Exit Sub
   
ErrorHandler:
Select Case Err.Number
        
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: EmptyIssuedDocumentsTBackUpIDFieldByTransactionID" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
End Sub

Public Sub EmptyIssuedDocumentFinancialDetailsTBackUpIDFieldByDocumentID(IssuedDocumentIDArg As Long) 'it also trigers empty backupID fields for LinkAttributeValuesToEntities
Debug.Print "Module Public Functions - " & "EmptyIssuedDocumentFinancialDetailsTBackUpIDFieldByDocumentID " & Time()
'On Error GoTo ErrorHandler

Dim db As DAO.Database
Dim rstIssuedDocumentFinancialDetails As DAO.Recordset

Set db = CurrentDb
Set rstIssuedDocumentFinancialDetails = db.OpenRecordset("Select * from IssuedDocumentFinancialDetailsT where Issued_Document_ID = " & IssuedDocumentIDArg)
Debug.Print rstIssuedDocumentFinancialDetails.recordcount

If Not rstIssuedDocumentFinancialDetails.EOF Then
  rstIssuedDocumentFinancialDetails.MoveLast
  rstIssuedDocumentFinancialDetails.MoveFirst

  Do Until rstIssuedDocumentFinancialDetails.EOF
   rstIssuedDocumentFinancialDetails.Edit
   rstIssuedDocumentFinancialDetails("IssuedDocumentFinancialDetails_Backup_ID") = Null
   rstIssuedDocumentFinancialDetails.Update
   Call EmptyLinkAttributeValuesToEntitiesTBackUpIDField(rstIssuedDocumentFinancialDetails("Issued_Document_Financial_Details_ID"), 4)
   rstIssuedDocumentFinancialDetails.MoveNext
  Loop
End If
  
ExitProcedure:
Exit Sub
   
ErrorHandler:
Select Case Err.Number
        
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: EmptyIssuedDocumentFinancialDetailsTBackUpIDFieldByDocumentID" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
End Sub

Public Sub EmptyIssuedDocumentProductDetailsTBackUpIDFieldByDocumentID(IssuedDocumentIDArg As Long) 'it also trigers empty backupID fields for  LinkAttributeValuesToEntities
'Debug.Print "Module Public Functions - " & "EmptyIssuedDocumentProductDetailsTBackUpIDFieldByDocumentID " & Time()
On Error GoTo ErrorHandler

Dim db As DAO.Database
Dim rstIssuedDocumentProductDetails As DAO.Recordset

Set db = CurrentDb
Set rstIssuedDocumentProductDetails = db.OpenRecordset("Select * from IssuedDocumentProductDetailsT where Issued_Document_ID = " & IssuedDocumentIDArg)

If Not rstIssuedDocumentProductDetails.EOF Then
  rstIssuedDocumentProductDetails.MoveLast
  rstIssuedDocumentProductDetails.MoveFirst

  Do Until rstIssuedDocumentProductDetails.EOF
   rstIssuedDocumentProductDetails.Edit
   rstIssuedDocumentProductDetails("IssuedDocumentProductDetails_Backup_ID") = Null
   rstIssuedDocumentProductDetails.Update
   Call EmptyLinkAttributeValuesToEntitiesTBackUpIDField(rstIssuedDocumentProductDetails("Issued_Document_Product_Details_ID"), 5)
   rstIssuedDocumentProductDetails.MoveNext
  Loop
End If
  
ExitProcedure:
Exit Sub
   
ErrorHandler:
Select Case Err.Number
        
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: EmptyIssuedDocumentProductDetailsTBackUpIDFieldByDocumentID" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
End Sub

Public Sub EmptyLinkAttributeValuesToEntitiesTBackUpIDField(EntityIDArg As Long, EntityTypeIDArg As Integer) 'it also trigers empty backupID fields for LinkAttributeValuesToEntities (for subattributes only)
Debug.Print "Module Public Functions - " & "EmptyLinkAttributeValuesToEntitiesTBackUpIDField " & Time()
'On Error GoTo ErrorHandler

Dim db As DAO.Database
Dim rstLinkAttributeValuesToEntities As DAO.Recordset

Set db = CurrentDb
Set rstLinkAttributeValuesToEntities = db.OpenRecordset("Select * from LinkAttributeValueToEntitiesT where Entity_ID = " & EntityIDArg & " AND Entity_Type_ID = " & EntityTypeIDArg)

If Not rstLinkAttributeValuesToEntities.EOF Then
  rstLinkAttributeValuesToEntities.MoveLast
  rstLinkAttributeValuesToEntities.MoveFirst

  Do Until rstLinkAttributeValuesToEntities.EOF
   rstLinkAttributeValuesToEntities.Edit
   rstLinkAttributeValuesToEntities("LinkAttributeValueToEntityID_Backup_ID") = Null
   rstLinkAttributeValuesToEntities.Update
   Call EmptyLinkAttributeValuesToEntitiesTBackUpIDFieldForSubAttributes(rstLinkAttributeValuesToEntities("Link_Attribute_Value_To_Entity_ID"))
   rstLinkAttributeValuesToEntities.MoveNext
  Loop
End If
  
ExitProcedure:
Exit Sub
   
ErrorHandler:
Select Case Err.Number
        
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: EmptyLinkAttributeValuesToEntitiesTBackUpIDField" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
End Sub

Public Sub EmptyLinkAttributeValuesToEntitiesTBackUpIDFieldForSubAttributes(LinkAttributeValueToEntityIDArg As Long)
Debug.Print "Module Public Functions - " & "EmptyLinkAttributeValuesToEntitiesTBackUpIDFieldForSubAttributes " & Time()
'On Error GoTo ErrorHandler

Dim db As DAO.Database
Dim rstLinkAttributeValuesToSubAttributes As DAO.Recordset

Set db = CurrentDb
Set rstLinkAttributeValuesToSubAttributes = db.OpenRecordset("Select * from LinkAttributeValueToEntitiesT where Entity_ID = " & LinkAttributeValueToEntityIDArg & " AND Entity_Type_ID = 7")

If Not rstLinkAttributeValuesToSubAttributes.EOF Then
  rstLinkAttributeValuesToSubAttributes.MoveLast
  rstLinkAttributeValuesToSubAttributes.MoveFirst

  Do Until rstLinkAttributeValuesToSubAttributes.EOF
   rstLinkAttributeValuesToSubAttributes.Edit
   rstLinkAttributeValuesToSubAttributes("LinkAttributeValueToEntityID_Backup_ID") = Null
   rstLinkAttributeValuesToSubAttributes.Update
   rstLinkAttributeValuesToSubAttributes.MoveNext
  Loop
End If
  
ExitProcedure:
Exit Sub
   
ErrorHandler:
Select Case Err.Number
        
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: EmptyLinkAttributeValuesToEntitiesTBackUpIDFieldForSubAttributes" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
End Sub

Public Sub EmptyDiscountLogsTBackUpIDField(IssuedDocumentIDArg As Long) 'it also trigers empty backupID fields for DiscountLogsDetailsT
Debug.Print "Module Public Functions - " & "EmptyDiscountLogsTBackUpIDField " & Time()
'On Error GoTo ErrorHandler

Dim db As DAO.Database
Dim rstDiscountLogs As DAO.Recordset

Set db = CurrentDb
Set rstDiscountLogs = db.OpenRecordset("Select * from DiscountLogsT where Issued_Document_ID = " & IssuedDocumentIDArg)

If Not rstDiscountLogs.EOF Then
  rstDiscountLogs.MoveLast
  rstDiscountLogs.MoveFirst

  Do Until rstDiscountLogs.EOF
   rstDiscountLogs.Edit
   rstDiscountLogs("DiscountLogs_Backup_ID") = Null
   rstDiscountLogs.Update
   Call EmptyDiscountLogsDetailsTBackUpIDField(rstDiscountLogs("Discount_Logs_ID"))
   rstDiscountLogs.MoveNext
  Loop
End If
  
ExitProcedure:
Exit Sub
   
ErrorHandler:
Select Case Err.Number
        
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: EmptyDiscountLogsTBackUpIDField" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
End Sub

Public Sub EmptyDiscountLogsDetailsTBackUpIDField(DiscountLogsIDArg As Long)
Debug.Print "Module Public Functions - " & "EmptyDiscountLogsTBackUpIDField " & Time()
On Error GoTo ErrorHandler

Dim db As DAO.Database
Dim rstDiscountLogDetails As DAO.Recordset

Set db = CurrentDb
Set rstDiscountLogDetails = db.OpenRecordset("Select * from DiscountLogsDetailsT where Discount_Logs_ID = " & DiscountLogsIDArg)

If Not rstDiscountLogDetails.EOF Then
  rstDiscountLogDetails.MoveLast
  rstDiscountLogDetails.MoveFirst

  Do Until rstDiscountLogDetails.EOF
   rstDiscountLogDetails.Edit
   rstDiscountLogDetails("DiscountLogsDetails_Backup_ID") = Null
   rstDiscountLogDetails.Update
   rstDiscountLogDetails.MoveNext
  Loop
End If
  
ExitProcedure:
Exit Sub
   
ErrorHandler:
Select Case Err.Number
        
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: EmptyDiscountLogsTBackUpIDField" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
End Sub

Public Sub UpdateAllFinancialTransactorsBalanceFromRecords()
Debug.Print "Module Public Functions - " & "UpdateAllFinancialTransactorsBalanceFromRecords " & Time()
'On Error GoTo ErrorHandler

If TableExists("RecalculatedFinancialBalancePerFinTransactorFromRecordsTempT") Then
CurrentDb.Execute ("DROP TABLE RecalculatedFinancialBalancePerFinTransactorFromRecordsTempT")
End If

CurrentDb.Execute "ZeroingTransactorsTotalDebitAndTotalCreditQ", dbFailOnError
CurrentDb.Execute "RecalculatedFinTranBalanceStoredToTempTQ", dbFailOnError
CurrentDb.Execute "UpdateFinancialTransactorsBalanceFromRecordsQ", dbFailOnError
CurrentDb.Execute ("DROP TABLE RecalculatedFinancialBalancePerFinTransactorFromRecordsTempT")

ExitProcedure:
If TableExists(" RecalculatedFinancialBalancePerFinTransactorFromRecordsTempT") Then
CurrentDb.Execute ("DROP TABLE RecalculatedFinancialBalancePerFinTransactorFromRecordsTempT")
End If

Exit Sub
   
   
ErrorHandler:
Select Case Err.Number
        
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: UpdateAllFinancialTransactorsBalanceFromRecords" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select


End Sub

Public Sub UpdateAllInventoryBalanceFromRecords()
Debug.Print "Module Public Functions - " & "UpdateAllInventoryBalanceFromRecords " & Time()
On Error GoTo ErrorHandler

CurrentDb.Execute "EmptyingProductInventoryBalanceTableQ", dbFailOnError
CurrentDb.Execute "AppendToInventoryBalanceAllRecalculatedRecordsFromRecordsQ", dbFailOnError

ExitProcedure:

Exit Sub
   
   
ErrorHandler:
Select Case Err.Number
        
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: UpdateAllInventoryBalanceFromRecords" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select


End Sub
Function TableExists(TableName As String) As Boolean
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim exists As Boolean
    
    exists = False
    Set db = CurrentDb()
    
    For Each tdf In db.TableDefs
        If tdf.Name = TableName Then
            exists = True
            Exit For
        End If
    Next tdf
    
    TableExists = exists
End Function

Public Sub ErrorLoging(ErrorNumberArg As Long, ErrorDescriptionArg As String, ModuleErrorAppearedInArg As String, ProcedureOrFunctionArg As String)
'On Error GoTo ErrorHandler
Dim VarProcedureTitle As String
Dim VarModuleName As String
VarProcedureTitle = "ErrorLoging"
VarModuleName = "Module Public Functions"
'Debug.Print VarModuleName & " - " & VarProcedureTitle & " - " & Time()


    Dim db As DAO.Database
    Set db = CurrentDb
 
    db.Execute "INSERT INTO ErrorLogingT (Error_Number, Error_Description, Module_Error_Appeared_In, Procedure_Or_Function, Error_Time, user_ID) " & _
               "VALUES (" & ErrorNumberArg & " , '" & Replace(ErrorDescriptionArg, "'", "''") & "', '" & Replace(ModuleErrorAppearedInArg, "'", "''") & "', '" & Replace(ProcedureOrFunctionArg, "'", "''") & "', #" & Format(Now(), "mm/dd/yyyy hh:nn:ss AM/PM") & "#, " & FetchUserID() & ");", dbFailOnErrorExitProcedure:
If Not db Is Nothing Then
db.Close
Set db = Nothing
End If

ExitProcedure:
Exit Sub
   
ErrorHandler:
MsgBox "Error while error loging with ErrorLoging procedure! Possible failure to insert also other errors!"
    Select Case Err.Number
         Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: " & VarModuleName & " - " & VarProcedureTitle & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
    End Select
End Sub

Public Sub ShowErrorMessage(ErrorNumArg As Variant, ErrorDescriptionArg As String, ModuleNameArg As String, ProcedureTitleArg As String)
On Error GoTo ErrorHandler
Debug.Print "Module Public Functions - " & "ShowErrorMessage " & Time()

MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & ErrorNumArg & vbCrLf & _
            "Error Source: " & ModuleNameArg & " - " & ProcedureTitleArg & vbCrLf & _
            "Error Description: " & ErrorDescriptionArg _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            
ExitProcedure:
Exit Sub
   
ErrorHandler:
Call ErrorLoging(Err.Number, Err.Description, "Module Public Functions", "ShowErrorMessage")
Resume Next
GoTo ExitProcedure
End Sub

Public Sub CloseDatabase(AskConfirmationArg As Boolean)
'On Error GoTo ErrorHandler
Dim VarProcedureTitle As String
Dim VarModuleName As String
VarProcedureTitle = "CloseDatabase"
VarModuleName = "Module Public Functions"
Debug.Print VarModuleName & " - " & VarProcedureTitle & " - " & Time()
        '---------------------------------------------------------------------
 Dim iResponse As Integer
 Dim obj As Object
 Dim VarLoginActivityID As Long
 
  If AskConfirmationArg = True Then
    ' Prompt user for confirmation
    iResponse = MsgBox("Do you want to log out?", vbYesNo + vbQuestion, "Logout Confirmation")

    ' If user confirms logout
    If iResponse = vbNo Then
      GoTo ExitProcedure
    End If
  End If
  
        ' Close all open forms
        For Each obj In Application.CurrentProject.AllForms
            If CheckIfObjectIsLoaded(obj.Name) Then
                DoCmd.Close acForm, obj.Name, acSaveNo
            End If
        Next obj
        
         ' Close all open reports
        For Each obj In Application.CurrentProject.AllReports
            If CheckIfObjectIsLoaded(obj.Name) Then
                DoCmd.Close acReport, obj.Name, acSaveNo
            End If
        Next obj
        
       
       
         ' Close all open Modules
        For Each obj In Application.CurrentProject.AllModules
            If CheckIfObjectIsLoaded(obj.Name) Then
                DoCmd.Close acReport, obj.Name, acSaveNo
            End If
        Next obj
        
        
        ' Storing Logout timestamp
         VarLoginActivityID = DLookup("[CurrentUserT]![Login_Activity_ID]", "[CurrentUserT]")
         CurrentDb.Execute "Update LoginActivityT SET Logout_Timestamp = #" & Format(Now(), "mm/dd/yyyy hh:nn:ss AM/PM") & "# Where Log_In_ID = " & VarLoginActivityID
        
        ' Quit the application
        Application.Quit acQuitSaveNone

    
ExitProcedure:
Exit Sub
   
ErrorHandler:
Dim VarErrorNum As Long
Dim VarErrorDescription As String
VarErrorNum = Err.Number
VarErrorDescription = Err.Description
Call ErrorLoging(Err.Number, Err.Description, VarModuleName, VarProcedureTitle)

    Select Case VarErrorNum
         Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & VarErrorNum & vbCrLf & _
            "Error Source: " & VarModuleName & " - " & VarProcedureTitle & vbCrLf & _
            "Error Description: " & VarErrorDescription _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
    End Select
End Sub

Private Function CheckIfObjectIsLoaded(ByVal strObjectName As String) As Boolean
On Error GoTo ErrorHandler
Dim VarProcedureTitle As String
Dim VarModuleName As String
VarProcedureTitle = "CheckIfObjectIsLoaded"
VarModuleName = "Module Public Functions"
Debug.Print VarModuleName & " - " & VarProcedureTitle & " - " & Time()
        '---------------------------------------------------------------------
    
    CheckIfObjectIsLoaded = (SysCmd(acSysCmdGetObjectState, acForm, strObjectName) And acObjStateOpen) <> 0

ExitProcedure:
Exit Function
   
ErrorHandler:
Dim VarErrorNum As Long
Dim VarErrorDescription As String
VarErrorNum = Err.Number
VarErrorDescription = Err.Description
Call ErrorLoging(Err.Number, Err.Description, VarModuleName, VarProcedureTitle)

    Select Case VarErrorNum
         Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & VarErrorNum & vbCrLf & _
            "Error Source: " & VarModuleName & " - " & VarProcedureTitle & vbCrLf & _
            "Error Description: " & VarErrorDescription _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
    End Select
End Function

Public Sub RestoreRibbonAndNavigationPane()
On Error GoTo ErrorHandler
Dim VarProcedureTitle As String
Dim VarModuleName As String
VarProcedureTitle = "RestoreRibbonAndNavigationPane"
VarModuleName = "Module Public Functions"
Debug.Print VarModuleName & " - " & VarProcedureTitle & " - " & Time()
        '---------------------------------------------------------------------
 
DoCmd.ShowToolbar "Ribbon", acToolbarYes
DoCmd.SelectObject acTable, , True

ExitProcedure:
Exit Sub
   
ErrorHandler:
Dim VarErrorNum As Long
Dim VarErrorDescription As String
VarErrorNum = Err.Number
VarErrorDescription = Err.Description
Call ErrorLoging(Err.Number, Err.Description, VarModuleName, VarProcedureTitle)

    Select Case VarErrorNum
         Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & VarErrorNum & vbCrLf & _
            "Error Source: " & VarModuleName & " - " & VarProcedureTitle & vbCrLf & _
            "Error Description: " & VarErrorDescription _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
    End Select
End Sub


Public Sub HideRibbonAndNavigationPane()
On Error GoTo ErrorHandler
Dim VarProcedureTitle As String
Dim VarModuleName As String
VarProcedureTitle = "HideRibbonAndNavigationPane"
VarModuleName = "Module Public Functions"
Debug.Print VarModuleName & " - " & VarProcedureTitle & " - " & Time()
        '---------------------------------------------------------------------
        
If CommandBars("Ribbon").Visible = True Then

DoCmd.ShowToolbar "Ribbon", acToolbarNo
DoCmd.NavigateTo "acNavigationCategoryObjectType"
DoCmd.RunCommand acCmdWindowHide

End If

ExitProcedure:
Exit Sub
   
ErrorHandler:
Dim VarErrorNum As Long
Dim VarErrorDescription As String
VarErrorNum = Err.Number
VarErrorDescription = Err.Description
Call ErrorLoging(Err.Number, Err.Description, VarModuleName, VarProcedureTitle)

    Select Case VarErrorNum
         Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & VarErrorNum & vbCrLf & _
            "Error Source: " & VarModuleName & " - " & VarProcedureTitle & vbCrLf & _
            "Error Description: " & VarErrorDescription _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
    End Select
End Sub