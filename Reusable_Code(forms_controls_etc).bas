'Purpose:   This module illustrates how to create a search form, _
            where the user can enter as many or few criteria as they wish, _
            and results are shown one per line.
'Note:      Only records matching ALL of the criteria are returned.
'Author:    Allen Browne (allen@allenbrowne.com), June 2006.
Option Compare Database
Dim FormName As Form
'Option Explicit



Private Sub cmdFilter_Click()
    'Purpose:   Build up the criteria string form the non-blank search boxes, and apply to the form's Filter.
    'Notes:     1. We tack " AND " on the end of each condition so you can easily add more search boxes; _
                        we remove the trailing " AND " at the end.
    '           2. The date range works like this: _
                        Both dates      = only dates between (both inclusive. _
                        Start date only = all dates from this one onwards; _
                        End date only   = all dates up to (and including this one).
    Dim strWhere As String                  'The criteria string.
    Dim lngLen As Long                      'Length of the criteria string to append to.
    Const conJetDate = "\#mm\/dd\/yyyy\#"   'The format expected for dates in a JET query string.
    
    '***********************************************************************
    'Look at each search box, and build up the criteria string from the non-blank ones.
    '***********************************************************************
    'Text field example. Use quotes around the value in the string.
    If Not IsNull(FormName.txtFilterCity) Then
        strWhere = strWhere & "([City] = """ & FormName.txtFilterCity & """) AND "
    End If
    
    'Another text field example. Use Like to find anywhere in the field.
    If Not IsNull(FormName.txtFilterMainNaFORMNAME) Then
        strWhere = strWhere & "([MainNaFORMNAME] Like ""*" & FormName.txtFilterMainNaFORMNAME & "*"") AND "
    End If
    
    'Number field example. Do not add the extra quotes.
    If Not IsNull(FormName.cboFilterLevel) Then
        strWhere = strWhere & "([LevelID] = " & FormName.cboFilterLevel & ") AND "
    End If
    
    'Yes/No field and combo example. If combo is blank or contains "ALL", we do nothing.
    If FormName.cboFilterIsCorporate = -1 Then
        strWhere = strWhere & "([IsCorporate] = True) AND "
    ElseIf FormName.cboFilterIsCorporate = 0 Then
        strWhere = strWhere & "([IsCorporate] = False) AND "
    End If
    
    'Date field example. Use the format string to add the # delimiters and get the right international format.
    If Not IsNull(FormName.txtStartDate) Then
        strWhere = strWhere & "([EnteredOn] >= " & Format(FormName.txtStartDate, conJetDate) & ") AND "
    End If
    
    'Another date field example. Use "less than the next day" since this field has tiFORMNAMEs as well as dates.
    If Not IsNull(FormName.txtEndDate) Then   'Less than the next day.
        strWhere = strWhere & "([EnteredOn] < " & Format(FormName.txtEndDate + 1, conJetDate) & ") AND "
    End If
    
    '***********************************************************************
    'Chop off the trailing " AND ", and use the string as the form's Filter.
    '***********************************************************************
    'See if the string has more than 5 characters (a trailng " AND ") to remove.
    lngLen = Len(strWhere) - 5
    If lngLen <= 0 Then     'Nah: there was nothing in the string.
        MsgBox "No criteria", vbInformation, "Nothing to do."
    Else                    'Yep: there is soFORMNAMEthing there, so remove the " AND " at the end.
        strWhere = Left$(strWhere, lngLen)
        'For debugging, remove the leading quote on the next line. Prints to ImFORMNAMEdiate Window (Ctrl+G).
        'Debug.Print strWhere
        
        'Finally, apply the string as the form's Filter.
        FormName.Filter = strWhere
        FormName.FilterOn = True
    End If
End Sub

Private Sub cmdReset_Click()
    'Purpose:   Clear all the search boxes in the Form Header, and show all records again.
    Dim ctl As Control
    
    'Clear all the controls in the Form Header section.
    For Each ctl In FormName.Section(acHeader).Controls
        Select Case ctl.ControlType
        Case acTextBox, acComboBox
            ctl.Value = Null
        Case acCheckBox
            ctl.Value = False
        End Select
    Next
    
    'Remove the form's filter.
    FormName.FilterOn = False
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    'To avoid problems if the filter returns no records, we did not set its AllowAdditions to No.
    'We prevent new records by cancelling the form's BeforeInsert event instead.
    'The problems are explained at http://allenbrowne.com/bug-06.html
    Cancel = True
    MsgBox "You cannot add new clients to the search form.", vbInformation, "Permission denied."
End Sub

Private Sub Form_Open(Cancel As Integer)
    'Remove the single quote from these lines if you want to initially show no records.
    'FORMNAME.Filter = "(False)"
    'FORMNAME.FilterOn = True
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'WE CAN USE THIS CODE TO ANY CONTINUOUS FORM'S "ON KEY DOWN EVENT" TO CALL THE PUBLIC FUNCTION THAT ENABLES ARROW KEY MOVEMENTS BETWEEN RECORDS
'KEY PREVIEW OF THE FORM MUST BE SET TO "YES"
    
   ' On Error GoTo Error_Handler
 
    'KeyCode = EnableArrowsScroll(KeyCode, Me)
     
'Error_Handler_Exit:
  '  On Error Resume Next
   ' Exit Sub
 
'Error_Handler:
   ' MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
        '   "Error Number: " & Err.Number & vbCrLf & _
        '   "Error Source: Form_KeyDown" & vbCrLf & _
        '   "Error Description: " & Err.Description & _
         '  Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
         '  , vbOKOnly + vbCritical, "An Error has Occurred!"
 '   Resume Error_Handler_Exit
End Sub

Public Sub IterateRecordsets(ByRef RecordsetArg As DAO.Recordset, Optional NotesArg As String)
On Error GoTo ErrorHandler

Dim dbForIterateRecords As DAO.Database
Dim rs As DAO.Recordset


Set dbForIterateRecords = CurrentDb
On Error GoTo 0
If NotesArg <> "" Then
Debug.Print NotesArg
End If

Debug.Print "Recordcount = " & RecordsetArg.recordcount
If RecordsetArg.recordcount = 0 Then
Debug.Print "No records found"
GoTo ExitProcedure
Else
RecordsetArg.MoveLast
RecordsetArg.MoveFirst
 Dim i As Long
    For i = 0 To RecordsetArg.Fields.Count - 1
        Debug.Print RecordsetArg.Fields(i).Name,
   Next
       RecordsetArg.MoveFirst
       Debug.Print
    Do Until RecordsetArg.EOF
         For i = 0 To RecordsetArg.Fields.Count - 1
         Debug.Print RecordsetArg.Fields(i).Value,
         Next
         Debug.Print
       RecordsetArg.MoveNext
    Loop
RecordsetArg.MoveFirst
End If



ExitProcedure:
If Not rs Is Nothing Then
    rs.Close
    Set rs = Nothing
End If

If Not dbForIterateRecords Is Nothing Then
    dbForIterateRecords.Close
    Set dbForIterateRecords = Nothing
End If
Exit Sub
   
ErrorHandler:
Select Case Err.Number
        Case 3167
        Response = acDataErrContinue
        Case 3314
        MsgBox "You have left empty fields which must be filled", vbInformation, "������ ���������"
        Response = acDataErrContinue
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: IterateRecordsets" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
    
End Sub
Sub iterateRecordsFromRecordset(SQLstrArg As String)
Dim db As DAO.Database
Dim Q As DAO.QueryDef
Dim rs As DAO.Recordset
Dim SQLstr As String


SQLstr = SQLstrArg

Set db = CurrentDb
Set Q = db.CreateQueryDef("", SQLstr)
Set rs = Q.OpenRecordset()

Debug.Print rs.recordcount

 Dim i As Long

    For i = 0 To rs.Fields.Count - 1
        Debug.Print rs.Fields(i).Name,
   Next
    rs.MoveFirst
    Do Until rs.EOF
        Debug.Print
        For i = 0 To rs.Fields.Count - 1
            Debug.Print rs.Fields(i).Value,
       Next
       rs.MoveNext
    Loop
    Set db = Nothing
    Set rs = Nothing
    Set Q = Nothing
End Sub


Sub ErrorHandling()
On Error GoTo ErrorHandler

ExitProcedure:
Exit Sub
   
ErrorHandler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: ��� ������� ��� sub � function" & vbCrLf & _
           "Error Description: " & Err.Description _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    Resume ExitProcedure
End Sub
Sub ErrorHandlingNew()
On Error GoTo ErrorHandler
Dim VarProcedureTitle As String
Dim VarModuleName As String
VarProcedureTitle = "��� ������� ��� ����� ��� PROCEDURE"
VarModuleName = "Me.Name" ' �� ����� �� "". �� ����� ����� ������� �� me.name ���� ����� test compile �� ����. �� ��������� ��� module ��� ��� ������ �� �����, ���� ���� ��� me.name ������� VarModuleName = "�� ����� ��� module"
Debug.Print VarModuleName & " - " & VarProcedureTitle & " - " & Time()
        '---------------------------------------------------------------------

        '---------------------------------------------------------------------
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


Sub Errorhandling_With_Select_Case()
On Error GoTo ErrorHandler

ExitProcedure:
Exit Sub
   
ErrorHandler:
Select Case Err.Number
        Case 3314
        MsgBox "You have left empty fields which must be filled", vbInformation, "������ ���������"
        Response = acDataErrContinue
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: ��� ������� ��� sub � function" & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select
End Sub


Sub OpenRecordsetOutput(rstOutput As Recordset)
     
       ' Enumerate the specified Recordset object.
       With rstOutput
          Do While Not .EOF
             Debug.Print , .Fields(0), .Fields(1)
             .MoveNext
          Loop
       End With
     
    End Sub
    
Private Sub Form_BeforeUpdate(Cancel As Integer) 'CheckForNullControlsBeforeFormUpdate

Dim ctl As Control

'For Each ctl In Me.Controls
     Select Case TypeName(ctl)
         Case "textBox", "combobox"
         With ctl
             If (IsNull(.Value) Or Len(.Value & vbNullString) = 0) And .Tag <> "x" Then
                 If MsgBox(.Name & " control not filled" & vbNewLine & _
                           "fill it..!", vbInformation) = vbOK Then
                           Cancel = True
                           ctl.SetFocus
                           Exit Sub
                 End If
             End If
         End With
         End Select
    ' Next
 Set ctl = Nothing

End Sub

Private Sub SourceTextBoxName_KeyPress(KeyAscii As Integer) ' Transfers focus and 1st keystroke from Textbox to Combobox as text, in order to continue search into Combobox field. It is placed in the SourceTextbox module
DestinationCbo.SetFocus
DestinationCbo.Text = ChrW(KeyAscii)
DestinationCbo.SelStart = DestinationCbo.SelLength + 1
End Sub

Private Sub LastInsertedID()
'it must follow right after the insert query in the same module and in the same procedure, otherwise it returns 0.
Dim db As DAO.Database
Dim rstForLastTransactionID As DAO.Recordset
Dim VarlastTransactionFinalID As Long
Set db = CurrentDb
Set rstForLastTransactionID = db.OpenRecordset("SELECT @@IDENTITY")
VarlastTransactionFinalID = rstForLastTransactionID(0)
rstForLastTransactionID.Close
db.Close
Set rstForLastTransactionID = Nothing
Set db = Nothing
End Sub

Private Sub CommonUseCodeForAllSubs()
  On Error GoTo ErrorHandler
  Dim VarProcedureTitle As String
  Dim VarModuleName As String
  VarProcedureTitle = "DeleteMonetaryCommercialDetailsRecord"
  VarModuleName = "Me.Name"
  Debug.Print VarModuleName & " - " & VarProcedureTitle & " - " & Time()
  '---------------------------------------------------------------------




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
        Call ShowErrorMessage(VarErrorNum, VarErrorDescription, VarModuleName, VarProcedureTitle)
    Resume ExitProcedure
  End Select
End Sub