Option Compare Database
Option Explicit
Public Sub UpdateTotalDebitAndTotalCreditForAllTransactors()
Debug.Print "Module Public Functions - " & "UpdateTotalDebitAndTotalCreditForAllTransactors " & Time()
'On Error GoTo Errorhandler



ExitProcedure:
Exit Sub
   
ErrorHandler:
Select Case Err.Number
       
        Case Else
            MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: UpdateTotalDebitAndTotalCreditForAllTransactors " & vbCrLf & _
            "Error Description: " & Err.Description _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
            Resume ExitProcedure
            End Select

End Sub

Option Compare Database
Option Explicit

Public Sub ExportTableSchemasToFile() ' this sub exports to a file all the tables and their fields in order to provide them to chatgrp
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    
    Dim strFile As String
    Dim intFileNum As Integer
    
    ' Set your desired file path here
    strFile = "C:\Users\User\Desktop\��������\TableSchemas.txt"
    
    ' Get an available file number
    intFileNum = FreeFile
    
    ' Open (create) the file for output
    Open strFile For Output As #intFileNum
    
    Set db = CurrentDb
    
    For Each tdf In db.TableDefs
        ' Skip system and temporary tables (often start with 'MSys' or '~tmp')
        If Not (tdf.Name Like "MSys*" Or tdf.Name Like "~tmp*") Then
            
            Print #intFileNum, "Table: " & tdf.Name
            
            For Each fld In tdf.Fields
                ' Print field name and the numeric DAO data type
                'Print #intFileNum, "    Field: " & fld.Name & " (Type: " & fld.Type & ")"
                Print #intFileNum, "    Field: " & fld.Name & " (Type: " & GetFieldTypeName(fld.Type) & ")"
            Next fld
            
            Print #intFileNum, '--- Blank line for readability
        End If
    Next tdf
    
    Close #intFileNum
    
    MsgBox "Table schemas exported to: " & strFile, vbInformation
End Sub



Public Sub ExportRelationshipInfoToFile() ' this sub exports to a file all the tables'  relationships in order to provide them to chatgrp
    Dim db As DAO.Database
    Dim rel As DAO.Relation
    Dim relField As DAO.Field
    
    Dim strFile As String
    Dim intFileNum As Integer
    
    ' Set your desired file path here

    strFile = "C:\Users\User\Desktop\��������\RelationshipInfo.txt"
    ' Get an available file number
    intFileNum = FreeFile
    
    ' Open (create) the file for output
    Open strFile For Output As #intFileNum
    
    Set db = CurrentDb
    
    For Each rel In db.Relations
        ' Skip system-defined relationships if they begin with "MSys"
        If Not (rel.Name Like "MSys*") Then
            
            Print #intFileNum, "Relationship: " & rel.Name
            Print #intFileNum, "  Table: " & rel.Table
            Print #intFileNum, "  Foreign Table: " & rel.ForeignTable
            
            ' Each Relation can have one or more matching fields
            For Each relField In rel.Fields
                Print #intFileNum, "    " & relField.Name & " -> " & relField.ForeignName
            Next relField
            
            Print #intFileNum,  ' Blank line for readability
        End If
    Next rel
    
    Close #intFileNum
    
    MsgBox "Relationship info exported to: " & strFile, vbInformation
End Sub
Private Function GetFieldTypeName(ByVal lngType As Long) As String

    Select Case lngType
        Case dbBoolean:     GetFieldTypeName = "Boolean"
        Case dbByte:        GetFieldTypeName = "Byte"
        Case dbInteger:     GetFieldTypeName = "Integer"
        Case dbLong:        GetFieldTypeName = "Long"
        Case dbCurrency:    GetFieldTypeName = "Currency"
        Case dbSingle:      GetFieldTypeName = "Single"
        Case dbDouble:      GetFieldTypeName = "Double"
        Case dbDate:        GetFieldTypeName = "Date/Time"
        Case dbText:        GetFieldTypeName = "Text"
        Case dbMemo:        GetFieldTypeName = "Memo"
        Case dbLongBinary:  GetFieldTypeName = "OLE Object"
        Case dbGUID:        GetFieldTypeName = "Replication ID"
        Case dbNumeric:     GetFieldTypeName = "Decimal"
        Case Else:          GetFieldTypeName = "Unknown"
    End Select
End Function

Public Sub ExportAllObjectsToText()
    Dim obj As AccessObject
    Dim exportPath As String
    Dim exportFile As String
    
    ' Where you want your files to go
    exportPath = "C:\Users\User\Desktop\��������\Access exports\"
    If Right(exportPath, 1) <> "\" Then exportPath = exportPath & "\"
    
    ' --- Standard modules ---
    For Each obj In CurrentProject.AllModules
        exportFile = exportPath & obj.Name & ".bas"
        Application.SaveAsText acModule, obj.Name, exportFile
    Next obj
    
    ' --- Forms (includes any code behind them) ---
    For Each obj In CurrentProject.AllForms
        exportFile = exportPath & obj.Name & ".txt"
        Application.SaveAsText acForm, obj.Name, exportFile
    Next obj
    
    ' --- Reports (includes any code behind them) ---
    For Each obj In CurrentProject.AllReports
        exportFile = exportPath & obj.Name & ".txt"
        Application.SaveAsText acReport, obj.Name, exportFile
    Next obj
    
    ' --- Macros ---
    For Each obj In CurrentProject.AllMacros
        exportFile = exportPath & obj.Name & ".txt"
        Application.SaveAsText acMacro, obj.Name, exportFile
    Next obj

    MsgBox "All object definitions and code have been exported to " & exportPath, vbInformation
End Sub