Attribute VB_Name = "Program"
Option Explicit

'Updater variables
Public Const VersionNumber As String = "2.0.5"
Public Const RepositoryName As String = "Open_Order_Report"

'Variables set by FrmImport117
Public Canceled As Boolean
Public sBranch As String
Public sSequence As String
Public sISN As String

Sub CreateExpedite()
    'Import 117 report
    FrmImport117.Show

    If Canceled = True Then GoTo IMPORT_CANCELED

    Application.ScreenUpdating = False

    'If not canceled but no reports could be found exit the macro
    If Sheets("117 DS").Range("A1").Value & _
       Sheets("117 BO").Range("A1").Value = "" Then
        Exit Sub
    End If

    On Error GoTo MAIN_ERR
    'Import supplier contact master
    ImportSupplierContacts Sheets("Supplier Master").Range("A1"), sBranch

    'Import previous open order report
    ImportPrevOOR

    'Import gaps
    ImportGaps Sheets("Gaps").Range("A1"), True, sBranch

    'Create open order report
    CreateOOR "BO"
    CreateOOR "DS"
    On Error GoTo 0

    'Add formatting to open order report
    FormatOOR "BO"
    FormatOOR "DS"

    'Save and email reports
    ExportOOR

    Clean
    Application.ScreenUpdating = True

    MsgBox "Complete!"
    Exit Sub

IMPORT_CANCELED:
    Clean
    Debug.Print Err.Description
    Application.ScreenUpdating = True
    Exit Sub

MAIN_ERR:
    MsgBox Err.Description, vbOKOnly, Err.Source
    Clean
    Application.ScreenUpdating = True
    Exit Sub
End Sub

Sub CreateCustExpedite()
    FrmImport117.Show

    If Canceled = True Then GoTo IMPORT_CANCELED

    'If not canceled but no reports could be found exit the macro
    If Sheets("117 DS").Range("A1").Value & _
       Sheets("117 BO").Range("A1").Value = "" Then
        Exit Sub
    End If

    'Import supplier contact master
    ImportSupplierContacts Sheets("Supplier Master").Range("A1"), sBranch

    'Import previous open order report
    ImportPrevOOR

    CreateCustOOR "BO"
    CreateCustOOR "DS"

    FormatCustOOR "BO"
    FormatCustOOR "DS"

    'Save and email
    ExportCustOOR

    Clean
    MsgBox "Complete!"
    Exit Sub

IMPORT_CANCELED:
    Clean
    Debug.Print Err.Description
End Sub

'---------------------------------------------------------------------------------------
' Proc : Clean
' Date : 10/20/2014
' Desc : Remove all data from the macro workbook
'---------------------------------------------------------------------------------------
Sub Clean()
    Dim s As Worksheet

    ThisWorkbook.Activate
    For Each s In ThisWorkbook.Sheets
        If s.Name <> "Macro" Then
            s.Select
            s.AutoFilterMode = False
            s.Columns.Hidden = False
            s.Rows.Hidden = False
            s.Cells.Delete
            s.Range("A1").Select
        End If
    Next

    Sheets("Macro").Select
    Range("C7").Select
End Sub
