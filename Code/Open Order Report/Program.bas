Attribute VB_Name = "Program"
Option Explicit

Sub Main()
    Dim ISN As String
    Dim Cancel As Boolean

    'Import 117 Report
    On Error GoTo ImportErr
    ISN = InputBox("Inside Sales Number:", "Please enter the ISN#")
    If ISN = "" Then Cancel = True
    Import117byISN ReportType.BO, Sheets("117 BO").Range("A1"), ISN, Cancel
    Import117byISN ReportType.DS, Sheets("117 DS").Range("A1"), ISN, Cancel
    On Error GoTo 0
    
    Format117 "117 DS"
    Format117 "117 BO"
    

ImportErr:
    Exit Sub
End Sub

Sub Clean()
    Dim s As Variant

    For Each s In ThisWorkbook.Sheets
        If s.Name <> "Macro" Then
            s.Cells.Delete
        End If
    Next
End Sub
