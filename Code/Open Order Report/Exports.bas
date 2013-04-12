Attribute VB_Name = "Exports"
Option Explicit

Sub Export117()
    Dim Wkbk As Workbook
    Dim sPath As String
    Dim FileName As String
    Dim ISN As String
    Dim PrevDispAlert As Boolean
    Dim PrevSheet As Worksheet
    Dim iCol As Integer

    Set PrevSheet = ActiveSheet

    Sheets("117 DS").Select
    PrevDispAlert = Application.DisplayAlerts

    iCol = FindColumn("IN")
    If iCol <> 0 Then
        ISN = Sheets("117 BO").Cells(2, FindColumn("IN")).Value
    End If
    
    Sheets("117 BO").Select
    If iCol <> 0 Then
        ISN = Sheets("117 DS").Cells(2, FindColumn("IN")).Value
    End If

    If ISN <> "" Then
        FileName = Format(Date, "m-dd-yy") & " OOR.xlsx"
        sPath = "\\br3615gaps\gaps\3615 Open Order Report\ByInsideSalesNumber\" & ISN & "\"

        Sheets("117 BO").Copy
        Set Wkbk = ActiveWorkbook
        ThisWorkbook.Sheets("117 DS").Copy After:=Wkbk.Sheets(Wkbk.Sheets.Count)

        If FolderExists(sPath) = False Then
            MkDir sPath
        End If

        On Error GoTo SAVE_ERR
        ActiveWorkbook.SaveAs sPath & FileName
        On Error GoTo 0
        Application.DisplayAlerts = False
        ActiveWorkbook.Close
    End If

    Application.DisplayAlerts = PrevDispAlert
    PrevSheet.Select

    Exit Sub

SAVE_ERR:
    If Err.Description = "Cannot access '" & FileName & "'." Then
        MsgBox Prompt:=Err.Description
    End If
    Resume Next
End Sub
