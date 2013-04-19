Attribute VB_Name = "Program"
Option Explicit

Sub Main()
    Dim ISN As String
    Dim Cancel As Boolean
    Dim ImportCheck As String

    Application.ScreenUpdating = False
    'Import 117 Report
    On Error GoTo ImportErr
    ISN = InputBox("Inside Sales Number:", "Please enter the ISN#")
    If ISN = "" Then Cancel = True

    Import117byISN ReportType.BO, Sheets("117 BO").Range("A1"), ISN, Cancel
    Import117byISN ReportType.DS, Sheets("117 DS").Range("A1"), ISN, Cancel
    On Error GoTo 0

    ImportCheck = Sheets("117 BO").Range("A1") & Sheets("117 DS").Range("A1")

    If ImportCheck <> "" Then
        ImportSupplierContacts
        ImportSalesContacts
        ImportOOR ISN

        Format117 "117 DS"
        Format117 "117 BO"

        If Sheets("117 BO").Range("A1").Value <> "" Then
            Sheets("117 BO").Select
        Else
            Sheets("117 DS").Select
        End If
    End If

    Application.ScreenUpdating = True

ImportErr:
    Exit Sub
End Sub

Sub SendMail()
    Dim ISN As String
    Dim EmailAddress As String
    Dim FileName As String
    Dim i As Long

    Sheets("117 BO").Select
    
    On Error Resume Next
    ISN = Cells(2, FindColumn("IN")).Value
    On Error GoTo 0
    
    FileName = Format(Date, "m-dd-yy") & " OOR.xlsx"

    If ISN = "" Then
        Sheets("117 DS").Select
        ISN = Cells(2, FindColumn("IN")).Value
    End If

    Sheets("Sales Contacts").Select
    For i = 2 To ActiveSheet.UsedRange.Rows.Count
        If Cells(i, 1).Value = ISN Then
            EmailAddress = Cells(i, 2).Value
            Exit For
        End If
    Next

    If EmailAddress = "" Then
        MsgBox Prompt:="Email for sales number " & ISN & " could not be found."
    Else
        Email SendTo:=EmailAddress, _
              Subject:="Open Order Report", _
              Body:="Please click the link to view the status of your open POs" & "<br><br>" & _
                    """\\br3615gaps\gaps\3615 Open Order Report\ByInsideSalesNumber\" & ISN & "\" & FileName & """"
    End If
End Sub

Sub Clean()
    Dim s As Variant

    For Each s In ThisWorkbook.Sheets
        If s.Name <> "Macro" Then
            s.Cells.Delete
        End If
    Next
End Sub
