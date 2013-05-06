Attribute VB_Name = "Imports"
Option Explicit

Sub ImportOOR(ISN As String)
    Dim Wkbk As Workbook
    Dim PrevDispAlert As Boolean
    Dim sPath As String
    Dim i As Integer

    PrevDispAlert = Application.DisplayAlerts

    For i = 0 To 10
        sPath = "\\br3615gaps\gaps\3615 Open Order Report\ByInsideSalesNumber\" & ISN & "\" & Format(Date - i, "m-dd-yy") & " OOR.xlsx"

        If FileExists(sPath) Then
            Workbooks.Open sPath
            Set Wkbk = ActiveWorkbook

            Sheets("117 BO").Select
            On Error Resume Next
            ActiveSheet.AutoFilter.ShowAllData
            ActiveSheet.ShowAllData
            On Error GoTo 0
            ActiveSheet.Name = "Previous 117 BO"
            ActiveSheet.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
            Wkbk.Activate

            Sheets("117 DS").Select
            On Error Resume Next
            ActiveSheet.AutoFilter.ShowAllData
            ActiveSheet.ShowAllData
            On Error GoTo 0
            ActiveSheet.Name = "Previous 117 DS"
            ActiveSheet.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
            Wkbk.Activate

            Application.DisplayAlerts = False
            ActiveWorkbook.Close
            Application.DisplayAlerts = PrevDispAlert
            Exit For
        End If
    Next
End Sub

Sub ImportSupplierContacts()
    Dim sPath As String
    Dim PrevStatus As Boolean

    sPath = "\\br3615gaps\gaps\Contacts\Supplier Contact Master.xlsx"

    Workbooks.Open sPath
    ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Sheets("Supplier Contacts").Range("A1")

    PrevStatus = Application.DisplayAlerts
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = PrevStatus
End Sub

Sub ImportSalesContacts()
    Dim sPath As String
    Dim PrevStatus As Boolean

    sPath = "\\br3615gaps\gaps\Contacts\Sales #s.xlsx"

    Workbooks.Open sPath
    ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Sheets("Sales Contacts").Range("A1")

    PrevStatus = Application.DisplayAlerts
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = PrevStatus
End Sub













