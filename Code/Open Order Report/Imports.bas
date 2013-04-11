Attribute VB_Name = "Imports"
Option Explicit

Sub ImportOOR(ISN As String)
    Dim PrevDispAlert As Boolean
    Dim sPath As String
    Dim i As Integer

    PrevDispAlert = Application.DisplayAlerts
    
    For i = 1 To 10
        sPath = "\\br3615gaps\gaps\3615 Open Order Report\ByInsideSalesNumber\" & ISN & "\" & Format(Date - i, "m-dd-yy") & " OOR.xlsx"

        If FileExists(sPath) Then
            Workbooks.Open sPath

            Sheets("117 BO").Select

            On Error Resume Next
            ActiveSheet.ShowAllData
            On Error GoTo 0
            ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Sheets("Previous 117 BO").Range("A1")

            Sheets("117 DS").Select
            On Error Resume Next
            ActiveSheet.ShowAllData
            On Error GoTo 0
            ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Sheets("Previous 117 DS").Range("A1")
        
            Application.DisplayAlerts = False
            ActiveWorkbook.Close
            Application.DisplayAlerts = PrevDispAlert
            Exit For
        End If
    Next
End Sub

Sub ImportContacts()
    Dim sPath As String
    Dim PrevStatus As Boolean

    sPath = "\\br3615gaps\gaps\Supplier Contacts\Supplier Contact Master.xlsx"

    Workbooks.Open sPath
    ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Sheets("Contacts").Range("A1")

    PrevStatus = Application.DisplayAlerts
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = PrevStatus

    FillInfo FunctionName:="ImportContacts", _
             Result:="Complete"
End Sub















