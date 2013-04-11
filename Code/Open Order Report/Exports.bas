Attribute VB_Name = "Exports"
Option Explicit

Sub Export117()
    Dim Wkbk As Workbook
    Dim sPath As String
    Dim ISN As String
    
    ISN = Sheets("117 BO").Cells(2, FindColumn("IN")).Value
    sPath = "\\br3615gaps\gaps\3615 Open Order Report\ByInsideSalesNumber\" & ISN & "\" & Format(Date, "m-dd-yy") & " OOR.xlsx"
    
    Sheets("117 BO").Copy
    Set Wkbk = ActiveWorkbook
    ThisWorkbook.Sheets("117 DS").Copy After:=Wkbk.Sheets(Wkbk.Sheets.Count)
    
    ActiveWorkbook.SaveAs sPath
End Sub
