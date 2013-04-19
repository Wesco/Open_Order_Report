Attribute VB_Name = "FormatData"
Option Explicit

Sub Format117(SheetName As String)
    Dim PrevSheet As Worksheet
    Dim iCol As Integer
    Dim iRows As Long

    Set PrevSheet = ActiveSheet
    Sheets(SheetName).Select

    If Range("A1").Value <> "" Then
        Rows(ActiveSheet.UsedRange.Rows.Count).Delete
        Rows(1).Delete

        DeleteColumn "QUOTED TO"
        DeleteColumn "EXT MARGIN $"
        DeleteColumn "MARGIN $"
        DeleteColumn "EXT COST"
        DeleteColumn "COST"
        DeleteColumn "SUSPENSION TYPE"
        DeleteColumn "QTY"
        DeleteColumn "BOX"
        DeleteColumn "PALLET"
        DeleteColumn "TRACK ID"
        DeleteColumn "CUSTOMER STATE"
        DeleteColumn "CUSTOMER CITY"
        DeleteColumn "CUSTOMER ADDRESS 2"
        DeleteColumn "CUSTOMER ADDRESS 1"
        DeleteColumn "CUSTOMER NAME"
        DeleteColumn "WIT QTY"
        DeleteColumn "WIP QTY"
        DeleteColumn "WIK QTY"
        DeleteColumn "PURCHASE DATE"
        DeleteColumn "OLD PROMISE DATE"
        DeleteColumn "SHIP COMPLETE"
        DeleteColumn "SHIP DATE"
        DeleteColumn "EXTENSION"
        DeleteColumn "REQUIRED DATE (LI)"
        DeleteColumn "DISCOUNT"
        DeleteColumn "UNIT PRICE"
        DeleteColumn "LGST"
        DeleteColumn "LPST"
        DeleteColumn "GROSS MARGIN"
        DeleteColumn "SUOM"
        DeleteColumn "CATALOG NUMBER"
        DeleteColumn "TYPE"
        DeleteColumn "KIT"
        DeleteColumn "OUT"
        DeleteColumn "SHIP TO"
        DeleteColumn "CUSTOMER PART NUMBER"
        DeleteColumn "CUST PO LINE #"
        DeleteColumn "REQUIRED DATE (HR)"
        DeleteColumn "TAX ACCOUNT"
        DeleteColumn "TAX"
        DeleteColumn "CYCLE"
        DeleteColumn "REMOTE ORDER"
        DeleteColumn "ERROR"
        DeleteColumn "WAREHOUSE"
        DeleteColumn "STATUS"

        iCol = FindColumn("SUPPLIER NUM")
        iRows = ActiveSheet.UsedRange.Rows.Count
        Range(Cells(2, iCol), Cells(iRows, iCol)).NumberFormat = "@"
        Range(Cells(2, iCol), Cells(iRows, iCol)).Value = Range(Cells(2, iCol), Cells(iRows, iCol)).Value

        ActiveSheet.ListObjects.Add(xlSrcRange, ActiveSheet.UsedRange, , xlYes).Name = "Table1"

        iCol = ActiveSheet.UsedRange.Columns.Count + 1
        Cells(1, iCol).Value = "UID"
        Cells(2, iCol).Formula = "=[@[ORDER NO]]&[@[LINE NO]]"
        Range(Cells(2, iCol), Cells(iRows, iCol)).Value = Range(Cells(2, iCol), Cells(iRows, iCol)).Value

        iCol = ActiveSheet.UsedRange.Columns.Count + 1
        Cells(1, iCol).Value = "Email"
        Cells(2, iCol).Formula = "=IFERROR(VLOOKUP(TRIM([@[SUPPLIER NUM]]),'Supplier Contacts'!A:B,2,FALSE),"""")"
        Range(Cells(2, iCol), Cells(iRows, iCol)).Value = Range(Cells(2, iCol), Cells(iRows, iCol)).Value

        iCol = ActiveSheet.UsedRange.Columns.Count + 1
        Cells(1, iCol).Value = "Notes"
        Cells(2, iCol).Formula = _
        "=IFERROR(IF(VLOOKUP([@UID],'Previous " & SheetName & "'!R:T,3,FALSE)=0,"""",VLOOKUP([@UID],'Previous " & SheetName & "'!R:T,3,FALSE)),"""")"
        Range(Cells(2, iCol), Cells(iRows, iCol)).Value = Range(Cells(2, iCol), Cells(iRows, iCol)).Value

        ActiveSheet.UsedRange.Columns.AutoFit
    End If
    PrevSheet.Select
End Sub



















