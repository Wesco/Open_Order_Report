Attribute VB_Name = "FormatData"
Option Explicit

Sub Format117(SheetName As String)
    Dim PrevSheet As Worksheet
    Dim iCol As Integer
    Dim iRows As Long
    Dim i As Long

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
        
        On Error GoTo COL_ERR
        If iCol = 0 Then Err.Raise 50000
        On Error GoTo 0
        
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

        iCol = ActiveSheet.UsedRange.Columns.Count + 1
        Cells(1, iCol).Value = "Address"
        Cells(2, iCol).Formula = "=IFERROR(CELL(""address"",INDEX('Previous " & SheetName & "'!R:R,MATCH([@UID],'Previous " & SheetName & "'!R:R,0),1)),"""")"
        Range(Cells(2, iCol), Cells(iRows, iCol)).Value = Range(Cells(2, iCol), Cells(iRows, iCol)).Value

        iCol = ActiveSheet.UsedRange.Columns.Count + 1
        Cells(1, iCol).Value = "Cell"
        Cells(2, iCol).Formula = "=RIGHT([@Address],LEN([@Address]) -" & Len(ActiveWorkbook.Name) + Len("Previous ") + Len(SheetName) + 5 & ")"
        Range(Cells(2, iCol), Cells(iRows, iCol)).Value = Range(Cells(2, iCol), Cells(iRows, iCol)).Value

        iCol = ActiveSheet.UsedRange.Columns.Count + 1
        Cells(1, iCol).Value = "Absolute"
        Cells(2, iCol).Formula = "=SUBSTITUTE([@Cell],""$"","""")"
        Range(Cells(2, iCol), Cells(iRows, iCol)).Value = Range(Cells(2, iCol), Cells(iRows, iCol)).Value

        Columns("U:V").Delete

        On Error Resume Next
        For i = 2 To iRows
            If Sheets("Previous " & SheetName).Range(Sheets(SheetName).Cells(i, 21).Value).Interior.Color <> "16777215" Then
                Range(Cells(i, 1), Cells(i, 20)).Interior.Color = _
                Sheets("Previous " & SheetName).Range(Sheets(SheetName).Cells(i, 21).Value).Interior.Color
            End If
        Next
        On Error GoTo 0

        Columns("U:U").Delete

        ActiveSheet.UsedRange.Columns.AutoFit

        With Sheets(SheetName).ListObjects("Table1").Sort
            .SortFields.Clear
            .SortFields.Add Key:=Range("Table1[[#All],[PO NUMBER]]"), _
                            SortOn:=xlSortOnValues, _
                            Order:=xlAscending, _
                            DataOption:=xlSortNormal
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End If

    PrevSheet.Select
Exit Sub

COL_ERR:
MsgBox "Column ""SUPPLIER NUM"" on " & ActiveSheet.Name & " could not be found.", vbOKOnly, "Error"
End Sub



















