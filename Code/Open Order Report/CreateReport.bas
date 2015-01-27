Attribute VB_Name = "CreateReport"
Option Explicit

'---------------------------------------------------------------------------------------
' Proc : CreateOOR
' Date : 10/23/2014
' Desc : Creates Supplier Open Order Report
' Ex   : CreateOOR "BO"
'---------------------------------------------------------------------------------------
Sub CreateOOR(RepType As String)
    Dim RemCols As Variant      'Columns to remove
    Dim RepCols As Variant      'Correct report columns
    Dim TotalRows As Long
    Dim TotalCols As Integer
    Dim Lookup As String
    Dim UID As Variant
    Dim i As Integer

    If Sheets("117 " & RepType).Range("A1").Value = "" Then Exit Sub

    Sheets("117 " & RepType).Select
    TotalRows = Rows(Rows.Count).End(xlUp).Row
    TotalCols = Columns(Columns.Count).End(xlToLeft).Column
    RemCols = Array("WAREHOUSE", "ERROR", "REMOTE ORDER", "CYCLE", "STATUS", _
                    "ORDER DATE", "TAX", "TAX ACCOUNT", "CUSTOMER DELIVERY DATE (HR)", _
                    "CUSTOMER REFERENCE NO", "CUST PO LINE #", "CUSTOMER PART NUMBER", "SHIP TO", _
                    "IN", "OUT", "KIT", "AVAILABLE QTY", "QTY SHIPPED", "GROSS MARGIN", "LPST", _
                    "LGST", "UNIT PRICE", "DISCOUNT", "EXTENSION", "PRINT PICK TICKET DATE", "SHIP COMPLETE", _
                    "WIK QTY", "WIP QTY", "WIT QTY", "CUSTOMER ADDRESS 1", _
                    "CUSTOMER ADDRESS 2", "CUSTOMER CITY", "CUSTOMER STATE", "TRACK ID", "PALLET", "BOX", _
                    "QTY", "SUSPENSION TYPE", "MARGIN $", "EXT MARGIN $", "QUOTED TO")

    RepCols = Array("CUSTOMER", "ORDER NO", "LINE NO", "TYPE", "ITEM NUMBER", "CATALOG NUMBER", "ITEM DESCRIPTION", _
                    "SUOM", "ORDER QTY", "QTY TO SHIP", "BO QTY", "CUSTOMER DELIVERY DATE (LI)", "PO NUMBER", _
                    "PROMISE DATE", "OLD PROMISE DATE", "PO LINE NUM", "SUPPLIER NUM", "PURCHASE DATE", "CUSTOMER NAME", _
                    "COST", "EXT COST")

    'Remove header and footer
    Rows(TotalRows).Delete
    Rows(1).Delete

    'Remove unneeded columns
    For i = 0 To UBound(RemCols)
        DeleteColumn RemCols(i)
    Next

    'Verify columns
    TotalCols = Columns(Columns.Count).End(xlToLeft).Column
    If TotalCols > UBound(RepCols) + 1 Then Err.Raise 50001, "CreateOOR", "117 report has changed"
    
    For i = 1 To TotalCols
        If Cells(1, i).Value <> RepCols(i - 1) Then
            Err.Raise 50001, "CreateOOR", "117 report has changed"
        End If
    Next

    TotalRows = Rows(Rows.Count).End(xlUp).Row

    'Create UID
    Columns(1).Insert
    Range("A1").Value = "UID"
    Range("A2:A" & TotalRows).Formula = "=" & NumToCol(FindColumn("ORDER NO")) & "2&" & NumToCol(FindColumn("LINE NO")) & "2"
    Range("A2:A" & TotalRows).Value = Range("A2:A" & TotalRows).Value

    'Create SIMs
    Columns(2).Insert
    Range("B1").Value = "SIM"
    UID = NumToCol(FindColumn("ITEM NUMBER"))
    Lookup = "=""=""&""""""""&SUBSTITUTE(TRIM(" & UID & "2),""-"","""")&"""""""""
    Range("B2:B" & TotalRows).Formula = Lookup
    Range("B2:B" & TotalRows).Value = Range("B2:B" & TotalRows).Value

    'Add lead time column
    AddColumn "LEAD TIME", "=IFERROR(IF(VLOOKUP(B2,Gaps!A:AW,49,FALSE)=0,"""",VLOOKUP(B2,Gaps!A:AW,49,FALSE)),"""")"

    'Add estimated delivery date column
    AddEstDelDt

    If Sheets("Prev 117 " & RepType).Range("A1").Value <> "" Then
        'Add previous notes
        Sheets("Prev 117 " & RepType).Select
        UID = FindColumn("NOTES")
        Sheets("117 " & RepType).Select
        Lookup = "VLOOKUP(A2,'Prev " & "117 " & RepType & "'!A:" & NumToCol(UID) & "," & UID & ",FALSE)"
        AddColumn "NOTES", "=IFERROR(IF(" & Lookup & "=0,""""," & Lookup & "),"""")"

        'Add previous customer notes
        Sheets("Prev 117 " & RepType).Select
        UID = FindColumn("CUST NOTES")
        Sheets("117 " & RepType).Select
        Lookup = "VLOOKUP(A2,'Prev " & "117 " & RepType & "'!A:" & NumToCol(UID) & "," & UID & ",FALSE)"
        AddColumn "CUST NOTES", "=IFERROR(IF(" & Lookup & "=0,""""," & Lookup & "),"""")"
    Else
        AddColumn "NOTES", ""
        AddColumn "CUST NOTES", ""
    End If

    UID = NumToCol(FindColumn("SUPPLIER NUM")) & "2"
    AddColumn "SUPPLIER NAME", "=IFERROR(IF(VLOOKUP(TRIM(" & UID & "),'Supplier Master'!A:B,2,FALSE)=0,"""",VLOOKUP(TRIM(" & UID & "),'Supplier Master'!A:B,2,FALSE)),"""")"
    AddColumn "WESCO ACCT NO.", "=IFERROR(IF(VLOOKUP(TRIM(" & UID & "),'Supplier Master'!A:C,3,FALSE)=0,"""",VLOOKUP(TRIM(" & UID & "),'Supplier Master'!A:C,3,FALSE)),"""")"
    AddColumn "REP AGENCY", "=IFERROR(IF(VLOOKUP(TRIM(" & UID & "),'Supplier Master'!A:D,4,FALSE)=0,"""",VLOOKUP(TRIM(" & UID & "),'Supplier Master'!A:D,4,FALSE)),"""")"
    AddColumn "CONTACT NAME", "=IFERROR(IF(VLOOKUP(TRIM(" & UID & "),'Supplier Master'!A:F,6,FALSE)=0,"""",VLOOKUP(TRIM(" & UID & "),'Supplier Master'!A:F,6,FALSE)),"""")"
    AddColumn "CONTACT PHONE NO.", "=IFERROR(IF(VLOOKUP(TRIM(" & UID & "),'Supplier Master'!A:G,7,FALSE)=0,"""",VLOOKUP(TRIM(" & UID & "),'Supplier Master'!A:G,7,FALSE)),"""")"
    AddColumn "CONTACT E-MAIL", "=IFERROR(IF(VLOOKUP(TRIM(" & UID & "),'Supplier Master'!A:I,9,FALSE)=0,"""",VLOOKUP(TRIM(" & UID & "),'Supplier Master'!A:I,9,FALSE)),"""")"
    AddColumn "CONTACT FAX NO.", "=IFERROR(IF(VLOOKUP(TRIM(" & UID & "),'Supplier Master'!A:H,8,FALSE)=0,"""",VLOOKUP(TRIM(" & UID & "),'Supplier Master'!A:H,8,FALSE)),"""")"
    AddColumn "ORD MIN", "=IFERROR(IF(VLOOKUP(TRIM(" & UID & "),'Supplier Master'!A:N,14,FALSE)=0,"""",VLOOKUP(TRIM(" & UID & "),'Supplier Master'!A:N,14,FALSE)),"""")"
    AddColumn "FRT ALLOWED AMT", "=IFERROR(IF(VLOOKUP(TRIM(" & UID & "),'Supplier Master'!A:O,15,FALSE)=0,"""",VLOOKUP(TRIM(" & UID & "),'Supplier Master'!A:O,15,FALSE)),"""")"
    AddColumn "SPECIAL INSTRUCTIONS", "=IFERROR(IF(VLOOKUP(TRIM(" & UID & "),'Supplier Master'!A:Q,17,FALSE)=0,"""",VLOOKUP(TRIM(" & UID & "),'Supplier Master'!A:Q,17,FALSE)),"""")"

    'Remove SIM
    Columns(2).Delete

    'Remove UID
    Columns(1).Delete
End Sub

Sub CreateCustOOR(RepType As String)
    Dim RemCols As Variant      'Columns to remove
    Dim RepCols As Variant      'Correct report columns
    Dim TotalRows As Long
    Dim TotalCols As Integer
    Dim Lookup As String
    Dim UID As Variant
    Dim i As Integer

    If Sheets("117 " & RepType).Range("A1").Value = "" Then Exit Sub

    Sheets("117 " & RepType).Select
    TotalRows = Rows(Rows.Count).End(xlUp).Row
    TotalCols = Columns(Columns.Count).End(xlToLeft).Column

    'Remove header and footer
    Rows(TotalRows).Delete
    Rows(1).Delete

    RemCols = Array("WAREHOUSE", "ERROR", "REMOTE ORDER", "STATUS", _
                    "ORDER DATE", "TAX", "TAX ACCOUNT", "CUSTOMER DELIVERY DATE (HR)", _
                    "SHIP TO", "IN", "OUT", "KIT", "GROSS MARGIN", "LPST", "LGST", _
                    "UNIT PRICE", "DISCOUNT", "EXTENSION", "PRINT PICK TICKET DATE", _
                    "SHIP COMPLETE", "PO NUMBER", "PROMISE DATE", "OLD PROMISE DATE", _
                    "PO LINE NUM", "SUPPLIER NUM", "PURCHASE DATE", "WIK QTY", "WIP QTY", _
                    "WIT QTY", "CUSTOMER ADDRESS 1", "CUSTOMER ADDRESS 2", _
                    "CUSTOMER CITY", "CUSTOMER STATE", "TRACK ID", "PALLET", "BOX", "QTY", _
                    "SUSPENSION TYPE", "COST", "EXT COST", "MARGIN $", "EXT MARGIN $", "QUOTED TO")

    RepCols = Array("CUSTOMER", "ORDER NO", "CYCLE", "CUSTOMER REFERENCE NO", "CUST PO LINE #", _
                    "CUSTOMER PART NUMBER", "LINE NO", "TYPE", "ITEM NUMBER", "CATALOG NUMBER", _
                    "ITEM DESCRIPTION", "SUOM", "ORDER QTY", "AVAILABLE QTY", "QTY TO SHIP", _
                    "BO QTY", "QTY SHIPPED", "CUSTOMER DELIVERY DATE (LI)", "CUSTOMER NAME")

    'Remove unneeded columns
    For i = 0 To UBound(RemCols)
        DeleteColumn RemCols(i)
    Next

    'Verify columns
    TotalCols = Columns(Columns.Count).End(xlToLeft).Column
    For i = 1 To TotalCols
        If Cells(1, i).Value <> RepCols(i - 1) Then
            Err.Raise 50001, "Format117", "117 report has changed"
        End If
    Next

    'Create UID
    Columns(1).Insert
    Range("A1").Value = "UID"
    Range("A2:A" & TotalRows).Formula = "=" & NumToCol(FindColumn("ORDER NO")) & "2&" & NumToCol(FindColumn("LINE NO")) & "2"
    Range("A2:A" & TotalRows).Value = Range("A2:A" & TotalRows).Value

    If Sheets("Prev 117 " & RepType).Range("A1").Value <> "" Then
        'Add previous customer notes
        Sheets("Prev 117 " & RepType).Select
        UID = FindColumn("CUST NOTES")
        Sheets("117 " & RepType).Select
        Lookup = "VLOOKUP(A2,'Prev " & "117 " & RepType & "'!A:" & NumToCol(UID) & "," & UID & ",FALSE)"
        AddColumn "CUST NOTES", "=IFERROR(IF(" & Lookup & "=0,""""," & Lookup & "),"""")"
    Else
        AddColumn "CUST NOTES", ""
    End If

    'Remove UID
    Columns(1).Delete
End Sub

Private Sub AddEstDelDt()
    Dim EstDelCol As Integer
    Dim TotalRows As Long
    Dim PromCol As String
    Dim PurchCol As String
    Dim LeadCol As String
    Dim i As Long

    TotalRows = Rows(Rows.Count).End(xlUp).Row
    EstDelCol = FindColumn("CUSTOMER DELIVERY DATE (LI)") + 1
    Columns(EstDelCol).Insert xlToLeft
    PromCol = NumToCol(FindColumn("PROMISE DATE"))
    PurchCol = NumToCol(FindColumn("PURCHASE DATE"))
    LeadCol = NumToCol(FindColumn("LEAD TIME"))

    'Add column header
    Cells(1, EstDelCol).Value = "EST DELIVERY DT"

    For i = 2 To TotalRows
        If Trim(Range(PromCol & i).Value) <> "" Then
            'If there is a promise date set it as the estimated delivery date
            Cells(i, EstDelCol).Value = Range(PromCol & i).Value
        ElseIf Range(LeadCol & i).Value <> "" And _
               Trim(Range(PurchCol & i).Value) <> "" Then
            'If there is a lead time and a purchase date
            'add the two to create and estimated delivery date
            'Floor((LT / 7))*2 is added since items dont ship on weekends
            Cells(i, EstDelCol).Value = CDate(Range(PurchCol & i).Value) + Range(LeadCol & i).Value + WorksheetFunction.RoundDown(Range(LeadCol & i) / 7, 0) * 2
        ElseIf Trim(Range(PurchCol & i).Value) <> "" Then
            'If there is a purchase date add 14 days
            'and use that as the estimated delivery date
            Cells(i, EstDelCol).Value = CDate(Range(PurchCol & i).Value) + 14
        End If
    Next

    Range(Cells(2, EstDelCol), Cells(TotalRows, EstDelCol)).NumberFormat = "m/d/yyyy"
End Sub

Private Sub AddColumn(Header As String, Optional Lookup As String)
    Dim TotalCols As Integer
    Dim TotalRows As Long

    TotalCols = Columns(Columns.Count).End(xlToLeft).Column + 1
    TotalRows = Rows(Rows.Count).End(xlUp).Row

    Cells(1, TotalCols).Value = Header
    Range(Cells(2, TotalCols), Cells(TotalRows, TotalCols)).Formula = Lookup
    Range(Cells(2, TotalCols), Cells(TotalRows, TotalCols)).Value = Range(Cells(2, TotalCols), Cells(TotalRows, TotalCols)).Value
End Sub

Private Function NumToCol(ByVal Col As Integer) As String
    NumToCol = Split(Columns(Col).Address(False, False), ":")(0)
End Function
