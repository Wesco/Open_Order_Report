Attribute VB_Name = "Imports"
Option Explicit

Sub ImportPrevOOR()
    Dim FilePath As String
    Dim FileName As String
    Dim Colorder As Variant
    Dim i As Integer
    Dim j As Integer

    FilePath = "\\br3615gaps\gaps\" & sBranch & " Open Order Report\" & sSequence & "\" & sISN & "\"
    Colorder = Array("UID", "CUSTOMER", "ORDER NO", "LINE NO", "TYPE", "ITEM NUMBER", "CATALOG NUMBER", "ITEM DESCRIPTION", _
                     "SUOM", "ORDER QTY", "QTY TO SHIP", "BO QTY", "CUSTOMER DELIVERY DATE (LI)", "PO NUMBER", _
                     "PROMISE DATE", "OLD PROMISE DATE", "PO LINE NUM", "SUPPLIER NUM", "PURCHASE DATE", "CUSTOMER NAME", _
                     "COST", "EXT COST", "LEAD TIME", "EST DELIVERY DT", "NOTES", "CUST NOTES", "SUPPLIER NAME", _
                     "WESCO ACCT NO.", "REP AGENCY", "CONTACT NAME", "CONTACT PHONE NO.", "CONTACT E-MAIL", _
                     "CONTACT FAX NO.", "ORD MIN", "FRT ALLOWED AMT", "SPECIAL INSTRUCTIONS")

    'Import the most recent BO OOR run in the last 90 days
    If Sheets("117 BO").Range("A1").Value <> "" Then
        For i = 0 To 90
            FileName = Format(Date - i, "yyyy-mm-dd") & " BO OOR.xlsx"
            If FileExists(FilePath & FileName) Then
                'UID column is added during import
                ImportFile FilePath, FileName, Sheets("Prev 117 BO").Range("A1")
                Sheets("Prev 117 BO").Select
                For j = 0 To UBound(Colorder)
                    FindColumn Colorder(i)
                Next
                Exit For
            End If
        Next
    End If

    'Import the most recent DS OOR run in the last 90 days
    If Sheets("117 DS").Range("A1").Value <> "" Then
        For i = 0 To 90
            FileName = Format(Date - i, "yyyy-mm-dd") & " DS OOR.xlsx"
            If FileExists(FilePath & FileName) Then
                'UID column is added during import
                ImportFile FilePath, FileName, Sheets("Prev 117 DS").Range("A1")
                Sheets("Prev 117 DS").Select
                For j = 0 To UBound(Colorder)
                    FindColumn Colorder(i)
                Next
                Exit For
            End If
        Next
    End If
End Sub

Sub ImportPrevCustOOR()
    Dim Colorder As Variant
    Dim FilePath As String
    Dim FileName As String
    Dim i As Integer
    Dim j As Integer

    FilePath = "\\br3615gaps\gaps\" & sBranch & " Open Order Report\" & sSequence & "\" & sISN & "\"
    Colorder = Array("UID", "ORDER NO", "CYCLE", "CUSTOMER REFERENCE NO", "CUST PO LINE #", "CUSTOMER PART NUMBER", _
                     "LINE NO", "TYPE", "ITEM NUMBER", "CATALOG NUMBER", "ITEM DESCRIPTION", "SUOM", "ORDER QTY", _
                     "AVAILABLE QTY", "QTY TO SHIP", "BO QTY", "QTY SHIPPED", "CUSTOMER DELIVERY DATE (LI)", "Notes")

    'Import the most recent CUST BO OOR run in the last 90 days
    If Sheets("117 BO").Range("A1").Value <> "" Then
        For i = 0 To 90
            FileName = Format(Date - i, "yyyy-mm-dd") & " CUST BO OOR.xlsx"
            If FileExists(FilePath & FileName) Then
                'UID column is added during import
                ImportFile FilePath, FileName, Sheets("Prev Cust BO").Range("A1")
                Sheets("Prev 117 BO").Select
                For j = 0 To UBound(Colorder)
                    If Cells(1, i + 1).Value <> Colorder(i) Then
                        Err.Raise CustErr.COLNOTFOUND, "ImportPrevOOR", "Column " & Colorder(i) & " was moved or is missing."
                    End If
                Next
                Exit For
            End If
        Next
    End If

    'Import the most recent CUST DS OOR run in the last 90 days
    If Sheets("117 DS").Range("A1").Value <> "" Then
        For i = 0 To 90
            FileName = Format(Date - i, "yyyy-mm-dd") & " CUST DS OOR.xlsx"
            If FileExists(FilePath & FileName) Then
                'UID column is added during import
                ImportFile FilePath, FileName, Sheets("Prev Cust DS").Range("A1")
                Sheets("Prev 117 DS").Select
                For j = 0 To UBound(Colorder)
                    If Cells(1, i + 1).Value <> Colorder(i) Then
                        Err.Raise CustErr.COLNOTFOUND, "ImportPrevOOR", "Column " & Colorder(i) & " was moved or is missing."
                    End If
                Next
                Exit For
            End If
        Next
    End If
End Sub

Private Sub ImportFile(FilePath As String, FileName As String, Destination As Range)
    Dim PrevDispAlert As Boolean
    Dim TotalRows As Long

    Workbooks.Open FilePath & FileName

    'Make sure all data is visible
    ActiveSheet.AutoFilterMode = False
    ActiveSheet.Columns.Hidden = False
    ActiveSheet.Rows.Hidden = False
    TotalRows = Rows(Rows.Count).End(xlUp).Row

    'Add UID column
    Columns(1).Insert
    Range("A1").Value = "UID"
    Range("A2:A" & TotalRows).Formula = "=" & NumToCol(FindColumn("ORDER NO")) & "2&" & NumToCol(FindColumn("LINE NO")) & "2"
    Range("A2:A" & TotalRows).Value = Range("A2:A" & TotalRows).Value

    ActiveSheet.UsedRange.Copy Destination:=Destination

    PrevDispAlert = Application.DisplayAlerts
    Application.DisplayAlerts = False
    ActiveWorkbook.Saved = True
    ActiveWorkbook.Close
    Application.DisplayAlerts = PrevDispAlert
End Sub

Private Function NumToCol(ByVal Col As Integer) As String
    NumToCol = Split(Columns(Col).Address(False, False), ":")(0)
End Function
