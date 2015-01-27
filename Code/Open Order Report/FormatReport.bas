Attribute VB_Name = "FormatReport"
Option Explicit

Sub FormatOOR(RepType As String)
    Dim TotalRows As Long
    Dim TotalCols As Integer
    Dim CustDelCol As String
    Dim EstDelCol As String
    Dim DateDiff As Integer
    Dim MovCols As Variant
    Dim i As Long

    If Sheets("117 " & RepType).Range("A1").Value = "" Then Exit Sub

    Sheets("117 " & RepType).Select
    TotalRows = Rows(Rows.Count).End(xlUp).Row
    TotalCols = Columns(Columns.Count).End(xlToLeft).Column
    CustDelCol = Split(Columns(FindColumn("CUSTOMER DELIVERY DATE (LI)")).Address(False, False), ":")(0)
    EstDelCol = Split(Columns(FindColumn("EST DELIVERY DT")).Address(False, False), ":")(0)
    MovCols = Array(Array("PROMISE DATE", "EST DELIVERY DT"), Array("", ""))

    'Add alternating line colors
    Sheets("117 " & RepType).ListObjects.Add xlSrcRange, Range(Cells(1, 1), Cells(TotalRows, TotalCols)), False, xlYes
    Sheets("117 " & RepType).ListObjects(1).Unlist

    For i = 2 To TotalRows
        If Trim(Range(CustDelCol & i).Value) <> "" And _
           Range(EstDelCol & i).Value <> "" Then
           
            On Error GoTo CDATE_ERR
            DateDiff = CDate(Trim(Range(CustDelCol & i).Value)) - CDate(Range(EstDelCol & i).Value)
            On Error GoTo 0

            If DateDiff <= 0 Then
                Range(EstDelCol & i).Interior.Color = RGB(230, 0, 0)
            ElseIf DateDiff <= 3 Then
                Range(EstDelCol & i).Interior.Color = RGB(255, 255, 0)
            End If
        Else
            Range(EstDelCol & i).Interior.Color = RGB(230, 0, 0)
        End If
        
        If CDate(Range(EstDelCol & i).Value) < Now Then
            Range(EstDelCol & i).Interior.Color = RGB(230, 0, 0)
        End If
    Next
    Exit Sub
    
CDATE_ERR:
    DateDiff = 0
    Resume Next
End Sub

Sub FormatCustOOR(RepType As String)
    Dim TotalRows As Long
    Dim TotalCols As Integer

    If Sheets("117 " & RepType).Range("A1").Value = "" Then Exit Sub

    Sheets("117 " & RepType).Select
    TotalRows = Rows(Rows.Count).End(xlUp).Row
    TotalCols = Columns(Columns.Count).End(xlToLeft).Column

    'Add alternating line colors
    Sheets("117 " & RepType).ListObjects.Add xlSrcRange, Range(Cells(1, 1), Cells(TotalRows, TotalCols)), False, xlYes
    Sheets("117 " & RepType).ListObjects(1).Unlist
End Sub
