Attribute VB_Name = "FormatData"
Option Explicit

Sub Format117(SheetName As String)
    Dim PrevSheet As Worksheet
    
    Set PrevSheet = ActiveSheet
    Sheets(SheetName).Select
    
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
    
    PrevSheet.Select
End Sub

Sub Format473()
    Rows(1).Delete
End Sub
    
