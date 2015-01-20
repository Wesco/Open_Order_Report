Attribute VB_Name = "Exports"
Option Explicit

Sub ExportOOR()
    Dim FilePath As String
    Dim FileName As String
    Dim EmailBody As String
    Dim Saved As Boolean

    FilePath = "\\br3615gaps\gaps\" & sBranch & " Open Order Report\" & sSequence & "\" & sISN & "\"

    If Sheets("117 DS").Range("A1").Value <> "" Then
        FileName = Format(Date, "yyyy-mm-dd") & " DS OOR.xlsx"
        Saved = SaveWorkbook(Sheets("117 DS"), FilePath, FileName)
        If Saved = True Then
            EmailBody = "<a href=""file:///" & FilePath & FileName & """>DS Report</a>"
        End If
    End If

    If Sheets("117 BO").Range("A1").Value <> "" Then
        FileName = Format(Date, "yyyy-mm-dd") & " BO OOR.xlsx"
        Saved = SaveWorkbook(Sheets("117 BO"), FilePath, FileName)
        If Saved = True Then
            If EmailBody = "" Then
                EmailBody = "<a href=""file:///" & FilePath & FileName & """>BO Report</a>"
            Else
                EmailBody = EmailBody & "<br><a href=""file://" & FilePath & FileName & """>BO Report</a>"
            End If
        End If
    End If
    Email Environ("username") & "@wesco.com", _
          Subject:="Open Order Report", _
          Body:=EmailBody, _
          MailType:=SMTP
End Sub

Sub ExportCustOOR()
    Dim FilePath As String
    Dim FileName As String
    Dim EmailBody As String
    Dim Files As Variant
    Dim Saved As Boolean

    FilePath = "\\br3615gaps\gaps\" & sBranch & " Open Order Report\" & sSequence & "\" & sISN & "\"

    If Sheets("117 DS").Range("A1").Value <> "" Then
        FileName = Format(Date, "yyyy-mm-dd") & " CUST DS OOR.xlsx"
        Saved = SaveWorkbook(Sheets("117 DS"), FilePath, FileName)
        If Saved = True Then
            Files = FilePath & FileName
            EmailBody = "Customer DS OOR attached."
        End If
    End If

    If Sheets("117 BO").Range("A1").Value <> "" Then
        FileName = Format(Date, "yyyy-mm-dd") & " CUST BO OOR.xlsx"
        Saved = SaveWorkbook(Sheets("117 BO"), FilePath, FileName)
        If Saved = True Then
            If TypeName(Files) = "String" Then
                Files = Array(Files, FilePath & FileName)
                EmailBody = "Customer DS & BO OOR attached."
            Else
                Files = FilePath & FileName
                EmailBody = "Customer BO OOR attached."
            End If
        End If
    End If

    Email Environ("username") & "@wesco.com", _
          Subject:="Customer Open Order Report", _
          Body:=EmailBody, _
          MailType:=SMTP, _
          Attachment:=Files
End Sub

Private Function SaveWorkbook(SourceSheet As Worksheet, FilePath As String, FileName As String) As Boolean
    Dim PrevDispAlert As Boolean

    PrevDispAlert = Application.DisplayAlerts
    Application.DisplayAlerts = False

    If Not FolderExists(FilePath) Then
        RecMkDir FilePath
    End If

    SourceSheet.Copy

    On Error GoTo FAILED_SAVE
    ActiveWorkbook.SaveAs FilePath & FileName, xlOpenXMLWorkbook
    On Error GoTo 0

    ActiveWorkbook.Close
    PrevDispAlert = PrevDispAlert
    SaveWorkbook = True
    Exit Function

FAILED_SAVE:
    MsgBox "An error occured while saving."
    ThisWorkbook.Activate
    PrevDispAlert = PrevDispAlert
    SaveWorkbook = False
End Function
