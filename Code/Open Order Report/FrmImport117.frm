VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmImport117 
   Caption         =   "117 Options"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4440
   OleObjectBlob   =   "FrmImport117.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmImport117"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
    Canceled = True
    Unload Me
End Sub

Private Sub btnOk_Click()
    'Check to see if either checkbox is clicked
    If chkBO + chkSO = 0 Then
        MsgBox "Please select the report criteria."
    ElseIf txtBranch.Text = "" And Len(txtBranch.Text) <> 4 Then
        MsgBox "Please enter a branch number."
    ElseIf Trim(txtSales.Text) = "" Then
        MsgBox "Please enter a sales number."
    Else
        Me.Hide
        sBranch = txtBranch.Text
        sISN = txtSales.Text

        If radInsideSales.Value = True Then
            sSequence = "ByInsideSalesperson"
        ElseIf radOutsideSales.Value = True Then
            sSequence = "ByOutsideSalesperson"
        End If

        On Error GoTo Import_Failed
        If chkBO = True Then
            If radInsideSales.Value = True Then
                Import117 BackOrders, ByInsideSalesperson, Now, One, txtSales.Text, txtBranch.Text, True, Sheets("117 BO").Range("A1")
            Else
                Import117 BackOrders, ByOutsideSalesperson, Now, One, txtSales.Text, txtBranch.Text, True, Sheets("117 BO").Range("A1")
            End If
        End If

        If chkSO = True Then
            If radInsideSales.Value = True Then
                Import117 DSOrders, ByInsideSalesperson, Now, One, txtSales.Text, txtBranch.Text, True, Sheets("117 DS").Range("A1")
            Else
                Import117 DSOrders, ByOutsideSalesperson, Now, One, txtSales.Text, txtBranch.Text, True, Sheets("117 DS").Range("A1")
            End If
        End If
        On Error GoTo 0

        Unload Me
    End If
    Exit Sub

Import_Failed:
    MsgBox Prompt:="Error " & Err.Number & " '" & Err.Description & "' occurred in " & Err.Source & ".", _
           Title:="Oops!"
    Resume Next
End Sub

Private Sub UserForm_Initialize()
    Canceled = False
    sBranch = ""
    sSequence = ""
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    lblBranch.Top = 62
    lblSalesNumber.Top = 62
    txtBranch.SetFocus
End Sub
