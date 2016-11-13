VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInputToRange 
   Caption         =   "Where to paste ?"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7965
   OleObjectBlob   =   "frmInputToRange.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmInputToRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =======================================================================
' dialog
'   by hhokawa777@gmail.com
' =======================================================================


Private Sub btnCancel_Click()
    toRange.Text = ""
        Me.Hide
End Sub

Private Sub btnOK_Click()
    Me.Hide
End Sub


Private Sub SpinColRepeat_SpinDown()
    If colRepeat.Text > 1 Then
        colRepeat.Text = colRepeat.Text - 1
    End If
End Sub

Private Sub SpinColRepeat_SpinUp()
        colRepeat.Text = colRepeat.Text + 1
End Sub


Private Sub SpinRowRepeat_SpinDown()
    If rowRepeat.Text > 1 Then
        rowRepeat.Text = rowRepeat.Text - 1
    End If
End Sub

Private Sub SpinRowRepeat_SpinUp()
    rowRepeat.Text = rowRepeat.Text + 1
End Sub


Private Sub UserForm_Initialize()
    toRange.Text = ""
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    toRange.Text = ""
End Sub
