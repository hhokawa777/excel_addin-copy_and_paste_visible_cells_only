Attribute VB_Name = "RibbonControl"
' =======================================================================
' ribbon related
'   by hhokawa777@gmail.com
' =======================================================================
Option Explicit

Sub SelectVisible(control As IRibbonControl)
    util_SelectVisible
End Sub

Sub CopyVisible(control As IRibbonControl)
    util_CopyVisible
End Sub


Sub ClearVisible(control As IRibbonControl)
    Dim rc As Integer
    rc = MsgBox("Content and format of visible cells on your selected area will be cleared." & vbNewLine & "You cannot undo this operation." & vbNewLine & vbNewLine & "Are you sure you want to proceed ?", vbYesNo + vbQuestion, "Confirm")
    If rc = vbYes Then
        util_ClearVisible
    End If
End Sub

Sub ClearContentsVisible(control As IRibbonControl)
    Dim rc As Integer
    rc = MsgBox("Content of visible cells on your selected area will be cleared." & vbNewLine & "You cannot undo this operation." & vbNewLine & vbNewLine & "Are you sure you want to proceed ?", vbYesNo + vbQuestion, "Šm”F")
    If rc = vbYes Then
       util_ClearContentsVisible
    End If
End Sub

Sub CopyPasteAllToVisible(control As IRibbonControl)
    util_CopyPasteToVisible xlPasteAll
End Sub

Sub CopyPasteValueToVisible(control As IRibbonControl)
    util_CopyPasteToVisible xlPasteValues
End Sub

Sub CopyPasteFormulaToVisible(control As IRibbonControl)
    util_CopyPasteToVisible xlPasteFormulas
End Sub

Sub CopyPasteFormulaAsItIsToVisible(control As IRibbonControl)
    util_CopyPasteToVisible 9994123
End Sub


Sub onAction(control As IRibbonControl)

Dim id As String

id = control.id     ' button id

Select Case True    'update here when control id or sheet name changed

'-----------------------------------------

'Case (id = "xxxxx")
'    dummy

'Case (id = "xxxxx")
'    dummy
    
Case Else

End Select

End Sub
