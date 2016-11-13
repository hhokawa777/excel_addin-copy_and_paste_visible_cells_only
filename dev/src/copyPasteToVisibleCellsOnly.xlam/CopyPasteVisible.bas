Attribute VB_Name = "CopyPasteVisible"
' =======================================================================
' Copy of paste area
'   by hhokawa777@gmail.com
' =======================================================================
Option Explicit

'------------------------------------------------------------
' Module Variables
'------------------------------------------------------------
Dim myBar As cProgress

Sub util_SelectVisible(Optional dummy)
    If ActiveWindow.selection.CountLarge > 1 Then
        ActiveWindow.selection.SpecialCells(xlCellTypeVisible).Select
    End If
End Sub

Sub util_CopyVisible(Optional dummy)
    util_SelectVisible
    ActiveWindow.selection.copy
End Sub


Sub util_ClearVisible(Optional dummy)
    util_SelectVisible
    ActiveWindow.selection.Clear
End Sub

Sub util_ClearContentsVisible(Optional dummy)
    util_SelectVisible
    ActiveWindow.selection.ClearContents
End Sub

Sub util_CopyPasteToVisible(pasteType As Long, Optional seekHorizonallyFirst As Boolean, Optional fromRange As Range, Optional toRange As Range, Optional numOfRowRepeat = 1, Optional numOfColRepeat = 1)

    '------------------------------------------------------------
    ' Variables
    '------------------------------------------------------------

    Dim tempMsg As String, rc As Long
    Dim warningFromCellCount As Currency, fromCellCount As Currency

    '------------------------------------------------------------
    ' Ask User for Destination cell
    '------------------------------------------------------------
    If fromRange Is Nothing Then
        Set fromRange = ActiveWindow.selection
    End If
    If toRange Is Nothing Then
        If pasteType = xlPasteAll Then
            warningFromCellCount = 10000
            If fromRange.CountLarge > 1 Then
                fromCellCount = fromRange.SpecialCells(xlCellTypeVisible).CountLarge
            Else
                fromCellCount = 1
            End If
            
            If fromCellCount > warningFromCellCount Then

                rc = MsgBox(warningFromCellCount & " or more cells (" & fromCellCount & " cells ) are selected for copying." & vbNewLine & _
                    "Paste (all) operation you chose will take long time." & vbNewLine & _
                    "Any clipboard operation won't work for other applications during this operation." & vbNewLine & _
                    vbNewLine & _
                    "Recommend to use paste (value) or (formula) instead." & _
                    vbNewLine & _
                    vbNewLine & _
                    "Are you sure you want to proceed ?" _
                    , vbYesNo + vbQuestion, "Confirm")


                If rc = vbNo Then
                    GoTo Finally
                End If
            End If

            tempMsg = "Copy & paste visible cells (all), where to paste ?"
        ElseIf pasteType = xlPasteFormulas Then
            tempMsg = "Copy & paste visible cells (formula), where to paste ?"
        ElseIf pasteType = 9994123 Then
            tempMsg = "Copy & paste visible cells (formula as is: formula's reference not changed), where to paste ?"
        ElseIf pasteType = xlPasteValues Then
            tempMsg = "Copy & paste visible cells (value), where to paste ?"
        Else
            MsgBox "wrong paste type specified in macro code."
            GoTo Finally
        End If
        
        frmInputToRange.rowRepeat = numOfRowRepeat
        frmInputToRange.colRepeat = numOfColRepeat
        frmInputToRange.Caption = tempMsg
        frmInputToRange.Show

        
        If frmInputToRange.toRange <> "" Then
            Set toRange = Application.Range(frmInputToRange.toRange.Text)
            numOfRowRepeat = frmInputToRange.rowRepeat.Text
            numOfColRepeat = frmInputToRange.colRepeat.Text
        End If
        Unload frmInputToRange
    End If
    
    If Not toRange Is Nothing Then
        Call util_CopyPasteToVisible_core(pasteType, seekHorizonallyFirst, fromRange, toRange, numOfRowRepeat, numOfColRepeat)
    End If
    
    '------------------------------------------------------------
    ' Finalization
    '------------------------------------------------------------
Finally:
    Exit Sub
    '------------------------------------------------------------
    ' Error handling
    '------------------------------------------------------------
ErrorHandler:
    MsgBox "Number : " & Err.Number & vbCr & _
         "Source : " & Err.Source & vbCr & _
         "Description : " & Err.Description, _
         vbCritical, _
         "SYSTEM ERROR"
    Resume Finally
End Sub
    
Sub util_CopyPasteToVisible_core(pasteType As Long, seekHorizonallyFirst As Boolean, fromRange As Range, toRange As Range, Optional numOfRowRepeat = 1, Optional numOfColRepeat = 1)
    'reference http://www.mrexcel.com/forum/excel-questions/85288-paste-visible-cells-only-3.html
    '------------------------------------------------------------
    ' Variables
    '------------------------------------------------------------
    Dim fromCell As Range, tempCell As Range
    Dim fromCellAddrs As Object, fromCellAddr As Variant
    Dim fromSheet As Worksheet, toSheet As Worksheet
    Dim rowZeroPadding As String, colZeroPadding As String, lenOfRowZeroPadding As Long, lenOfColZeroPadding As Long
    Dim originalCalcState As Variant, StartTime As Date, elaspedTime As Date, cnt As Currency, cntSum As Currency
    Dim FirstCell As Range, lastCell As Range, newToRange As Range
    Dim colRepeat As Long, rowRepeat As Long, rowIndex As Long, colIndex As Long
    
'    On Error GoTo ErrorHandler
    originalCalcState = Null
    Application.ScreenUpdating = False
    originalCalcState = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    '------------------------------------------------------------
    ' Initialization
    '------------------------------------------------------------
    StartTime = Now
    Application.StatusBar = "Initializing ..."
    Set myBar = New cProgress
    myBar.Init
    Set fromSheet = fromRange.Worksheet
    Set toSheet = toRange.Worksheet
   
    '--- set only visible cells -----
    If fromRange.CountLarge > 1 Then
        Set fromRange = fromRange.SpecialCells(xlCellTypeVisible)
    End If
    
    
    '--- Sort destination cells order ---------------------------
    Set fromCellAddrs = CreateObject("System.Collections.ArrayList")
    lenOfRowZeroPadding = Len(Rows.Count & "")
    lenOfColZeroPadding = Len(Columns.Count & "")
    rowZeroPadding = String(lenOfRowZeroPadding, "0")
    colZeroPadding = String(lenOfColZeroPadding, "0")
    If seekHorizonallyFirst Then
        For Each fromCell In fromRange
            fromCellAddrs.Add Format(fromCell.Row, rowZeroPadding) & "-" & Format(fromCell.Column, colZeroPadding)
        Next
    Else
        For Each fromCell In fromRange
            fromCellAddrs.Add Format(fromCell.Column, colZeroPadding) & "-" & Format(fromCell.Row, rowZeroPadding)
        Next
    End If
    cntSum = fromCellAddrs.Count
    fromCellAddrs.Sort

    '------------------------------------------------------------
    ' Copy and Paste
    '------------------------------------------------------------
    '--- initialization -----------------------------------------
    With myBar
        .View = 3
        .Min = 0
        .Max = cntSum * numOfRowRepeat * numOfColRepeat
        .Msg = ", Copying ..."
    End With
    myBar.Start
    cnt = 0
    '--- copy and paste repeatedly-------------------------------
    Set FirstCell = toRange.Range("A1")
    For colRepeat = 1 To numOfColRepeat
        rowIndex = FirstCell.Row
        If colRepeat = 1 Then
            colIndex = FirstCell.Column
        Else
            Set tempCell = util_findNextVisibleCellHorizonally(lastCell)
            If tempCell Is Nothing Then
                Exit For
            End If
            colIndex = tempCell.Column
        End If
    
        For rowRepeat = 1 To numOfRowRepeat
            If rowRepeat = 1 Then
                rowIndex = FirstCell.Row
            Else
                Set tempCell = util_findNextVisibleCellVertically(lastCell)
                If tempCell Is Nothing Then
                    Exit For
                End If
                rowIndex = tempCell.Row
            End If
            
            Set newToRange = toSheet.Cells(rowIndex, colIndex)
            Set lastCell = util_CopyPasteToVisible_core2(pasteType, seekHorizonallyFirst, fromRange, newToRange, fromCellAddrs, cnt)
    
        Next
    
    Next
    
    '--- finalization -------------------------------------------
    Set myBar = Nothing
    elaspedTime = Now - StartTime
    Debug.Print cnt & " copied." & "Elapsed: " & elaspedTime
    
    toSheet.Activate
    toSheet.Range(toRange.Range("A1"), lastCell).Select
    
    '------------------------------------------------------------
    ' Finalization
    '------------------------------------------------------------
Finally:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    If Not IsNull(originalCalcState) Then
        Application.Calculation = originalCalcState
    End If
    If Not fromRange Is Nothing Then
        fromRange.copy
    End If
    Exit Sub
    '------------------------------------------------------------
    ' Error handling
    '------------------------------------------------------------
ErrorHandler:
    MsgBox "Number : " & Err.Number & vbCr & _
         "Source : " & Err.Source & vbCr & _
         "Description : " & Err.Description, _
         vbCritical, _
         "SYSTEM ERROR"
    Resume Finally
End Sub

Function util_findNextVisibleCellHorizonally(aCell As Range) As Range
    If aCell.Column < Columns.Count Then                           ' if not over max column
        Set aCell = aCell.Offset(0, 1)                              ' move to next column
        Set util_findNextVisibleCellHorizonally = util_findVisibleCellHorizonally(aCell)
    Else
        Set util_findNextVisibleCellHorizonally = Nothing
    End If
End Function
Function util_findVisibleCellHorizonally(aCell As Range) As Range
    Dim flgFound As Boolean
                
    flgFound = False
    Do While flgFound = False
        If aCell.EntireColumn.ColumnWidth > 0 Then                 ' if not hidden column
            flgFound = True
        Else
            If aCell.Column < Columns.Count Then                   ' if hidden column
                Set aCell = aCell.Offset(0, 1)                     ' move to next column
            Else
                Exit Do
            End If
        End If
    Loop
    
    If flgFound = True Then
        Set util_findVisibleCellHorizonally = aCell
    Else
        Set util_findVisibleCellHorizonally = Nothing
    End If
    
End Function

Function util_findNextVisibleCellVertically(aCell As Range) As Range
    If aCell.Row < Rows.Count Then                                  ' if not over max row
        Set aCell = aCell.Offset(1, 0)                              ' move to next row
        Set util_findNextVisibleCellVertically = util_findVisibleCellVertically(aCell)
    Else
        Set util_findNextVisibleCellVertically = Nothing
    End If
End Function

Function util_findVisibleCellVertically(aCell As Range) As Range
    Dim flgFound As Boolean
    flgFound = False
    Do While flgFound = False
        If aCell.EntireRow.RowHeight > 0 Then                       ' if not hidden row
            flgFound = True
        Else                                                        ' if hidden row
            If aCell.Row < Rows.Count Then
                Set aCell = aCell.Offset(1, 0)                      ' move to next row
            Else
                Exit Do
            End If
        End If
    Loop
        
    If flgFound = True Then
        Set util_findVisibleCellVertically = aCell
    Else
        Set util_findVisibleCellVertically = Nothing
    End If
    
End Function

    
Function util_CopyPasteToVisible_core2(pasteType As Long, seekHorizonallyFirst As Boolean, fromRange As Range, toRange As Range, fromCellAddrs As Object, ByRef cnt As Currency) As Range
    '------------------------------------------------------------
    ' Variables
    '------------------------------------------------------------
    Dim fromCell As Range, toCell As Range, lastCell As Range
    Dim fromCurrRow As Long, fromCurrCol As Long, fromPrevRow As Long, fromPrevCol As Long, fromStartRow As Long, fromStartCol As Long
    Dim toStartRow As Long, toStartCol As Long, fromCellAddr As Variant
    Dim fromSheet As Worksheet, toSheet As Worksheet
    Dim cntToUpdateBar As Long
    Dim tempRange As Range, flgFound As Boolean, tempMsg As String, rc As Long, flgMaxRowReached As Boolean, flgMaxColReached As Boolean
    Dim lenOfRowZeroPadding As Long, lenOfColZeroPadding As Long
    
    '------------------------------------------------------------
    ' Initialization
    '------------------------------------------------------------
    lenOfRowZeroPadding = Len(Rows.Count & "")
    lenOfColZeroPadding = Len(Columns.Count & "")
    Set fromSheet = fromRange.Worksheet
    Set toSheet = toRange.Worksheet
    '------------------------------------------------------------
    ' Copy and Paste
    '------------------------------------------------------------
    '--- initialization -----------------------------------------
    fromPrevCol = -1
    fromPrevRow = -1
    Set fromCell = fromRange.Range("A1")
    fromStartRow = fromCell.Row
    fromStartCol = fromCell.Column
    Set toCell = toRange.Range("A1")
    toStartRow = toCell.Row
    toStartCol = toCell.Column
    flgMaxRowReached = False
    flgMaxColReached = False
    If pasteType = xlPasteAll Then
        cntToUpdateBar = 1000
    Else
        cntToUpdateBar = 10000
    End If
    
    '--- copy and paste for each cells --------------------------
    For Each fromCellAddr In fromCellAddrs
        '--- in-loop initialization -----------------------------
        If seekHorizonallyFirst Then
            Set fromCell = fromSheet.Cells(CLng(Left(fromCellAddr, lenOfRowZeroPadding)), CLng(Right(fromCellAddr, lenOfColZeroPadding)))
        Else
            Set fromCell = fromSheet.Cells(CLng(Right(fromCellAddr, lenOfRowZeroPadding)), CLng(Left(fromCellAddr, lenOfColZeroPadding)))
        End If
        
        '--- in-loop main logic ---------------------------------
        fromCurrRow = fromCell.Row
        fromCurrCol = fromCell.Column
        If seekHorizonallyFirst Then                                            ' if specified to seek horizonal direction
            '--- if new row ------------------                                  ' TODO: replace logic with find visible cell function above
            If (fromPrevRow <> -1) And (fromPrevRow <> fromCurrRow) Then        ' if new row
                If toCell.Row < Rows.Count Then                                 ' if not over max row
                    Set toCell = toSheet.Cells(toCell.Row + 1, toStartCol)      ' move to next row and reset column position
                Else
                    flgMaxRowReached = True
                    Exit For
                End If
                flgFound = False
                Do While flgFound = False
                    If toCell.EntireRow.RowHeight > 0 Then                      ' if not hidden row
                        flgFound = True
                    Else                                                        ' if hidden row
                        If toCell.Row < Rows.Count Then
                            Set toCell = toCell.Offset(1, 0)                    ' move to next row
                        Else
                            Exit For
                        End If
                    End If
                Loop
                flgMaxColReached = False
            End If
            '--- seek copy location to horizonal direction ------------------
            If flgMaxColReached = False Then
                flgFound = False
                Do While flgFound = False
                    If toCell.EntireColumn.ColumnWidth > 0 Then                     ' if not hidden row, do copy and paste
                        If pasteType = xlPasteAll Then
                            fromCell.copy toCell
                        ElseIf pasteType = xlPasteFormulas Then
                            toCell.FormulaR1C1 = fromCell.FormulaR1C1
                        ElseIf pasteType = 9994123 Then
                            toCell.formula = fromCell.formula
                        ElseIf pasteType = xlPasteValues Then
                            toCell.Value = fromCell.Value
                        Else
                            MsgBox "wrong paste type specified in macro code."
                            GoTo Finally
                        End If
                        Set lastCell = toCell
                        flgFound = True
                    End If
                    If toCell.Column < Columns.Count Then                           ' if not max column
                        Set toCell = toCell.Offset(0, 1)                            ' move to next column
                    Else
                        flgMaxColReached = True
                        Exit Do
                    End If
                Loop
            End If
        Else
            '--- if new column ------------------
            If (fromPrevCol <> -1) And (fromPrevCol <> fromCurrCol) Then        ' if new column
                If toCell.Column < Columns.Count Then                           ' if not over max column
                    Set toCell = toSheet.Cells(toStartRow, toCell.Column + 1)   ' move to next column and reset row position
                Else
                    flgMaxColReached = True
                    Exit For
                End If
                flgFound = False
                Do While flgFound = False
                    If toCell.EntireColumn.ColumnWidth > 0 Then                 ' if not hidden column
                        flgFound = True
                    Else
                        If toCell.Column < Columns.Count Then                   ' if hidden column
                            Set toCell = toCell.Offset(0, 1)                    ' move to next column
                        Else
                            Exit For
                        End If
                    End If
                Loop
                flgMaxRowReached = False
            End If
            '--- seek copy location to vertical direction ------------------
            If flgMaxRowReached = False Then
                flgFound = False
                Do While flgFound = False
                    If toCell.EntireRow.RowHeight > 0 Then                          ' if not hidden row, do copy and paste
                        If pasteType = xlPasteAll Then
                            fromCell.copy toCell
                        ElseIf pasteType = xlPasteFormulas Then
                            toCell.FormulaR1C1 = fromCell.FormulaR1C1
                        ElseIf pasteType = 9994123 Then
                            toCell.formula = fromCell.formula
                        ElseIf pasteType = xlPasteValues Then
                            toCell.Value = fromCell.Value
                        Else
                            MsgBox "wrong paste type specified in macro code."
                            GoTo Finally
                        End If
                        Set lastCell = toCell
                        flgFound = True
                    End If
                    If toCell.Row < Rows.Count Then                                 ' if not max column
                        Set toCell = toCell.Offset(1, 0)                            ' move to next row
                    Else
                        flgMaxRowReached = True
                        Exit Do
                    End If
                Loop
            End If
        End If
        
        '--- in-loop finalization -------------------------------
        fromPrevRow = fromCurrRow
        fromPrevCol = fromCurrCol
        cnt = cnt + 1
        If (cnt Mod cntToUpdateBar = 0) Then
            myBar.Value = myBar.Value + cntToUpdateBar
'            DoEvents
        End If
    Next
    Set util_CopyPasteToVisible_core2 = lastCell
    
    '------------------------------------------------------------
    ' Finalization
    '------------------------------------------------------------
Finally:
    Exit Function
    '------------------------------------------------------------
    ' Error handling
    '------------------------------------------------------------
ErrorHandler:
    MsgBox "Number : " & Err.Number & vbCr & _
         "Source : " & Err.Source & vbCr & _
         "Description : " & Err.Description, _
         vbCritical, _
         "SYSTEM ERROR"
    Set util_CopyPasteToVisible_core2 = Nothing
    Resume Finally
End Function

