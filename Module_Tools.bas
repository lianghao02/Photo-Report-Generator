Attribute VB_Name = "Module_Tools"
Option Explicit

' ==================================================================================
'  第三部分：輔助功能
' ==================================================================================

Function GetSelectedRowRange() As Range
    Dim selObj As Object
    Set selObj = Selection
    If TypeName(selObj) = "Range" Then
        Set GetSelectedRowRange = selObj.EntireRow
        Exit Function
    End If
    On Error Resume Next
    If TypeName(selObj) = "Picture" Or TypeName(selObj) = "DrawingObject" Or TypeName(selObj) = "Shape" Then
        Set GetSelectedRowRange = selObj.TopLeftCell.EntireRow
        Exit Function
    End If
    On Error GoTo 0
    Set GetSelectedRowRange = Nothing
End Function

Sub MoveRowUp()
    Dim rng As Range, startRowSetting As Long, selStart As Long, selEnd As Long
    On Error Resume Next: startRowSetting = ThisWorkbook.Sheets("設定").Range("B11").Value: On Error GoTo 0
    If startRowSetting < 2 Then startRowSetting = 6
    
    Set rng = GetSelectedRowRange()
    If rng Is Nothing Then MsgBox "請先點選範圍", vbExclamation: Exit Sub
    If rng.Areas.Count > 1 Then MsgBox "請選擇連續範圍", vbExclamation: Exit Sub
    
    selStart = rng.Row: selEnd = selStart + rng.Rows.Count - 1
    If selStart <= startRowSetting Then Exit Sub
    
    Application.ScreenUpdating = False
    rng.Cut
    Rows(selStart - 1).Insert Shift:=xlDown
    Rows(selStart - 1 & ":" & selEnd - 1).Select
    Call ResetSerialNumbers(startRowSetting)
    Application.ScreenUpdating = True
End Sub

Sub MoveRowDown()
    Dim rng As Range, startRowSetting As Long, selStart As Long, selEnd As Long
    On Error Resume Next: startRowSetting = ThisWorkbook.Sheets("設定").Range("B11").Value: On Error GoTo 0
    If startRowSetting < 2 Then startRowSetting = 6
    
    Set rng = GetSelectedRowRange()
    If rng Is Nothing Then MsgBox "請先點選範圍", vbExclamation: Exit Sub
    selStart = rng.Row: selEnd = selStart + rng.Rows.Count - 1
    
    If selStart < startRowSetting Then Exit Sub
    If ActiveSheet.Cells(selEnd + 1, 2).Value = "" Then Exit Sub
    
    Application.ScreenUpdating = False
    rng.Cut
    Rows(selEnd + 2).Insert Shift:=xlDown
    Rows(selStart + 1 & ":" & selEnd + 1).Select
    Call ResetSerialNumbers(startRowSetting)
    Application.ScreenUpdating = True
End Sub

Sub ClearAllData()
    Dim startRow As Long
    On Error Resume Next: startRow = ThisWorkbook.Sheets("設定").Range("B11").Value: On Error GoTo 0
    If startRow < 2 Then startRow = 6
    If MsgBox("確定要清除所有資料嗎？", vbYesNo + vbCritical) = vbYes Then
        Rows(startRow & ":" & Rows.Count).ClearContents
        Range("A" & startRow).Select
    End If
End Sub

Sub ResetSerialNumbers(startRow As Long)
    Dim ws As Object, lastRow As Long, i As Long
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastRow < startRow Then Exit Sub
    For i = startRow To lastRow
        ws.Cells(i, 1).Value = i - (startRow - 1)
    Next i
End Sub
