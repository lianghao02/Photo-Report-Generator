Attribute VB_Name = "Module_Import"
Option Explicit

' ==================================================================================
'  第一部分：照片匯入與整理
' ==================================================================================

Sub ImportProcessedPhotos()
    Dim dialog As FileDialog
    Dim fullPath As String, fileName As String, finalPath As String
    Dim fso As Object
    Dim wsData As Worksheet, wsSet As Worksheet
    Dim targetRow As Long, currentRow As Long, lastRow As Long
    Dim i As Integer
    Dim startSerial As Long, inputSerial As String
    Dim startRowSetting As Long
    
    On Error Resume Next
    Set wsSet = ThisWorkbook.Sheets("設定")
    If wsSet Is Nothing Then MsgBox "[錯誤] 找不到 [設定] 工作表！", vbCritical: Exit Sub
    On Error GoTo 0
    
    Dim opMode As String, autoClean As String
    opMode = wsSet.Range("B2").Value
    autoClean = wsSet.Range("B3").Value
    
    If IsNumeric(wsSet.Range("B11").Value) And wsSet.Range("B11").Value > 1 Then
        startRowSetting = CLng(wsSet.Range("B11").Value)
    Else
        startRowSetting = 6
    End If
    
    Set wsData = ActiveSheet
    lastRow = wsData.Cells(wsData.Rows.Count, 2).End(xlUp).Row
    If lastRow < (startRowSetting - 1) Then targetRow = startRowSetting Else targetRow = lastRow + 1
    
    If wsData.Cells(targetRow, 2).Value <> "" Then
        If MsgBox("目前位置已有資料，是否從第 " & targetRow & " 列開始新增？", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim smallFolderPath As String, smallFilePath As String
    
    If opMode = "複製備份" Then
        smallFolderPath = ThisWorkbook.Path & "\Print_Images\"
        If Not fso.FolderExists(smallFolderPath) Then fso.CreateFolder (smallFolderPath)
        If autoClean = "是" Then
            On Error Resume Next: fso.DeleteFile smallFolderPath & "*.*", True: On Error GoTo 0
        End If
    End If

    Set dialog = Application.FileDialog(msoFileDialogFilePicker)
    With dialog
        .Title = "請選取照片"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "圖片檔案", "*.jpg; *.jpeg; *.png; *.bmp"
        
        If .Show = -1 Then
            Application.ScreenUpdating = False
            If IsNumeric(wsData.Cells(targetRow - 1, 1).Value) And wsData.Cells(targetRow - 1, 1).Value <> "" Then
                startSerial = wsData.Cells(targetRow - 1, 1).Value + 1
            Else
                startSerial = 1
            End If
            
            Application.ScreenUpdating = True
            inputSerial = InputBox("匯入 " & .SelectedItems.Count & " 張。起始編號為：", "編號", startSerial)
            Application.ScreenUpdating = False
            If inputSerial = "" Or Not IsNumeric(inputSerial) Then startSerial = 0 Else startSerial = CLng(inputSerial)
            
            For i = 1 To .SelectedItems.Count
                currentRow = targetRow + (i - 1)
                Application.StatusBar = "處理中..." & i & "/" & .SelectedItems.Count
                
                If startSerial > 0 Then wsData.Cells(currentRow, 1).Value = startSerial + (i - 1)
                
                fullPath = .SelectedItems(i)
                fileName = fso.GetFileName(fullPath)
                
                If opMode = "複製備份" Then
                    smallFilePath = smallFolderPath & fileName
                    On Error Resume Next: fso.CopyFile fullPath, smallFilePath, True: On Error GoTo 0
                    finalPath = smallFilePath
                Else
                    finalPath = fullPath
                End If
                
                finalPath = Replace(finalPath, "\", "\\")
                wsData.Cells(currentRow, 2).Value = finalPath
                wsData.Cells(currentRow, 9).Value = fileName
                
                If currentRow > startRowSetting Then
                    If wsData.Cells(currentRow - 1, 3).Value <> "" Then
                        wsData.Range("C" & currentRow & ":H" & currentRow).Value = _
                        wsData.Range("C" & (currentRow - 1) & ":H" & (currentRow - 1)).Value
                    End If
                End If
            Next i
            Call FormatDateTimeColumns
        End If
    End With
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Set dialog = Nothing
    Set fso = Nothing
End Sub

Sub FormatDateTimeColumns()
    Dim ws As Object, wsSet As Object
    Dim lastRowD As Long, lastRowE As Long, i As Long
    Dim startRowSetting As Long
    
    Set ws = ActiveSheet
    On Error Resume Next: Set wsSet = ThisWorkbook.Sheets("設定"): On Error GoTo 0
    
    If wsSet Is Nothing Then
        startRowSetting = 6
    Else
        If IsNumeric(wsSet.Range("B11").Value) And wsSet.Range("B11").Value > 1 Then
            startRowSetting = wsSet.Range("B11").Value
        Else
            startRowSetting = 6
        End If
    End If
    
    lastRowD = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    lastRowE = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    
    If lastRowD >= startRowSetting Then ws.Range("D" & startRowSetting & ":D" & lastRowD).NumberFormat = "@"
    If lastRowE >= startRowSetting Then ws.Range("E" & startRowSetting & ":E" & lastRowE).NumberFormat = "@"
    
    Dim cellValue As String, parts As Variant
    For i = startRowSetting To lastRowD
        If ws.Cells(i, "D").Value <> "" Then
            cellValue = CStr(ws.Cells(i, "D").Value)
            If InStr(cellValue, "/") > 0 Then
                parts = Split(cellValue, "/")
                If UBound(parts) = 2 And IsNumeric(parts(0)) Then
                    If Len(parts(0)) = 4 Then parts(0) = CStr(CInt(parts(0)) - 1911)
                    ws.Cells(i, "D").Value = parts(0) & "年" & CInt(parts(1)) & "月" & CInt(parts(2)) & "日"
                End If
            ElseIf IsNumeric(cellValue) Then
                If Len(cellValue) = 7 Then ws.Cells(i, "D").Value = Left(cellValue, 3) & "年" & CInt(Mid(cellValue, 4, 2)) & "月" & CInt(Right(cellValue, 2)) & "日"
                If Len(cellValue) = 8 Then ws.Cells(i, "D").Value = (CInt(Left(cellValue, 4)) - 1911) & "年" & CInt(Mid(cellValue, 5, 2)) & "月" & CInt(Right(cellValue, 2)) & "日"
            End If
        End If
    Next i
    
    For i = startRowSetting To lastRowE
        If ws.Cells(i, "E").Value <> "" Then
            cellValue = CStr(ws.Cells(i, "E").Value)
            If InStr(cellValue, ":") > 0 Then
                parts = Split(cellValue, ":")
                If UBound(parts) >= 1 Then ws.Cells(i, "E").Value = CInt(parts(0)) & "時" & CInt(parts(1)) & "分" & IIf(UBound(parts) = 2, CInt(parts(2)) & "秒", "")
            ElseIf IsNumeric(cellValue) Then
                If Len(cellValue) = 6 Then ws.Cells(i, "E").Value = CInt(Left(cellValue, 2)) & "時" & CInt(Mid(cellValue, 3, 2)) & "分" & CInt(Right(cellValue, 2)) & "秒"
                If Len(cellValue) = 4 Then ws.Cells(i, "E").Value = CInt(Left(cellValue, 2)) & "時" & CInt(Right(cellValue, 2)) & "分"
            End If
        End If
    Next i
End Sub
