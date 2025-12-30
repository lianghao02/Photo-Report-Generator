Attribute VB_Name = "Module_Generate"
Option Explicit

' ==================================================================================
'  第二部分：Word 輸出指揮官
' ==================================================================================

Sub GeneratePDF_Landscape()
    Dim wsSet As Worksheet
    Dim wVal As Variant, hVal As Variant, tVal As Variant
    
    On Error Resume Next
    Set wsSet = ThisWorkbook.Sheets("設定")
    If wsSet Is Nothing Then MsgBox "[嚴重錯誤] 找不到「設定」工作表！", vbCritical: Exit Sub
    On Error GoTo 0
    
    tVal = wsSet.Range("B9").Value
    wVal = wsSet.Range("B5").Value ' 一般照片_寬
    hVal = wsSet.Range("B6").Value ' 一般照片_高
    
    Call CoreGenerateDocs(CStr(tVal), CSng(Val(wVal)), CSng(Val(hVal)), "一般")
End Sub

Sub GeneratePDF_Mobile()
    Dim wsSet As Worksheet
    Dim wVal As Variant, hVal As Variant, tVal As Variant
    
    On Error Resume Next
    Set wsSet = ThisWorkbook.Sheets("設定")
    If wsSet Is Nothing Then MsgBox "[嚴重錯誤] 找不到「設定」工作表！", vbCritical: Exit Sub
    On Error GoTo 0
    
    tVal = wsSet.Range("B10").Value
    wVal = wsSet.Range("B7").Value  ' 手機截圖_寬
    hVal = wsSet.Range("B8").Value  ' 手機截圖_高
    
    Call CoreGenerateDocs(CStr(tVal), CSng(Val(wVal)), CSng(Val(hVal)), "手機")
End Sub

Private Sub CoreGenerateDocs(templateName As String, maxCmW As Single, maxCmH As Single, typeSuffix As String)
    ' 防呆檢查
    If maxCmW <= 0 Or maxCmH <= 0 Then
        MsgBox "[錯誤] 讀取到的尺寸為 0 或無效！" & vbCrLf & "請檢查 Excel 設定頁 B5~B8 數值。", vbCritical
        Exit Sub
    End If

    Dim wdApp As Object, wdDoc As Object, newDoc As Object
    Dim templatePath As String, saveWordPath As String, savePdfPath As String, excelPath As String
    Dim shp As Object
    Dim maxPointW As Single, maxPointH As Single
    Dim ratioW As Single, ratioH As Single, finalRatio As Single
    Dim wsSet As Worksheet, ws As Worksheet
    Dim lastRow As Long, lastCol As Long, startRow As Long
    
    Set wsSet = ThisWorkbook.Sheets("設定")
    If IsNumeric(wsSet.Range("B11").Value) Then startRow = wsSet.Range("B11").Value Else startRow = 6
    
    ' 模板檔名智慧補全
    Dim lowerTmp As String
    lowerTmp = LCase(templateName)
    If Right(lowerTmp, 5) <> ".docx" And Right(lowerTmp, 4) <> ".doc" Then
        templateName = templateName & ".docx"
    End If
    
    excelPath = ThisWorkbook.FullName
    templatePath = ThisWorkbook.Path & "\" & templateName
    If Dir(templatePath) = "" Then MsgBox "[錯誤] 找不到模板：" & templateName, vbCritical: Exit Sub
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row
    If lastRow < startRow Then lastRow = startRow
    lastCol = ws.Cells(5, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 9 Then lastCol = 9
    
    On Error Resume Next: ActiveWorkbook.Names("WordData").Delete: On Error GoTo 0
    ActiveWorkbook.Names.Add Name:="WordData", RefersTo:=ws.Range(ws.Cells(startRow - 1, 1), ws.Cells(lastRow, lastCol))
    ThisWorkbook.Save
    
    On Error Resume Next: Set wdApp = GetObject(, "Word.Application"): On Error GoTo 0
    If wdApp Is Nothing Then Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True
    wdApp.Activate
    
    Set wdDoc = wdApp.Documents.Open(templatePath)
    wdApp.DisplayAlerts = 0
    With wdDoc.MailMerge
        .MainDocumentType = 0
        .OpenDataSource Name:=excelPath, SQLStatement:="SELECT * FROM [WordData]", ReadOnly:=True, LinkToSource:=True
        .Destination = 0
        .SuppressBlankLines = True
        .Execute Pause:=False
    End With
    wdApp.DisplayAlerts = -1
    
    DoEvents
    Set newDoc = wdApp.ActiveDocument
    newDoc.Activate
    
    On Error Resume Next
    wdApp.ActiveWindow.View.Type = 3
    On Error GoTo 0
    
    newDoc.Fields.Update
    DoEvents
    newDoc.Fields.Unlink
    
    Dim retryCount As Integer
    retryCount = 0
    Do While newDoc.InlineShapes.Count = 0 And retryCount < 50
        DoEvents
        Application.Wait (Now + TimeValue("0:00:01") / 10)
        retryCount = retryCount + 1
    Loop
    
    maxPointW = maxCmW * 28.346
    maxPointH = maxCmH * 28.346
    
    Dim picCount As Integer
    picCount = 0
    
    For Each shp In newDoc.InlineShapes
        picCount = picCount + 1
        shp.LockAspectRatio = 0
        shp.ScaleWidth = 100
        shp.ScaleHeight = 100
        
        If shp.Width > 0 And shp.Height > 0 Then
            ratioW = maxPointW / shp.Width
            ratioH = maxPointH / shp.Height
            
            If ratioW < ratioH Then finalRatio = ratioW Else finalRatio = ratioH
            
            shp.Width = shp.Width * finalRatio
            shp.Height = shp.Height * finalRatio
            shp.LockAspectRatio = -1
            
            If shp.Width > (maxPointW + 10) Then
                shp.Width = maxPointW
            End If
        End If
    Next shp
    
    With newDoc.Content.Find
        .ClearFormatting: .Replacement.Text = ""
        .Text = "錯誤! 尚未指定檔名。": .Execute Replace:=2
        .Text = "Error! Filename not specified.": .Execute Replace:=2
    End With
    
    ' === 檔名生成與防呆邏輯 ===
    Dim baseName As String, counter As Integer
    Dim nameFormat As String
    
    nameFormat = wsSet.Range("B12").Value
    If nameFormat = "" Then nameFormat = "[案由]_[類型]_[日期]_[時間]"
    
    Dim varCase As String, varType As String, varDate As String, varTime As String
    
    varCase = Trim(ws.Cells(startRow, 3).Value)
    If varCase = "" Then varCase = "報告"
    
    Dim InvalidChars As String, charIdx As Integer
    InvalidChars = "\/:*?""<>|"
    For charIdx = 1 To Len(InvalidChars)
        varCase = Replace(varCase, Mid(InvalidChars, charIdx, 1), "_")
    Next charIdx
    varCase = Replace(varCase, vbCr, ""): varCase = Replace(varCase, vbLf, "")
    
    varType = typeSuffix
    varDate = CStr(Year(Now) - 1911) & Format(Now, "MMdd")
    varTime = Format(Now, "HHmm")
    
    baseName = nameFormat
    baseName = Replace(baseName, "[案由]", varCase)
    baseName = Replace(baseName, "[類型]", varType)
    baseName = Replace(baseName, "[日期]", varDate)
    baseName = Replace(baseName, "[時間]", varTime)
    
    If LCase(Right(baseName, 5)) = ".docx" Then baseName = Left(baseName, Len(baseName) - 5)
    If LCase(Right(baseName, 4)) = ".doc" Then baseName = Left(baseName, Len(baseName) - 4)
    If LCase(Right(baseName, 4)) = ".pdf" Then baseName = Left(baseName, Len(baseName) - 4)
    
    If Trim(baseName) = "" Then baseName = "報告"
    
    ' === [新增] 輸出資料夾邏輯 ===
    Dim outFolderName As String, saveRootPath As String
    
    ' 1. 讀取設定頁 B13
    outFolderName = Trim(wsSet.Range("B13").Value)
    
    ' 淨化資料夾名稱
    For charIdx = 1 To Len(InvalidChars)
        outFolderName = Replace(outFolderName, Mid(InvalidChars, charIdx, 1), "")
    Next charIdx
    
    ' 預設資料夾名稱
    If outFolderName = "" Then outFolderName = "產出報告"
    
    ' 2. 組合完整路徑
    saveRootPath = ThisWorkbook.Path & "\" & outFolderName
    
    ' 3. 檢查並建立資料夾
    If Dir(saveRootPath, vbDirectory) = "" Then
        MkDir saveRootPath
    End If
    
    ' 4. 檔名重複檢查與存檔 (加上資料夾路徑)
    counter = 0
    Do
        savePdfPath = saveRootPath & "\" & baseName & IIf(counter > 0, "(" & counter & ")", "") & ".pdf"
        If Dir(savePdfPath) = "" Then Exit Do
        counter = counter + 1
    Loop
    
    baseName = baseName & IIf(counter > 0, "(" & counter & ")", "")
    saveWordPath = saveRootPath & "\" & baseName & ".docx"
    savePdfPath = saveRootPath & "\" & baseName & ".pdf"
    
    On Error Resume Next
    newDoc.SaveAs2 fileName:=saveWordPath, FileFormat:=12
    newDoc.ExportAsFixedFormat OutputFileName:=savePdfPath, ExportFormat:=17
    
    If Err.Number <> 0 Then
        MsgBox "[錯誤] 存檔失敗！可能檔案已開啟。", vbCritical
        Err.Clear
    Else
        wdDoc.Close SaveChanges:=False
        MsgBox "報告製作完成！" & vbCrLf & "位置：" & outFolderName & vbCrLf & "檔名：" & baseName, vbInformation
    End If
    On Error GoTo 0
    
    Set newDoc = Nothing
    Set wdDoc = Nothing
    Set wdApp = Nothing
End Sub
