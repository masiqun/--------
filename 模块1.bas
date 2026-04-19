Attribute VB_Name = "模块1"

Sub terminplan()
    ' 禁用事件、屏幕更新和自动计算以提高性能
    With Application
        .EnableEvents = False
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
    
    On Error GoTo ErrorHandler
    
    ' 执行主要功能
    Call ProcessGanttChart
    
    Exit Sub
    
ErrorHandler:
    MsgBox "错误: " & Err.Description, vbCritical
    
Finally:
    ' 恢复应用程序设置
    With Application
        .DisplayAlerts = True
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub

Sub ProcessGanttChart()
    Dim ws As Worksheet
    Set ws = Worksheets("terminplan")
    
    ' 获取工作表数据范围
    Dim lastRow As Long, lastCol As Long
    Call GetWorksheetDimensions(ws, lastRow, lastCol)
    
    ' 检查是否有数据
    If lastRow = 0 Or lastCol = 0 Then
        MsgBox "工作表为空", vbExclamation
        Exit Sub
    End If
    
    ' 初始化甘特图列
    Dim startCol As Long
    startCol = GetStartColumn(ws, lastCol)
    
    ' 清理和设置表头
    Call SetupGanttHeader(ws, startCol, lastCol)
    
    ' 查找日期范围
    Dim tagBegin As Date, tagEnd As Date
    Dim tagBeginColumn As Long, tagEndColumn As Long
    Call FindDateRange(ws, tagBegin, tagEnd, tagBeginColumn, tagEndColumn)
    
    ' 计算日期差异
    Dim tagDiff As Long
    tagDiff = tagEnd - tagBegin + 1
    
    ' 设置单元格格式
    Call FormatGanttCells(ws, startCol, tagDiff)
    
    ' 填充日期数据
    Call FillDateData(ws, startCol, tagDiff, tagBegin)
    
    ' 合并相同值的单元格
    Call MergeSimilarCells(ws, startCol, tagDiff)
    
    ' 复制表头到底部
    Call CopyHeaderToBottom(ws, startCol, tagDiff, lastRow)
    
    ' 填充甘特图颜色
    Call FillGanttColors(ws, startCol, tagDiff, tagBegin, tagBeginColumn, tagEndColumn)
End Sub

Sub GetWorksheetDimensions(ws As Worksheet, ByRef lastRow As Long, ByRef lastCol As Long)
    Dim rngLast As Range
    
    ' 查找最后一行
    On Error Resume Next
    Set rngLast = ws.Cells.Find( _
        What:="*", _
        After:=ws.Range("A1"), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious _
    )
    If Not rngLast Is Nothing Then lastRow = rngLast.Row
    
    ' 查找最后一列
    Set rngLast = ws.Cells.Find( _
        What:="*", _
        After:=ws.Range("A1"), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByColumns, _
        SearchDirection:=xlPrevious _
    )
    If Not rngLast Is Nothing Then lastCol = rngLast.Column
    On Error GoTo 0
End Sub

Function GetStartColumn(ws As Worksheet, lastCol As Long) As Long
    Dim rngLast As Range
    
    If Application.WorksheetFunction.CountA(ws.Rows(1)) Then
        Set rngLast = ws.Rows(1).Cells.Find( _
            What:="*", _
            After:=ws.Range("A1"), _
            LookIn:=xlFormulas, _
            LookAt:=xlPart, _
            SearchOrder:=xlByColumns, _
            SearchDirection:=xlNext _
        )
        If Not rngLast Is Nothing Then
            GetStartColumn = rngLast.Column
        Else
            GetStartColumn = lastCol
        End If
    Else
        GetStartColumn = lastCol
    End If
End Function

Sub SetupGanttHeader(ws As Worksheet, startCol As Long, lastCol As Long)
    ' 清理现有数据
    ws.Range(ws.Columns(startCol), ws.Columns(lastCol)).Delete
    ws.Rows().Hidden = False
    
    ' 设置表头
    With ws
        .Cells(1, startCol) = "Jahr"
        .Cells(2, startCol) = "Monat"
        .Cells(3, startCol) = "Woche"
        .Cells(4, startCol) = "Tag"
        .Columns(startCol).EntireColumn.AutoFit
    End With
End Sub

Sub FindDateRange(ws As Worksheet, ByRef tagBegin As Date, ByRef tagEnd As Date, _
                  ByRef tagBeginColumn As Long, ByRef tagEndColumn As Long)
    Dim foundCell As Range
    
    ' 查找开始日期
    Set foundCell = ws.Cells.Find( _
        What:="开始日期", _
        After:=ws.Range("A1"), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext _
    )
    
    If Not foundCell Is Nothing Then
        tagBegin = ws.Cells(foundCell.Row - 1, foundCell.Column).Value
        tagBeginColumn = foundCell.Column
    Else
        MsgBox "未找到开始日期"
        Exit Sub
            ' 查找结束日期
    End If

    Set foundCell = ws.Cells.Find( _
    What:="结束日期", _
    After:=ws.Range("A1"), _
    LookIn:=xlFormulas, _
    LookAt:=xlPart, _
    SearchOrder:=xlByRows, _
    SearchDirection:=xlNext _
    )
    If Not foundCell Is Nothing Then
        tagEnd = ws.Cells(foundCell.Row - 1, foundCell.Column).Value
        
            tagEndColumn = foundCell.Column
        Else
            MsgBox "未找到结束日期"
            Exit Sub
            End If
        End Sub
            
            Sub FormatGanttCells(ws As Worksheet, startCol As Long, tagDiff As Long)
                Dim colWidth As Double
                colWidth = 0.2
                
    ' 设置字体和格式
    With ws.Range(ws.Columns(startCol), ws.Columns(startCol + tagDiff))
    .Font.Name = "微软雅黑"
        .Font.Name = "VW Text Office"
        .Font.Size = 10
    End With
    
    With ws.Range(ws.Columns(startCol + 1), ws.Columns(startCol + tagDiff))
        .ColumnWidth = colWidth
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' 设置标题格式
    With ws.Range(ws.Cells(1, startCol + 1), ws.Cells(1, startCol + tagDiff))
        .Font.Size = 14
        .Font.Bold = True
    End With
    
    With ws.Range(ws.Cells(2, startCol + 1), ws.Cells(2, startCol + tagDiff))
        .Font.Size = 11
        .Font.Bold = True
    End With
    
    ' 设置日期格式
    ws.Range(ws.Cells(4, startCol + 1), ws.Cells(4, startCol + tagDiff)).NumberFormat = "yyyy/mm/dd"
    
    ' 自动调整行高
    ws.Range(ws.Rows(1), ws.Rows(4)).EntireRow.AutoFit
    
    ' 设置标签列格式
    With ws.Columns(startCol)
        .Font.Size = 11
        .Font.Bold = True
    End With
End Sub

Sub FillDateData(ws As Worksheet, startCol As Long, tagDiff As Long, tagBegin As Date)
    Dim i As Long
    
    ' 填充开始日期并自动填充
    ws.Cells(4, startCol + 1) = tagBegin
    ws.Cells(4, startCol + 1).AutoFill _
        Destination:=ws.Range(ws.Cells(4, startCol + 1), ws.Cells(4, startCol + tagDiff))
    
    ' 填充年、月、周数据
    For i = startCol + 1 To startCol + tagDiff
        With ws
            .Cells(3, i) = WorksheetFunction.WeekNum(.Cells(4, i), 21)
            .Cells(2, i) = Month(.Cells(4, i))
            .Cells(1, i) = Year(.Cells(4, i))
        End With
    Next i
End Sub

Sub MergeSimilarCells(ws As Worksheet, startCol As Long, tagDiff As Long)
    Dim k As Long, j As Long, m As Long
    
    ' 禁用警告和屏幕更新
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    ' 合并相同年、月、周的单元格
    For k = 1 To 3
        j = startCol + tagDiff
        m = j - 1
        
        Do While j > startCol + 1
            Do While ws.Cells(k, j).Value = ws.Cells(k, m).Value
                m = m - 1
            Loop
            ws.Range(ws.Cells(k, m + 1), ws.Cells(k, j)).Merge
            j = m
            m = m - 1
        Loop
    Next k
    
    ' 恢复设置
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Sub CopyHeaderToBottom(ws As Worksheet, startCol As Long, tagDiff As Long, lastRow As Long)
    ' 复制周和月数据到底部
    ws.Range(ws.Cells(3, startCol), ws.Cells(3, startCol + tagDiff)).Copy _
        Destination:=ws.Range(ws.Cells(lastRow + 1, startCol), ws.Cells(lastRow + 1, startCol + tagDiff))
    
    ws.Range(ws.Cells(2, startCol), ws.Cells(2, startCol + tagDiff)).Copy _
        Destination:=ws.Range(ws.Cells(lastRow + 2, startCol), ws.Cells(lastRow + 2, startCol + tagDiff))
End Sub

Sub FillGanttColors(ws As Worksheet, startCol As Long, tagDiff As Long, tagBegin As Date, _
                    tagBeginColumn As Long, tagEndColumn As Long)
    Dim i As Long, m As Long
    Dim typeColumn As Long, departmentColumn As Long, testColumn As Long
    Dim diffBegin As Long, diffEnd As Long
    Dim colorColumn As Long
    Dim eaColorColumn As Long, ekColorColumn As Long, exColorColumn As Long
    Dim epColorColumn As Long, egColorColumn As Long, ecColorColumn As Long
    
    ' 查找列位置
    typeColumn = FindColumn(ws, "类型")
    departmentColumn = FindColumn(ws, "部门")
    testColumn = FindColumn(ws, "试验内容")
    
    ' 查找颜色列
    eaColorColumn = FindColumn(ws, "EA")
    ekColorColumn = FindColumn(ws, "EK")
    exColorColumn = FindColumn(ws, "EX")
    epColorColumn = FindColumn(ws, "EP")
    egColorColumn = FindColumn(ws, "EG")
    ecColorColumn = FindColumn(ws, "EC")
    
    ' 填充甘特图颜色
    For i = 6 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ws.Cells(i, typeColumn) = "试验项" Then
            ' 计算日期差异
            diffBegin = ws.Cells(i, tagBeginColumn) - tagBegin + 1
            diffEnd = ws.Cells(i, tagEndColumn) - tagBegin + 1
            
            ' 根据部门设置颜色
            Select Case Left(ws.Cells(i, departmentColumn).Value, 2)
                Case "EA": colorColumn = eaColorColumn
                Case "EP": colorColumn = epColorColumn
                Case "EX": colorColumn = exColorColumn
                Case "EC": colorColumn = ecColorColumn
                Case "EG": colorColumn = egColorColumn
                Case "EK": colorColumn = ekColorColumn
            End Select
            
            ' 填充颜色
            With ws.Range(ws.Cells(i, startCol + diffBegin), ws.Cells(i, startCol + diffEnd))
                .Interior.Color = ws.Cells(4, colorColumn).Interior.Color
            End With
            
            ' 添加试验内容
            ws.Cells(i, Int((startCol + diffBegin + startCol + diffEnd) / 2)).Value = ws.Cells(i, testColumn).Value
            
            ' 复制到车辆行
            If m > 0 Then
                ws.Range(ws.Cells(i, startCol + diffBegin), ws.Cells(i, startCol + diffEnd)).Copy _
                    Destination:=ws.Range(ws.Cells(m, startCol + diffBegin), ws.Cells(m, startCol + diffEnd))
                ws.Range(ws.Cells(m, startCol + diffBegin), ws.Cells(m, startCol + diffEnd)).Merge
            End If
        ElseIf ws.Cells(i, typeColumn) = "车辆" Then
            m = i ' 记录车辆行号
        End If
    Next i
End Sub

Function FindColumn(ws As Worksheet, columnName As String) As Long
    Dim foundCell As Range
    
    Set foundCell = ws.Cells.Find( _
        What:=columnName, _
        After:=ws.Range("A1"), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext _
    )
    
    If Not foundCell Is Nothing Then
        FindColumn = foundCell.Column
    Else
        FindColumn = 0
    End If
End Function