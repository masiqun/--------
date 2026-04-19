Attribute VB_Name = "模块1"
'==============================================================================
' 模块名称: 模块1
' 功能描述: 车辆甘特图生成模块
'             - 自动查找工作表中的数据范围
'             - 生成年、月、周、日的日期表头
'             - 根据试验类型和部门填充不同的颜色
'             - 生成可视化的甘特图
' 作者: (自动生成)
' 日期: 2026-04-19
'==============================================================================

'==============================================================================
' 过程名称: terminplan
' 功能描述: 主入口过程，负责初始化环境和错误处理
'             这是整个宏的入口点，会被Excel宏引擎直接调用
' 参数说明: 无
' 返回值:   无
' 使用示例: 在Excel中按 Alt+F8，选择 terminplan，点击运行
'==============================================================================
Sub terminplan()
    '----------------------------------------------------------------------------
    ' 性能优化部分：
    ' Excel应用程序设置调整，用于提升宏执行速度
    ' DisableApplicationSettings 会关闭一些实时功能，大幅提升执行效率
    '----------------------------------------------------------------------------
    With Application
        .EnableEvents = False          ' 禁用事件触发，避免每次单元格改变时执行事件代码
        .ScreenUpdating = False        ' 禁用屏幕刷新，避免每次操作都重绘屏幕
        .Calculation = xlCalculationManual  ' 关闭自动计算，改为手动计算
    End With
    
    '----------------------------------------------------------------------------
    ' 错误处理设置：
    ' 如果后续代码发生错误，跳转到 ErrorHandler 处执行
    '----------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    
    '----------------------------------------------------------------------------
    ' 执行主要功能：
    ' 调用 ProcessGanttChart 过程来生成完整的甘特图
    '----------------------------------------------------------------------------
    Call ProcessGanttChart
    
    Exit Sub
    
ErrorHandler:
    '----------------------------------------------------------------------------
    ' 错误处理块：
    ' 当代码执行出错时，显示错误描述信息
    ' vbCritical 常数表示显示严重错误图标
    '----------------------------------------------------------------------------
    MsgBox "错误: " & Err.Description, vbCritical
    
Finally:
    '----------------------------------------------------------------------------
    ' 清理和恢复部分：
    ' 无论代码执行成功还是发生错误，都会执行此处的代码
    ' 恢复所有被禁用的Excel功能，确保Excel恢复正常状态
    '----------------------------------------------------------------------------
    With Application
        .DisplayAlerts = True           ' 恢复警告提示
        .ScreenUpdating = True         ' 恢复屏幕刷新
        .EnableEvents = True           ' 恢复事件触发
        .Calculation = xlCalculationAutomatic  ' 恢复自动计算
    End With
End Sub

'==============================================================================
' 过程名称: ProcessGanttChart
' 功能描述: 甘特图生成的主控制流程
'             协调各个子过程的工作，按顺序执行甘特图生成的所有步骤
' 参数说明: 无
' 返回值:   无
' 调用关系: 由 terminplan 过程调用
'==============================================================================
Sub ProcessGanttChart()
    Dim ws As Worksheet                ' 定义工作表对象变量
    
    '----------------------------------------------------------------------------
    ' 设置工作表引用：
    ' Worksheets("terminplan") 引用名为 "terminplan" 的工作表
    ' 这是甘特图数据所在的工作表
    '----------------------------------------------------------------------------
    Set ws = Worksheets("terminplan")
    
    '----------------------------------------------------------------------------
    ' 步骤1: 获取工作表数据范围
    ' 使用 GetWorksheetDimensions 过程查找数据的最后一行和最后一列
    ' lastRow: 数据区域最下面的行号
    ' lastCol: 数据区域最右边的列号
    '----------------------------------------------------------------------------
    Dim lastRow As Long, lastCol As Long
    Call GetWorksheetDimensions(ws, lastRow, lastCol)
    
    '----------------------------------------------------------------------------
    ' 步骤2: 检查数据有效性
    ' 如果 lastRow 或 lastCol 为 0，说明工作表中没有数据
    ' 此时弹出提示框并退出过程
    '----------------------------------------------------------------------------
    If lastRow = 0 Or lastCol = 0 Then
        MsgBox "工作表为空", vbExclamation
        Exit Sub
    End If
    
    '----------------------------------------------------------------------------
    ' 步骤3: 确定甘特图起始列
    ' GetStartColumn 函数查找第一行中最后一个有数据的列
    ' 甘特图将在此列之后开始生成
    '----------------------------------------------------------------------------
    Dim startCol As Long
    startCol = GetStartColumn(ws, lastCol)
    
    '----------------------------------------------------------------------------
    ' 步骤4: 清理和设置表头
    ' 删除旧数据，设置新的表头标签：年、月、周、日
    ' 这些标签分别代表：Jahr(年)、Monat(月)、Woche(周)、Tag(日)
    '----------------------------------------------------------------------------
    Call SetupGanttHeader(ws, startCol, lastCol)
    
    '----------------------------------------------------------------------------
    ' 步骤5: 查找日期范围
    ' 在工作表中查找"开始日期"和"结束日期"单元格
    ' 获取试验计划的起始和结束日期
    '----------------------------------------------------------------------------
    Dim tagBegin As Date, tagEnd As Date
    Dim tagBeginColumn As Long, tagEndColumn As Long
    Call FindDateRange(ws, tagBegin, tagEnd, tagBeginColumn, tagEndColumn)
    
    '----------------------------------------------------------------------------
    ' 步骤6: 计算日期差异
    ' tagDiff 表示从开始日期到结束日期的总天数
    ' +1 是因为需要包含结束日期当天
    '----------------------------------------------------------------------------
    Dim tagDiff As Long
    tagDiff = tagEnd - tagBegin + 1
    
    '----------------------------------------------------------------------------
    ' 步骤7: 设置单元格格式
    ' 设置字体、字号、列宽、对齐方式等格式
    ' 使甘特图看起来更美观
    '----------------------------------------------------------------------------
    Call FormatGanttCells(ws, startCol, tagDiff)
    
    '----------------------------------------------------------------------------
    ' 步骤8: 填充日期数据
    ' 在表头下方的单元格中填入具体的日期值
    ' 自动填充年、月、周、日的序列数据
    '----------------------------------------------------------------------------
    Call FillDateData(ws, startCol, tagDiff, tagBegin)
    
    '----------------------------------------------------------------------------
    ' 步骤9: 合并相同值的单元格
    ' 将相邻的相同年、月、周的单元格合并，使图表更简洁
    ' 例如：所有1月的单元格会合并成一个
    '----------------------------------------------------------------------------
    Call MergeSimilarCells(ws, startCol, tagDiff)
    
    '----------------------------------------------------------------------------
    ' 步骤10: 复制表头到底部
    ' 将年/月/周/日的表头复制到数据区域底部
    ' 方便在打印或查看时始终能看到表头信息
    '----------------------------------------------------------------------------
    Call CopyHeaderToBottom(ws, startCol, tagDiff, lastRow)
    
    '----------------------------------------------------------------------------
    ' 步骤11: 填充甘特图颜色
    ' 根据试验类型（车辆/试验项）和部门（EA/EG/EC等）
    ' 为甘特图的每个时间段填充对应的颜色
    ' 这是甘特图可视化的核心步骤
    '----------------------------------------------------------------------------
    Call FillGanttColors(ws, startCol, tagDiff, tagBegin, tagBeginColumn, tagEndColumn)
End Sub

'==============================================================================
' 过程名称: GetWorksheetDimensions
' 功能描述: 获取工作表中数据区域的边界
'             查找最后一行和最后一列的位置
' 参数说明: ws        - 输入参数，工作表对象
'           lastRow   - 输出参数，最后一行的行号
'           lastCol   - 输出参数，最后一列的列号
' 返回值:   无（通过ByRef参数返回结果）
'==============================================================================
Sub GetWorksheetDimensions(ws As Worksheet, ByRef lastRow As Long, ByRef lastCol As Long)
    Dim rngLast As Range              ' 定义查找结果单元格变量
    
    '----------------------------------------------------------------------------
    ' 查找最后一行：
    ' 使用 Cells.Find 方法从后向前查找（xlPrevious）
    ' What:="*" 表示查找任何内容
    ' SearchDirection:=xlPrevious 表示从后向前搜索
    '----------------------------------------------------------------------------
    On Error Resume Next              ' 临时忽略错误，防止未找到时出错
    Set rngLast = ws.Cells.Find( _
        What:="*", _                  ' 查找任何非空单元格
        After:=ws.Range("A1"), _      ' 从A1之后开始查找
        LookIn:=xlFormulas, _         ' 在公式中搜索
        LookAt:=xlPart, _             ' 部分匹配
        SearchOrder:=xlByRows, _      ' 按行搜索
        SearchDirection:=xlPrevious _ ' 从后向前搜索
    )
    If Not rngLast Is Nothing Then
        lastRow = rngLast.Row         ' 如果找到，记录行号
    End If
    
    '----------------------------------------------------------------------------
    ' 查找最后一列：
    ' 逻辑与查找最后一行相同，但按列搜索（xlByColumns）
    '----------------------------------------------------------------------------
    Set rngLast = ws.Cells.Find( _
        What:="*", _
        After:=ws.Range("A1"), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByColumns, _   ' 按列搜索
        SearchDirection:=xlPrevious _
    )
    If Not rngLast Is Nothing Then
        lastCol = rngLast.Column      ' 如果找到，记录列号
    End If
    
    On Error GoTo 0                  ' 恢复正常的错误处理
End Sub

'==============================================================================
' 函数名称: GetStartColumn
' 功能描述: 确定甘特图数据起始列
'             查找第一行中最后一个有数据的单元格位置
' 参数说明: ws      - 工作表对象
'           lastCol - 最后一列的位置（备用值）
' 返回值:   起始列的列号
'==============================================================================
Function GetStartColumn(ws As Worksheet, lastCol As Long) As Long
    Dim rngLast As Range
    
    '----------------------------------------------------------------------------
    ' 检查第一行是否有数据：
    ' CountA 函数统计非空单元格数量
    ' 如果大于0，说明第一行有数据
    '----------------------------------------------------------------------------
    If Application.WorksheetFunction.CountA(ws.Rows(1)) Then
        '----------------------------------------------------------------------------
        ' 第一行有数据，从左向右查找第一个非空单元格
        ' SearchDirection:=xlNext 表示从前向后搜索
        '----------------------------------------------------------------------------
        Set rngLast = ws.Rows(1).Cells.Find( _
            What:="*", _
            After:=ws.Range("A1"), _
            LookIn:=xlFormulas, _
            LookAt:=xlPart, _
            SearchOrder:=xlByColumns, _
            SearchDirection:=xlNext _
        )
        
        ' 如果找到，返回该列号；否则使用 lastCol 作为默认值
        If Not rngLast Is Nothing Then
            GetStartColumn = rngLast.Column
        Else
            GetStartColumn = lastCol
        End If
    Else
        ' 如果第一行没有数据，使用 lastCol
        GetStartColumn = lastCol
    End If
End Function

'==============================================================================
' 过程名称: SetupGanttHeader
' 功能描述: 清理并设置甘特图表头
'             删除旧数据，设置年、月、周、日标签
' 参数说明: ws       - 工作表对象
'           startCol - 起始列号
'           lastCol  - 最后一列号
' 返回值:   无
'==============================================================================
Sub SetupGanttHeader(ws As Worksheet, startCol As Long, lastCol As Long)
    '----------------------------------------------------------------------------
    ' 清理现有数据：
    ' 删除从 startCol 开始到最后一列的所有数据
    ' 为生成新的甘特图腾出空间
    '----------------------------------------------------------------------------
    ws.Range(ws.Columns(startCol), ws.Columns(lastCol)).Delete
    ws.Rows().Hidden = False          ' 取消所有隐藏的行
    
    '----------------------------------------------------------------------------
    ' 设置表头标签：
    ' 在指定列的第一到四行设置年/月/周/日标签
    ' 这些标签用于标识时间维度
    '----------------------------------------------------------------------------
    With ws
        .Cells(1, startCol) = "Jahr"  ' 年（Year）
        .Cells(2, startCol) = "Monat" ' 月（Month）
        .Cells(3, startCol) = "Woche" ' 周（Week）
        .Cells(4, startCol) = "Tag"   ' 日（Day）
        .Columns(startCol).EntireColumn.AutoFit  ' 自动调整列宽
    End With
End Sub

'==============================================================================
' 过程名称: FindDateRange
' 功能描述: 查找试验计划的开始和结束日期
'             在工作表中定位"开始日期"和"结束日期"单元格
' 参数说明: ws             - 工作表对象
'           tagBegin       - 输出参数，开始日期
'           tagEnd         - 输出参数，结束日期
'           tagBeginColumn - 输出参数，开始日期所在列
'           tagEndColumn   - 输出参数，结束日期所在列
' 返回值:   无
'==============================================================================
Sub FindDateRange(ws As Worksheet, ByRef tagBegin As Date, ByRef tagEnd As Date, _
                  ByRef tagBeginColumn As Long, ByRef tagEndColumn As Long)
    Dim foundCell As Range
    
    '----------------------------------------------------------------------------
    ' 查找"开始日期"单元格：
    ' 在工作表中搜索包含"开始日期"文本的单元格
    ' 找到后，取其上一行一列的单元格值作为实际日期
    '----------------------------------------------------------------------------
    Set foundCell = ws.Cells.Find( _
        What:="开始日期", _           ' 搜索"开始日期"
        After:=ws.Range("A1"), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext _
    )
    
    If Not foundCell Is Nothing Then
        ' 获取开始日期的实际值（开始日期单元格上方一格）
        tagBegin = ws.Cells(foundCell.Row - 1, foundCell.Column).Value
        tagBeginColumn = foundCell.Column
    Else
        MsgBox "未找到开始日期"        ' 未找到时提示用户
        Exit Sub
    End If
    
    '----------------------------------------------------------------------------
    ' 查找"结束日期"单元格：
    ' 逻辑与查找"开始日期"相同
    '----------------------------------------------------------------------------
    Set foundCell = ws.Cells.Find( _
        What:="结束日期", _
        After:=ws.Range("A1"), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext _
    )
    
    If Not foundCell Is Nothing Then
        ' 获取结束日期的实际值
        tagEnd = ws.Cells(foundCell.Row - 1, foundCell.Column).Value
        tagEndColumn = foundCell.Column
    Else
        MsgBox "未找到结束日期"
        Exit Sub
    End If
End Sub

'==============================================================================
' 过程名称: FormatGanttCells
' 功能描述: 设置甘特图表格的格式
'             包括字体、字号、列宽、对齐方式等
' 参数说明: ws       - 工作表对象
'           startCol - 起始列号
'           tagDiff  - 日期跨度（天数）
' 返回值:   无
'==============================================================================
Sub FormatGanttCells(ws As Worksheet, startCol As Long, tagDiff As Long)
    Dim colWidth As Double
    
    '----------------------------------------------------------------------------
    ' 设置列宽：
    ' colWidth = 0.2 是一个很窄的列宽
    ' 因为甘特图通常有很多天的列，窄列宽可以显示更多天数
    '----------------------------------------------------------------------------
    colWidth = 0.2
    
    '----------------------------------------------------------------------------
    ' 设置字体：
    ' 使用"微软雅黑"和"VW Text Office"两种字体
    ' 字号设置为10号
    '----------------------------------------------------------------------------
    With ws.Range(ws.Columns(startCol), ws.Columns(startCol + tagDiff))
        .Font.Name = "微软雅黑"
        .Font.Name = "VW Text Office"
        .Font.Size = 10
    End With
    
    '----------------------------------------------------------------------------
    ' 设置列宽和对齐：
    ' 将日期列的宽度设置为很窄
    ' 内容居中对齐
    '----------------------------------------------------------------------------
    With ws.Range(ws.Columns(startCol + 1), ws.Columns(startCol + tagDiff))
        .ColumnWidth = colWidth
        .HorizontalAlignment = xlCenter  ' 水平居中
        .VerticalAlignment = xlCenter    ' 垂直居中
    End With
    
    '----------------------------------------------------------------------------
    ' 设置标题行格式：
    ' 第一行（年份）设置为14号加粗字体
    ' 第二行（月份）设置为11号加粗字体
    '----------------------------------------------------------------------------
    With ws.Range(ws.Cells(1, startCol + 1), ws.Cells(1, startCol + tagDiff))
        .Font.Size = 14
        .Font.Bold = True
    End With
    
    With ws.Range(ws.Cells(2, startCol + 1), ws.Cells(2, startCol + tagDiff))
        .Font.Size = 11
        .Font.Bold = True
    End With
    
    '----------------------------------------------------------------------------
    ' 设置日期格式：
    ' 将第4行（日期）的格式设置为 yyyy/mm/dd
    '----------------------------------------------------------------------------
    ws.Range(ws.Cells(4, startCol + 1), ws.Cells(4, startCol + tagDiff)).NumberFormat = "yyyy/mm/dd"
    
    '----------------------------------------------------------------------------
    ' 自动调整行高：
    ' 根据内容自动调整第1-4行的高度
    '----------------------------------------------------------------------------
    ws.Range(ws.Rows(1), ws.Rows(4)).EntireRow.AutoFit
    
    '----------------------------------------------------------------------------
    ' 设置标签列格式：
    ' 第一列（标签列）的字号设置为11号加粗
    '----------------------------------------------------------------------------
    With ws.Columns(startCol)
        .Font.Size = 11
        .Font.Bold = True
    End With
End Sub

'==============================================================================
' 过程名称: FillDateData
' 功能描述: 填充日期序列数据
'             自动填充年、月、周、日的连续日期
' 参数说明: ws       - 工作表对象
'           startCol - 起始列号
'           tagDiff  - 日期跨度
'           tagBegin - 开始日期
' 返回值:   无
'==============================================================================
Sub FillDateData(ws As Worksheet, startCol As Long, tagDiff As Long, tagBegin As Date)
    Dim i As Long
    
    '----------------------------------------------------------------------------
    ' 填充第一个日期：
    ' 在第4行（日期行）的第一个位置填入开始日期
    ' 使用 AutoFill 自动填充功能来填充整个日期序列
    '----------------------------------------------------------------------------
    ws.Cells(4, startCol + 1) = tagBegin
    ws.Cells(4, startCol + 1).AutoFill _
        Destination:=ws.Range(ws.Cells(4, startCol + 1), ws.Cells(4, startCol + tagDiff))
    
    '----------------------------------------------------------------------------
    ' 填充年、月、周数据：
    ' 遍历每个日期单元格，计算并填入对应的年、月、周值
    ' For 循环从 startCol + 1 到 startCol + tagDiff
    '----------------------------------------------------------------------------
    For i = startCol + 1 To startCol + tagDiff
        With ws
            ' WeekNum 函数计算给定日期是当年的第几周
            ' 第二个参数21表示一周从星期一开始
            .Cells(3, i) = WorksheetFunction.WeekNum(.Cells(4, i), 21)
            ' Month 函数提取月份
            .Cells(2, i) = Month(.Cells(4, i))
            ' Year 函数提取年份
            .Cells(1, i) = Year(.Cells(4, i))
        End With
    Next i
End Sub

'==============================================================================
' 过程名称: MergeSimilarCells
' 功能描述: 合并相邻的相同值单元格
'             将年、月、周列中相邻的相同值合并，使图表更简洁
' 参数说明: ws       - 工作表对象
'           startCol - 起始列号
'           tagDiff  - 日期跨度
' 返回值:   无
'==============================================================================
Sub MergeSimilarCells(ws As Worksheet, startCol As Long, tagDiff As Long)
    Dim k As Long, j As Long, m As Long
    
    '----------------------------------------------------------------------------
    ' 性能优化：
    ' 在执行合并操作前禁用屏幕更新和警告提示
    '----------------------------------------------------------------------------
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    '----------------------------------------------------------------------------
    ' 合并相同值单元格：
    ' For 循环处理三行：1=年、2=月、3=周
    ' 使用 Do While 循环从右向左扫描单元格
    ' 当发现连续相同值时，使用 Merge 方法合并它们
    '----------------------------------------------------------------------------
    For k = 1 To 3  ' 1=年、2=月、3=周
        j = startCol + tagDiff  ' j 指向最右边的单元格
        m = j - 1               ' m 指向 j 的左边一个单元格
        
        Do While j > startCol + 1  ' 直到处理完所有单元格
            ' 比较两个单元格的值是否相同
            Do While ws.Cells(k, j).Value = ws.Cells(k, m).Value
                m = m - 1            ' 如果相同，继续向左查找
            Loop
            
            ' 找到不同值时，合并从 m+1 到 j 的所有单元格
            ws.Range(ws.Cells(k, m + 1), ws.Cells(k, j)).Merge
            
            j = m                   ' j 移动到新位置
            m = m - 1               ' m 继续向左移动
        Loop
    Next k
    
    '----------------------------------------------------------------------------
    ' 恢复设置：
    ' 合并完成后，重新启用屏幕更新和警告提示
    '----------------------------------------------------------------------------
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

'==============================================================================
' 过程名称: CopyHeaderToBottom
' 功能描述: 将年/月/周/日表头复制到数据区域底部
'             方便在滚动查看或打印时始终能看到表头信息
' 参数说明: ws       - 工作表对象
'           startCol - 起始列号
'           tagDiff  - 日期跨度
'           lastRow  - 最后一行行号
' 返回值:   无
'==============================================================================
Sub CopyHeaderToBottom(ws As Worksheet, startCol As Long, tagDiff As Long, lastRow As Long)
    '----------------------------------------------------------------------------
    ' 复制周数据到底部：
    ' 将第3行（周）的数据复制到 lastRow + 1 行
    '----------------------------------------------------------------------------
    ws.Range(ws.Cells(3, startCol), ws.Cells(3, startCol + tagDiff)).Copy _
        Destination:=ws.Range(ws.Cells(lastRow + 1, startCol), ws.Cells(lastRow + 1, startCol + tagDiff))
    
    '----------------------------------------------------------------------------
    ' 复制月数据到底部：
    ' 将第2行（月）的数据复制到 lastRow + 2 行
    '----------------------------------------------------------------------------
    ws.Range(ws.Cells(2, startCol), ws.Cells(2, startCol + tagDiff)).Copy _
        Destination:=ws.Range(ws.Cells(lastRow + 2, startCol), ws.Cells(lastRow + 2, startCol + tagDiff))
End Sub

'==============================================================================
' 过程名称: FillGanttColors
' 功能描述: 填充甘特图的颜色
'             根据试验类型和部门，为每个时间段填充对应的颜色
' 参数说明: ws             - 工作表对象
'           startCol       - 起始列号
'           tagDiff        - 日期跨度
'           tagBegin       - 开始日期
'           tagBeginColumn - 开始日期列
'           tagEndColumn   - 结束日期列
' 返回值:   无
'==============================================================================
Sub FillGanttColors(ws As Worksheet, startCol As Long, tagDiff As Long, tagBegin As Date, _
                    tagBeginColumn As Long, tagEndColumn As Long)
    Dim i As Long, m As Long          ' i=循环计数器，m=车辆行行号
    Dim typeColumn As Long             ' 类型列的列号
    Dim departmentColumn As Long        ' 部门列的列号
    Dim testColumn As Long             ' 试验内容列的列号
    Dim diffBegin As Long              ' 开始位置的列偏移
    Dim diffEnd As Long                ' 结束位置的列偏移
    Dim colorColumn As Long            ' 颜色列的列号
    Dim eaColorColumn As Long          ' EA部门颜色列
    Dim ekColorColumn As Long          ' EK部门颜色列
    Dim exColorColumn As Long          ' EX部门颜色列
    Dim epColorColumn As Long          ' EP部门颜色列
    Dim egColorColumn As Long          ' EG部门颜色列
    Dim ecColorColumn As Long          ' EC部门颜色列
    
    '----------------------------------------------------------------------------
    ' 查找关键列的位置：
    ' 使用 FindColumn 函数查找各列的位置
    '----------------------------------------------------------------------------
    typeColumn = FindColumn(ws, "类型")         ' 查找"类型"列
    departmentColumn = FindColumn(ws, "部门")     ' 查找"部门"列
    testColumn = FindColumn(ws, "试验内容")       ' 查找"试验内容"列
    
    '----------------------------------------------------------------------------
    ' 查找各部门的颜色列：
    ' 每个部门（EA、EK、EX、EP、EG、EC）可能有不同的颜色
    ' 这些颜色定义在工作表的特定列中
    '----------------------------------------------------------------------------
    eaColorColumn = FindColumn(ws, "EA")
    ekColorColumn = FindColumn(ws, "EK")
    exColorColumn = FindColumn(ws, "EX")
    epColorColumn = FindColumn(ws, "EP")
    egColorColumn = FindColumn(ws, "EG")
    ecColorColumn = FindColumn(ws, "EC")
    
    '----------------------------------------------------------------------------
    ' 遍历数据行，填充颜色：
    ' 从第6行开始（假设前5行是表头或特殊数据）
    ' 遍历到数据区域的最后一行
    '----------------------------------------------------------------------------
    For i = 6 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        '----------------------------------------------------------------------------
        ' 判断当前行是"试验项"还是"车辆"
        '----------------------------------------------------------------------------
        If ws.Cells(i, typeColumn) = "试验项" Then
            '----------------------------------------------------------------------------
            ' 计算该试验项在甘特图中的开始和结束位置
            ' diffBegin 和 diffEnd 是相对于 startCol 的列偏移量
            '----------------------------------------------------------------------------
            diffBegin = ws.Cells(i, tagBeginColumn) - tagBegin + 1
            diffEnd = ws.Cells(i, tagEndColumn) - tagBegin + 1
            
            '----------------------------------------------------------------------------
            ' 根据部门代码确定颜色列：
            ' 使用 Left 函数提取部门代码的前两个字符
            ' 例如："EAA-1" -> "EA"
            '----------------------------------------------------------------------------
            Select Case Left(ws.Cells(i, departmentColumn).Value, 2)
                Case "EA": colorColumn = eaColorColumn
                Case "EP": colorColumn = epColorColumn
                Case "EX": colorColumn = exColorColumn
                Case "EC": colorColumn = ecColorColumn
                Case "EG": colorColumn = egColorColumn
                Case "EK": colorColumn = ekColorColumn
            End Select
            
            '----------------------------------------------------------------------------
            ' 填充颜色：
            ' 使用 Interior.Color 属性设置单元格的背景颜色
            ' 颜色值取自该部门对应的颜色列（第4行）的颜色
            '----------------------------------------------------------------------------
            With ws.Range(ws.Cells(i, startCol + diffBegin), ws.Cells(i, startCol + diffEnd))
                .Interior.Color = ws.Cells(4, colorColumn).Interior.Color
            End With
            
            '----------------------------------------------------------------------------
            ' 添加试验内容标签：
            ' 在时间段的中间位置添加试验内容描述
            ' 使用 Int 函数取中间位置的整数
            '----------------------------------------------------------------------------
            ws.Cells(i, Int((startCol + diffBegin + startCol + diffEnd) / 2)).Value = ws.Cells(i, testColumn).Value
            
            '----------------------------------------------------------------------------
            ' 复制到车辆行：
            ' 如果存在对应的车辆行（m > 0），将颜色和内容也复制到车辆行
            ' 然后合并车辆行的时间段单元格
            '----------------------------------------------------------------------------
            If m > 0 Then
                ws.Range(ws.Cells(i, startCol + diffBegin), ws.Cells(i, startCol + diffEnd)).Copy _
                    Destination:=ws.Range(ws.Cells(m, startCol + diffBegin), ws.Cells(m, startCol + diffEnd))
                ws.Range(ws.Cells(m, startCol + diffBegin), ws.Cells(m, startCol + diffEnd)).Merge
            End If
            
        ElseIf ws.Cells(i, typeColumn) = "车辆" Then
            '----------------------------------------------------------------------------
            ' 如果是"车辆"行，记录该行号
            ' 后续的试验项将把自己的甘特图也显示在这个车辆行中
            '----------------------------------------------------------------------------
            m = i
        End If
    Next i
End Sub

'==============================================================================
' 函数名称: FindColumn
' 功能描述: 在工作表中查找指定名称的列
'             返回该列的列号，如果未找到返回0
' 参数说明: ws          - 工作表对象
'           columnName  - 要查找的列名称
' 返回值:   列号（未找到返回0）
'==============================================================================
Function FindColumn(ws As Worksheet, columnName As String) As Long
    Dim foundCell As Range
    
    '----------------------------------------------------------------------------
    ' 使用 Find 方法在工作表中搜索指定列名
    ' SearchDirection:=xlNext 表示从前向后搜索，找到第一个匹配项
    '----------------------------------------------------------------------------
    Set foundCell = ws.Cells.Find( _
        What:=columnName, _
        After:=ws.Range("A1"), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext _
    )
    
    '----------------------------------------------------------------------------
    ' 返回结果：
    ' 如果找到，返回该单元格的列号
    ' 如果未找到，返回0
    '----------------------------------------------------------------------------
    If Not foundCell Is Nothing Then
        FindColumn = foundCell.Column
    Else
        FindColumn = 0
    End If
End Function