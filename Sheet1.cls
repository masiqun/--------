VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub terminplan()
Application.EnableEvents = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

'查找最后一行和最后一列
    Dim ws As Worksheet
    Dim LastRow As Long, LastCol As Long
    Dim rngLast, foundCell As Range
    
    Set ws = Worksheets("terminplan")  ' 操作当前活动工作表
    
    ' 查找最后一行（包括隐藏行）
    On Error Resume Next  ' 处理空表情况
    Set rngLast = ws.Cells.Find( _
        What:="*", _
        After:=ws.Range("A1"), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious _
    )
    If Not rngLast Is Nothing Then LastRow = rngLast.Row
    
    ' 查找最后一列（包括隐藏列）
    Set rngLast = ws.Cells.Find( _
        What:="*", _
        After:=ws.Range("A1"), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByColumns, _
        SearchDirection:=xlPrevious _
    )
    If Not rngLast Is Nothing Then LastCol = rngLast.Column
    
    On Error GoTo 0  ' 恢复错误处理
    
    ' 如果没有数据，提示并退出
    If LastRow = 0 Or LastCol = 0 Then
        MsgBox "工作表为空！", vbExclamation
        Exit Sub
    End If
    
    ' 选中从 A1 到最后一个单元格的区域
    'ws.Range(ws.Cells(1, 1), ws.Cells(LastRow, LastCol)).Select
    'MsgBox "最后一行：" & LastRow & vbCrLf & "最后一列：" & LastCol, vbInformation


'初始化
    Dim startCol As Long
    If Application.WorksheetFunction.CountA(Rows(1)) Then

        Set rngLast = Rows(1).Cells.Find( _
        What:="*", _
        After:=ws.Range("A1"), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByColumns, _
        SearchDirection:=xlNext _
    )
        startCol = rngLast.Column
    Else
        startCol = LastCol
    End If

    'Worksheets("Terminplan").PageSetup.PrintArea = ""
    Range(Columns(startCol), Columns(LastCol)).Delete
    Rows().Hidden = False
    Cells(1, startCol) = "Jahr"
    Cells(2, startCol) = "Monat"
    Cells(3, startCol) = "Woche"
    Cells(4, startCol) = "Tag"

    Columns(startCol).EntireColumn.AutoFit
    
    
'定义变量
    Dim i, j, k, m, Tag_diff, SOP_diff, last_row As Integer
    Dim Col3_width, Col3_widthmm, Col_width, Col_widthmm, Row_height As Double
    Dim Tag_begin_column, Tag_end_column As Long
    Dim Tag_begin, Tag_end As Date
    
'找开始时间和结束时间
' 在第五行中搜索开始日期
    Set foundCell = ws.Cells.Find( _
        What:="开始日期", _
        After:=ws.Range("A1"), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext _
    )
    
    If Not foundCell Is Nothing Then
        ' 返回位置信息
        Tag_begin = Cells(foundCell.Row - 1, foundCell.Column).Value
        Tag_begin_column = foundCell.Column
        'MsgBox "找到内容在：" & foundCell.Address & vbCrLf & _
        '       "行号：" & foundCell.Row & vbCrLf & _
         '      "列号：" & foundCell.Column
    Else
        MsgBox "未找到开始日期"
    End If
    ' 在第五行 列中搜索结束日期
    Set foundCell = ws.Cells.Find( _
        What:="结束日期", _
        After:=ws.Range("A1"), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext _
    )
    
    If Not foundCell Is Nothing Then
        ' 返回位置信息
        Tag_end = Cells(foundCell.Row - 1, foundCell.Column).Value
        Tag_end_column = foundCell.Column
        'MsgBox "找到内容在：" & foundCell.Address & vbCrLf & _
         '      "行号：" & foundCell.Row & vbCrLf & _
          '     "列号：" & foundCell.Column
    Else
        MsgBox "未找到结束日期"
    End If

    Tag_diff = Tag_end - Tag_begin + 1

    
'定义单元格列宽
'    Col3_width = Columns(3).ColumnWidth
 '   If Col3_width > 1 Then
 '       Col3_widthmm = 2.116375 * Col3_width + 1.321594
 '   Else
 '       Col3_widthmm = 3.442196 * Col3_width
 '   End If
 '   Col_widthmm = (297 - 20 - Col3_widthmm) / Tag_diff
 '   If Col_widthmm / 3.442196 > 1 Then
 '       Col_width = (Col_widthmm - 1.321594) / 2.116375
 '   Else
  '      Col_width = WorksheetFunction.Max(Col_widthmm / 3.442196, 0.077)
  '  End If
    Col_width = 0.2
'定义工作表格式
    Worksheets("terminplan").Range(Columns(startCol), Columns(startCol + Tag_diff)).Select
    With Selection
        .Font.Name = "微软雅黑"
        .Font.Name = "VW Text Office"
        .Font.Size = 10
    End With
    Range(Columns(startCol + 1), Columns(startCol + Tag_diff)).Select
    With Selection
        .ColumnWidth = Col_width
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    Range(Cells(1, startCol + 1), Cells(1, startCol + Tag_diff)).Font.Size = 14
    Range(Cells(1, startCol + 1), Cells(1, startCol + Tag_diff)).Font.Bold = True
    Range(Cells(2, startCol + 1), Cells(2, startCol + Tag_diff)).Font.Size = 11
    Range(Cells(2, startCol + 1), Cells(2, startCol + Tag_diff)).Font.Bold = True
    Range(Cells(4, startCol + 1), Cells(4, startCol + Tag_diff)).NumberFormat = "yyyy/mm/dd"
    Range(Rows(1), Rows(4)).EntireRow.AutoFit
    'Range("B2:B3").NumberFormat = "yyyy/mm/dd"
    Columns(startCol).Font.Size = 11
    Columns(startCol).Font.Bold = True
    
'输入日、周、月、年
    Cells(4, startCol + 1) = Tag_begin
    Cells(4, startCol + 1).AutoFill Destination:=Range(Cells(4, startCol + 1), Cells(4, Tag_diff + startCol))
    For i = startCol + 1 To Tag_diff + startCol
        Cells(3, i) = WorksheetFunction.WeekNum(Cells(4, i), 21)
        Cells(2, i) = Month(Cells(4, i))
        Cells(1, i) = Year(Cells(4, i))
'        Select Case Cells(8, 3)
 '       Case "Tag vor SOP"
 '           Cells(8, i) = DateDiff("d", SOP, Cells(7, i))
 '       Case "Wo vor SOP"
 '           Cells(8, i) = DateDiff("ww", SOP, Cells(7, i), vbMonday)
 '       Case "Mo vor SOP"
 '           Cells(8, i) = DateDiff("m", SOP, Cells(7, i))
 '       End Select
    Next
    
'合并单元格时不显示警告,不显示计算过程
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
'合并相同年、月、周单元格
    For k = 1 To 3
            j = Tag_diff + startCol
            m = Tag_diff + startCol - 1
            Do While j > startCol + 1
                Do While Cells(k, j).Value = Cells(k, m).Value
                    m = m - 1
                Loop
                Range(Cells(k, m + 1), Cells(k, j)).Merge
                j = m
                m = m - 1
            Loop
    Next
    

    
'如果列宽太窄，则隐藏Tag和Woche行
    If Col_width < 0.2 Then
    'Rows(7).Hidden = True
    'Rows(6).Hidden = True
    End If
    
'底部日历
'找最后一行
    Set rngLast = ws.Cells.Find( _
        What:="*", _
        After:=ws.Range("A1"), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious _
    )
    If Not rngLast Is Nothing Then LastRow = rngLast.Row

    Range(Cells(3, startCol), Cells(3, startCol + Tag_diff)).Copy Destination:= _
    Range(Cells(LastRow + 1, startCol), Cells(LastRow + 1, startCol + Tag_diff))
    Range(Cells(2, startCol), Cells(2, startCol + Tag_diff)).Copy Destination:= _
    Range(Cells(LastRow + 2, startCol), Cells(LastRow + 2, startCol + Tag_diff))
    

    
'定义打印区域
    'Worksheets("Terminplan").PageSetup.PrintArea = Range(Cells(4, 3), Cells(Worksheets("Terminplan").UsedRange.Rows.Count, Tag_diff + 3)).Address

'填冲试验项
    Dim TypeColumn, DepartmentColumn, Diff_begin, Diff_end, DepartmentColor, TestColumn As Long
    Dim EAColorColumn, EKColorColumn, EXColorColumn, EPColorColumn, EGColorColumn, ECColorColumn, ColorColumn As Long
 '找类型字段
        Set rngLast = ws.Cells.Find( _
        What:="类型", _
        After:=ws.Range("A1"), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext _
    )
    If Not rngLast Is Nothing Then TypeColumn = rngLast.Column
'找部门字段
        Set rngLast = ws.Cells.Find( _
        What:="部门", _
        After:=ws.Range("A1"), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext _
    )
    If Not rngLast Is Nothing Then DepartmentColumn = rngLast.Column
'找试验中文名称字段

        Set rngLast = ws.Cells.Find( _
        What:="试验中文名称", _
        After:=ws.Range("A1"), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext _
    )
    If Not rngLast Is Nothing Then TestColumn = rngLast.Column
'找颜色字段
        Set rngLast = ws.Cells.Find( _
        What:="EK", _
        After:=ws.Range("A1"), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext _
    )
    If Not rngLast Is Nothing Then EKColorColumn = rngLast.Column
        Set rngLast = ws.Cells.Find( _
        What:="EX", _
        After:=ws.Range("A1"), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext _
    )
    If Not rngLast Is Nothing Then EXColorColumn = rngLast.Column
        Set rngLast = ws.Cells.Find( _
        What:="EP", _
        After:=ws.Range("A1"), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext _
    )
    If Not rngLast Is Nothing Then EPColorColumn = rngLast.Column
            Set rngLast = ws.Cells.Find( _
        What:="EG", _
        After:=ws.Range("A1"), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext _
    )
    If Not rngLast Is Nothing Then EGColorColumn = rngLast.Column
        Set rngLast = ws.Cells.Find( _
        What:="EC", _
        After:=ws.Range("A1"), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext _
    )
    If Not rngLast Is Nothing Then ECColorColumn = rngLast.Column
        Set rngLast = ws.Cells.Find( _
        What:="EA", _
        After:=ws.Range("A1"), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext _
    )
    If Not rngLast Is Nothing Then EAColorColumn = rngLast.Column
        
    
    For i = 6 To LastRow

    If Cells(i, TypeColumn) = "试验项" Then
        Diff_begin = Cells(i, Tag_begin_column) - Tag_begin + 1
        Diff_end = Cells(i, Tag_end_column) - Tag_begin + 1
        Select Case Left(Cells(i, DepartmentColumn), 2)
            Case "EA"
                ColorColumn = EAColorColumn
            Case "EP"
                ColorColumn = EPColorColumn
            Case "EX"
                ColorColumn = EXColorColumn
            Case "EC"
                ColorColumn = ECColorColumn
            Case "EG"
                ColorColumn = EGColorColumn
            Case "EK"
                ColorColumn = EKColorColumn
        End Select
        'Range(Cells(i, startCol + Diff_begin), Cells(i, startCol + Diff_end)).Merge '合并单元格
        'Cells(i, startCol + Diff_begin).Interior.Color = Cells(4, ColorColumn).Interior.Color
        Range(Cells(i, startCol + Diff_begin), Cells(i, startCol + Diff_end)).Interior.Color = Cells(4, ColorColumn).Interior.Color
        Cells(i, Int((startCol + Diff_begin + startCol + Diff_end) / 2)).Value = Cells(i, TestColumn).Value
        Range(Cells(i, startCol + Diff_begin), Cells(i, startCol + Diff_end)).Copy Destination:=Range(Cells(m, startCol + Diff_begin), Cells(m, startCol + Diff_end))
        Range(Cells(m, startCol + Diff_begin), Cells(m, startCol + Diff_end)).Merge
    ElseIf Cells(i, TypeColumn) = "车辆" Then
        m = i
    End If
    
    Next
    
Application.DisplayAlerts = True
Application.ScreenUpdating = True

Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic

End Sub





