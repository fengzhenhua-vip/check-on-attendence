Attribute VB_Name = "V56得力E考勤模块"
' 得力E考勤系统v55
' 作者：冯振华
' 日期: 2021年4月25日--2021年4月28日
' 属性：针对得力E+考勤机生成的签到数据，进行汇总及上色处理
' 按键：尚未设置
' 说明：将校准表集成到到根据配置文件直接生成为Correct ,这样做的好处在于可以根据自习变化及时调整校准表，而这个集成后多用的时间几乎可以忽略不计，就方便程序来讲采用了这个方案

Public DeLiOriginal As Variant
Sub 得力E考勤()
    Application.ScreenUpdating = False
    Call 基础变量设置
    Call 得力E预处理
    Call 标准通用考勤(DeLiOriginal)
    Application.ScreenUpdating = True
End Sub

Sub 得力E预处理()
    Sheets(1).name = NameOriginal
    Sheets(NameOriginal).Select
    Call 得力E删除空格
    Call 得力E时间刷新
    Call 得力E源表排序
    Call 得力E调取有效数据
End Sub
Sub 得力E源表排序()
' 取消前两行合并的单元格
    Rows("1:2").Select
    Selection.UnMerge
    Dim i As Integer
'表头加A
    For i = 1 To 13
        Cells(1, i + 13) = Cells(1, i)
        Cells(1, i) = "A" & Cells(1, i)
    Next
'按姓名排序
    Columns("C:C").Select
    ActiveWorkbook.Worksheets(NameOriginal).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(NameOriginal).Sort.SortFields.Add Key:=Range("C1"), SortOn _
        :=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(NameOriginal).Sort
        .SetRange Range("A1:Z1928")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
'表头去A
    For i = 14 To 26
        Cells(1, i - 13) = Cells(1, i)
        Cells(1, i) = ""
    Next
End Sub
Sub 得力E时间刷新()
'
' 针对得力e+系统进行的时间设置 ，调用数据到指定数组前，应当先转换成相应的时间格式

    Columns("L:M").NumberFormatLocal = TimeFormat
    Dim i As Integer
    For i = 2 To Sheets(NameOriginal).Range("a65536").End(xlUp).Row
     Cells(i, 12) = Cells(i, 12).Value
     Cells(i, 13) = Cells(i, 13).Value
    Next
End Sub
Sub 得力E删除空格()
' 删除所有空格,当EF列是时间格式时，其中空格也会被完全删除，导致时间上的错误，所以要恢复
    Range("A1").Select
    Cells.Replace What:=" ", Replacement:="", LookAt:=xlPart, SearchOrder:= _
       xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Cells.Replace What:="AM", Replacement:=" AM", LookAt:=xlPart, SearchOrder:= _
       xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Cells.Replace What:="PM", Replacement:=" PM", LookAt:=xlPart, SearchOrder:= _
       xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
End Sub
Sub 得力E调取有效数据()
' 生成标准通用考勤数组格式
    Dim DeLiSource As Variant
    Dim DeRmax As Integer
    Dim DeCmax As Integer
    Dim i, j As Integer
    DeRmax = Sheets(NameOriginal).Range("a65536").End(xlUp).Row
    DeCmax = 13
    Columns("F:F").NumberFormatLocal = DateFormat
    For i = 1 To DeRmax
     Cells(i, 6) = Cells(i, 6).Value
    Next
    DeLiSource = Range(Cells(1, 3), Cells(DeRmax, DeCmax)).Value
    ReDim DeLiOriginal(1 To DeRmax, 1 To 15) As Variant
    j = 0
    For i = 1 To DeRmax
        j = j + 1
        DeLiOriginal(j, 1) = DeLiSource(i, 1)
        DeLiOriginal(j, 2) = DeLiSource(i, 4)
        DeLiOriginal(j, 3) = DeLiSource(i, 7)
        DeLiOriginal(j, 4) = DeLiSource(i, 5)
        DeLiOriginal(j, 5) = DeLiSource(i, 10)
        DeLiOriginal(j, 6) = DeLiSource(i, 11)
    Next
End Sub

