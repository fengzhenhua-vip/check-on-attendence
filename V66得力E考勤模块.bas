Attribute VB_Name = "V66得力E考勤模块"
' 得力E考勤系统V66
' 作者：冯振华
' 日期: 2021年4月25日--2021年6月12日
' 属性：针对得力E+考勤机生成的签到数据，进行汇总及上色处理
' 按键：尚未设置
' 说明：将校准表集成到到根据配置文件直接生成为Correct ,这样做的好处在于可以根据自习变化及时调整校准表，而这个集成后多用的时间几乎可以忽略不计，就方便程序来讲采用了这个方案
' 日志：在2021年6月11日到12日，对标准通用考勤模块进行了升级，增强的兼容性，同时提高了效率，所以此次对得力E预处理系统进行升级，排序更加简洁合理，故同步版本号为V66.
' 日志：增加Trim函数（用于去除字符串两边空格），增加了Mid函数（用于取得字符串）V66

Public DeLiOriginal As Variant
Public DeRmax As Integer
Public DeCmax As Integer

Sub 得力E考勤()
    Application.ScreenUpdating = False
    Call 基础变量设置
    Call GetDeliOriginal
    Call 标准通用考勤(DeLiOriginal)
    Application.ScreenUpdating = True
End Sub

Sub DeLiEplusSort()
' 取消前两行合并的单元格
    Rows("1:2").UnMerge
'按姓名、日期、班次共同排序
    ActiveWorkbook.Worksheets(NameOriginal).Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets(NameOriginal).Sort.SortFields. _
        Add Key:=Range("C2:C" & DeRmax), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets(NameOriginal).Sort.SortFields. _
        Add Key:=Range("F2:F" & DeRmax), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets(NameOriginal).Sort.SortFields. _
        Add Key:=Range("I2:I" & DeRmax), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(NameOriginal).Sort
        .SetRange Range("A1:Q" & DeRmax)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub GetDeliOriginal()
' 更改Sheets名
    Sheets(1).name = NameOriginal
    Sheets(NameOriginal).Select
' 取得源表行数和列数并排序
    DeRmax = Sheets(NameOriginal).Range("a65536").End(xlUp).Row
    DeCmax = 13
    Call DeLiEplusSort                                                                                        '对源表排序
' 生成标准通用考勤数组格式
    Dim DeLiSource As Variant
    Dim i, j As Integer
    DeLiSource = Range(Cells(1, 3), Cells(DeRmax, DeCmax)).Value
    ReDim DeLiOriginal(1 To DeRmax, 1 To 15) As Variant
    j = 1
    DeLiOriginal(j, 1) = DeLiSource(1, 1)                                                                     '通用程序中有改名的部分，所以此处可以省略而没有影响
    DeLiOriginal(j, 2) = DeLiSource(1, 4)
    DeLiOriginal(j, 3) = DeLiSource(1, 7)
    DeLiOriginal(j, 4) = DeLiSource(1, 5)
    DeLiOriginal(j, 5) = DeLiSource(1, 10)
    DeLiOriginal(j, 6) = DeLiSource(1, 11)
    For i = 2 To DeRmax
        If DeLiSource(i, 1) > 0 Then                                                                          '过滤姓名为空的行
            j = j + 1
            DeLiOriginal(j, 1) = VBA.Trim(Mid(DeLiSource(i, 1), 1, 1)) & VBA.Trim(Mid(DeLiSource(i, 1), 2))   '简单和去除二字姓名中间的空格,不再使用统一的空格消除
            DeLiOriginal(j, 2) = CDate(DeLiSource(i, 4))
            DeLiOriginal(j, 3) = DeLiSource(i, 7)
            DeLiOriginal(j, 4) = DeLiSource(i, 5)
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 以下设置解决了时间格式未设置时自动转换为时间格式,经测试与标准时间有10的负17次方的误差，这在比较时    '
' 间以获得统计信息及对统计结果上色时在时间节点上会出现误判。所以将标准配置信息中的标准时间节点改为     '
' 59秒，而CDate获得的时间皆以00秒结束，故此时将获得严格统计结果，消除以上误差。这个工作在V66版中修改。 '
'                                                                                                      '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If DeLiSource(i, 10) > 0 Then
                DeLiOriginal(j, 5) = CDate(DeLiSource(i, 10))
            End If
            If DeLiSource(i, 11) > 0 Then
                DeLiOriginal(j, 6) = CDate(DeLiSource(i, 11))
            End If
        End If
    Next
End Sub

