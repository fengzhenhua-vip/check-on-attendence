Attribute VB_Name = "标准化宏模块"
Public Const STRowMax As Single = 65536 '按office2007标准设定
Sub 标准VBA排名(ShName, RowB, ICol, OCol)
'版本： V1.0
'作者：冯振华
'时间：2022年1月14日17：20
'功能：对于某一列数值，按大小排名，并标准好名次
'四个参量依次为：工作表名，开始行，输入列，输出列
    Dim ARR As Variant
    Dim TempValue, TempMin As Single
    Dim i, j, k, p, q, RowE As Integer
    Dim ARROut(1 To STRowMax, 1 To 2) As Variant
    If ICol = OCol Then
        MsgBox "排名输出列与输入列相同 ！"                   '输入列与输出列相同时，退出程序，不与排名
        End
    End If
    RowE = Sheets(ShName).Cells(STRowMax, ICol).End(xlUp).Row
    If RowB >= RowE Then
        MsgBox "排序结束行数小于等于开始行数 ！"
        End
    End If
    ARR = Sheets(ShName).Range(Sheets(ShName).Cells(RowB, ICol), Sheets(ShName).Cells(RowE, ICol))
    TempMin = Application.WorksheetFunction.Small(ARR, 1) - 1
    k = 1
ArrBegin:
    q = 0
    TempValue = Application.WorksheetFunction.Large(ARR, 1) ' 此处借用了excel函数求最大，也可以自己采用for next来求最大值
    For i = 1 To UBound(ARR, 1)
        If ARR(i, 1) = TempValue Then
            ARR(i, 1) = TempMin: ARROut(i, 1) = k: q = q + 1
        End If
    Next
    k = k + q
    If k <= UBound(ARR, 1) Then
        GoTo ArrBegin:
    End If
    Sheets(ShName).Range(Sheets(ShName).Cells(RowB, OCol), Sheets(ShName).Cells(RowE, OCol)) = ARROut
End Sub
Sub 标准VBA提取(ShName, RowB, ICol, OShName, ORowB, OCol)
' 版本：V1.0
' 作者：冯振华
' 时间：2022年1月14日17:29
    Dim ARR As Variant
    Dim i, j, k, p, q, RowE As Integer
    Dim ARROut(1 To STRowMax, 1 To 2) As Variant
    Dim sht As Worksheet
    j = 0: q = 0
    For Each sht In ActiveWorkbook.Sheets
        If sht.Name = ShName Then
            j = 1
        End If
        If sht.Name = OShName Then
            q = 1
        End If
    Next
    If j = 0 Then
        MsgBox ShName & "不存在，请填写正确的“输入报表”名称 ！"
        End
    End If
    If q = 0 Then
        MsgBox OShName & "不存在，请填写正确的“输出报表”名称 ！"
        End
    End If
    RowE = Sheets(ShName).Cells(STRowMax, ICol).End(xlUp).Row
    If ShName = OShName Then
        If ICol = OCol Then
            MsgBox "输出列与输入列相同 ！"                   '输入列与输出列相同时，退出程序，不与排名
            End
        End If
    End If
    If RowB >= RowE Then
        MsgBox "结束行数小于等于开始行数 ！"
        End
    End If
    ARR = Sheets(ShName).Range(Sheets(ShName).Cells(RowB, ICol), Sheets(ShName).Cells(RowE, ICol))
    k = 1: p = 0
TQBegin:
    p = p + 1
    ARROut(p, 1) = ARR(k, 1)
    For i = k To UBound(ARR, 1)
        If ARR(i, 1) = ARROut(p, 1) Then
            ARR(i, 1) = Empty
        End If
    Next
    Do While Len(ARR(k, 1)) = 0 And k < UBound(ARR, 1)
        k = k + 1
    Loop
    If k <= UBound(ARR, 1) And Len(ARR(k, 1)) > 0 Then
        GoTo TQBegin:
    End If
    Sheets(OShName).Range(Sheets(OShName).Cells(ORowB, OCol), Sheets(OShName).Cells(ORowB + p - 1, OCol)) = ARROut
'    MsgBox Application.WorksheetFunction.Large(ARROut, 1)
End Sub
Sub 标准VBA拼音排序(IShName, IRowB, ICol)
' 版本：V1.0
' 作者：冯振华
' 时间：2022年1月14日19:19
' 功能：只对一列按拼音排序，不是对于拓展行排序
    Dim RowE As Integer
    RowE = Sheets(IShName).Cells(STRowMax, ICol).End(xlUp).Row
    ActiveWorkbook.Worksheets(IShName).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(IShName).Sort.SortFields.Add2 Key:=Sheets(IShName).Range(Sheets(IShName).Cells(IRowB, ICol), Sheets(IShName).Cells(RowE, ICol)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(IShName).Sort
        .SetRange Sheets(IShName).Range(Sheets(IShName).Cells(IRowB, ICol), Sheets(IShName).Cells(RowE, ICol))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub 标准VBA工具箱()
    标准VBA办公系统.Show
End Sub
Sub test1()
    Call 标准VBA排名("Sheet1", 1, 2, 3)
End Sub
Sub test2()
    Call 标准VBA提取("山东", 3, 2, "Out", 1, 4)
End Sub
Sub test3()
   Call 标准VBA拼音排序("Out", 2, 2)
End Sub
