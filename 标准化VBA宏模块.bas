Attribute VB_Name = "标准化VBA宏模块"
Public Const STRowMax As Double = 65536 '按office2007标准设定1048576
Public Const STColMax As Double = 256 '按office2007标准设定16384
Public Const CheckLine As Double = 100  '100行（列）内应该按通常处理的数据来讲是包含了最大值了，所以为了提高效率，设置此值足够了

Sub 标准VBA排名(ShName, RowB, ICol, OCol)
'版本： V1.0
'作者：冯振华
'时间：2022年1月14日17：20
'功能：对于某一列数值，按大小排名，并标准好名次
'四个参量依次为：工作表名，开始行，输入列，输出列
    Dim Arr As Variant
    Dim TempValue, TempMin As Double
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
    Arr = Sheets(ShName).Range(Sheets(ShName).Cells(RowB, ICol), Sheets(ShName).Cells(RowE, ICol))
    TempMin = CDbl(Application.WorksheetFunction.Small(Arr, 1) - 1)
    k = 1
ArrBegin:
    q = 0
    TempValue = CDbl(Application.WorksheetFunction.Large(Arr, 1)) ' 此处借用了excel函数求最大，也可以自己采用for next来求最大值
    For i = 1 To UBound(Arr, 1)
        If Arr(i, 1) = TempValue Then
            Arr(i, 1) = TempMin: ARROut(i, 1) = k: q = q + 1
        End If
    Next
    k = k + q
    If k <= UBound(Arr, 1) Then
        GoTo ArrBegin:
    End If
    Sheets(ShName).Range(Sheets(ShName).Cells(RowB, OCol), Sheets(ShName).Cells(RowE, OCol)) = ARROut
End Sub
Sub 标准VBA提取(ShName, RowB, ICol, OShName, ORowB, OCol)
' 版本：V1.0
' 作者：冯振华
' 时间：2022年1月14日17:29
    Dim Arr As Variant
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
    Arr = Sheets(ShName).Range(Sheets(ShName).Cells(RowB, ICol), Sheets(ShName).Cells(RowE, ICol))
    k = 1: p = 0
TQBegin:
    p = p + 1
    ARROut(p, 1) = Arr(k, 1)
    For i = k To UBound(Arr, 1)
        If Arr(i, 1) = ARROut(p, 1) Then
            Arr(i, 1) = Empty
        End If
    Next
    Do While Len(Arr(k, 1)) = 0 And k < UBound(Arr, 1)
        k = k + 1
    Loop
    If k <= UBound(Arr, 1) And Len(Arr(k, 1)) > 0 Then
        GoTo TQBegin:
    End If
    Sheets(OShName).Range(Sheets(OShName).Cells(ORowB, OCol), Sheets(OShName).Cells(ORowB + p - 1, OCol)) = ARROut
'    MsgBox Application.WorksheetFunction.Large(ARROut, 1)
End Sub
Sub 标准VBA排序(IShName, IRowB, ICol, ShunXu)
' 版本：V1.0
' 作者：冯振华
' 时间：2022年1月14日19:19
' 功能：只对一列按拼音排序，不是对于拓展行排序
    Dim RowE As Integer
    Dim ShengJiang As String
    If ShunXu = 1 Then
        ShengJiang = xlDescending   '降序
    Else
        ShengJiang = xlAscending    '升序
    End If
    RowE = Sheets(IShName).Cells(STRowMax, ICol).End(xlUp).Row
    ActiveWorkbook.Worksheets(IShName).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(IShName).Sort.SortFields.Add2 Key:=Sheets(IShName).Range(Sheets(IShName).Cells(IRowB, ICol), Sheets(IShName).Cells(RowE, ICol)), _
        SortOn:=xlSortOnValues, Order:=ShengJiang, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(IShName).Sort
        .SetRange Sheets(IShName).Range(Sheets(IShName).Cells(IRowB, ICol), Sheets(IShName).Cells(RowE, ICol))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub 标准VBA填充(IShName, IRowB, ICol1, ICol2, OShName, ORowB, OCol1, OCol2)
' 版本：V1.0
' 作者：冯振华
' 时间：2022年1月15日14:32
' 功能：只对一列按拼音排序，不是对于拓展行排序
    Dim IArr, OArr As Variant
    Dim IColMin, IColMax, OColMin, OColMax As Integer
    Dim IRowE, ORowE As Integer
    Dim i, j, k, m, n, p, q As Integer
    If ICol1 < ICol2 Then
        IColMin = ICol1: IColMax = ICol2
    ElseIf ICol1 > ICol2 Then
        IColMin = ICol2: IColMax = ICol1
    Else
        MsgBox IShName & ICol1 & "与" & ICol2 & "相同，请重新输入 ！"
    End If
    IRowE = Sheets(IShName).Cells(STRowMax, ICol1).End(xlUp).Row
    IArr = Sheets(IShName).Range(Sheets(IShName).Cells(IRowB, IColMin), Sheets(IShName).Cells(IRowE, IColMax))
    If OCol1 < OCol2 Then
        OColMin = OCol1: OColMax = OCol2
    ElseIf OCol1 > OCol2 Then
        OColMin = OCol2: OColMax = OCol1
    Else
        MsgBox OShName & OCol1 & "与" & OCol2 & "相同，请重新输入 ！"
    End If
    ORowE = Sheets(OShName).Cells(STRowMax, OCol1).End(xlUp).Row
    OArr = Sheets(OShName).Range(Sheets(OShName).Cells(ORowB, OColMin), Sheets(OShName).Cells(ORowE, OColMax))
    m = ICol1 - IColMin + 1: n = OCol1 - OColMin + 1
    p = ICol2 - IColMin + 1: q = OCol2 - OColMin + 1
    For i = 1 To UBound(OArr, 1)
        For j = 1 To UBound(IArr, 1)
            If OArr(i, n) = IArr(j, m) Then
                OArr(i, q) = IArr(j, p)
            End If
        Next
    Next
    Sheets(OShName).Range(Sheets(OShName).Cells(ORowB, OColMin), Sheets(OShName).Cells(ORowE, OColMax)) = OArr
End Sub
Sub 标准VBA分数段人数(IShName, IRowB, ICol1, ICol2, OShName, ORowB, OCol1, OCol2, SMin, SMax)
' 版本：V1.0
' 作者：冯振华
' 时间：2022年1月15日14:32
' 功能：只对一列按拼音排序，不是对于拓展行排序
    Dim IArr, OArr As Variant
    Dim IColMin, IColMax, OColMin, OColMax As Integer
    Dim IRowE, ORowE As Integer
    Dim i, j, k, m, n, p, q As Integer
    If ICol1 < ICol2 Then
        IColMin = ICol1: IColMax = ICol2
    ElseIf ICol1 > ICol2 Then
        IColMin = ICol2: IColMax = ICol1
    Else
        MsgBox IShName & ICol1 & "与" & ICol2 & "相同，请重新输入 ！"
    End If
    IRowE = Sheets(IShName).Cells(STRowMax, ICol1).End(xlUp).Row
    IArr = Sheets(IShName).Range(Sheets(IShName).Cells(IRowB, IColMin), Sheets(IShName).Cells(IRowE, IColMax))
    If OCol1 < OCol2 Then
        OColMin = OCol1: OColMax = OCol2
    ElseIf OCol1 > OCol2 Then
        OColMin = OCol2: OColMax = OCol1
    Else
        MsgBox OShName & OCol1 & "与" & OCol2 & "相同，请重新输入 ！"
    End If
    ORowE = Sheets(OShName).Cells(STRowMax, OCol1).End(xlUp).Row
    OArr = Sheets(OShName).Range(Sheets(OShName).Cells(ORowB, OColMin), Sheets(OShName).Cells(ORowE, OColMax))
    m = ICol1 - IColMin + 1: n = OCol1 - OColMin + 1
    p = ICol2 - IColMin + 1: q = OCol2 - OColMin + 1
    For i = 1 To UBound(OArr, 1)
        OArr(i, q) = Empty
        For j = 1 To UBound(IArr, 1)
            If InStr(OArr(i, n), IArr(j, m)) > 0 Then
                If CDbl(SMin) <= CDbl(IArr(j, p)) And CDbl(IArr(j, p)) <= CDbl(SMax) Then
                    OArr(i, q) = OArr(i, q) + 1
                End If
            End If
        Next
    Next
    Sheets(OShName).Range(Sheets(OShName).Cells(ORowB, OColMin), Sheets(OShName).Cells(ORowE, OColMax)) = OArr
    Sheets(OShName).Select
End Sub
Public Function ScoreLine(ShName, NumLine, RowB, ICol)
' 版本：V1.0
' 作者：冯振华
' 时间：2022年1月15日14:32
' 功能：取得某一列分数线，目前是按平原县第一中学的本求取得。以上一次某科的人数的分数为基准，向上数的人数与向下人数比较，以较少者分数作为当前考试分数线
    Dim Arr, ArrBak As Variant
    Dim RowE As Integer
    Dim i, j, k As Integer
    Dim UpNum, DownNum As Integer
    RowE = Sheets(ShName).Cells(STRowMax, ICol).End(xlUp).Row
    ArrBak = Sheets(ShName).Range(Sheets(ShName).Cells(RowB, ICol), Sheets(ShName).Cells(RowE, ICol))
    Call 标准VBA排序(ShName, RowB, ICol, 1)
    Arr = Sheets(ShName).Range(Sheets(ShName).Cells(RowB, ICol), Sheets(ShName).Cells(RowE, ICol))
    Sheets(ShName).Range(Sheets(ShName).Cells(RowB, ICol), Sheets(ShName).Cells(RowE, ICol)) = ArrBak
    i = 0
UpBegin:
    If NumLine > i Then
        If Arr(NumLine - i, 1) = Arr(NumLine, 1) Then
            i = i + 1
            GoTo UpBegin:
        Else
            UpNum = i
        End If
    Else
        UpNum = i
    End If
    i = 0
DownBegin:
    If NumLine > i Then
        If Arr(NumLine + i, 1) = Arr(NumLine, 1) Then
            i = i + 1
            GoTo DownBegin:
        Else
            DownNum = i
        End If
    Else
        DownNum = i
    End If
    If UpNum < DownNum Then
        ScoreLine = Arr(NumLine - UpNum, 1)
    Else
        ScoreLine = Arr(NumLine, 1)
     End If
End Function
Sub 标准VBA拆分工作薄(CFName)
    Dim sht As Worksheet
    Dim CFPath, CFFolder As String
    Set SFO = CreateObject("Scripting.FileSystemObject")
    CFPath = ActiveWorkbook.Path
    CFFolder = CFPath & "\" & CFName & "拆分结果"
    If SFO.folderExists(CFFolder) = False Then
        MkDir CFFolder
    End If
    Application.DisplayAlerts = False
    For Each sht In ActiveWorkbook.Sheets
        sht.Select: sht.Copy
        ActiveWorkbook.SaveAs Filename:=CFFolder & "\" & sht.Name & CFName, FileFormat:=xlNormal  '将工作簿另存为EXCEL默认格式
        ActiveWorkbook.Close
    Next
    Application.DisplayAlerts = True
        MsgBox ActiveWorkbook.Name & "拆分完毕 !"
End Sub
Sub 标准VBA合并工作薄()
    Dim fpath, fname As String
    Dim Arr(STRowMax) As String
    Dim CurFil, OtherFil As String
    Dim OArr As Variant
    Dim CurBook As Workbook
    Dim Imax, Jmax As Integer
    Imax = 1000: Jmax = 100
    Dim i, j, k, m, n, p, q As Integer
    Dim sht, osht As Worksheet
    Dim ActiveShtName As String
    For Each sht In ActiveWorkbook.Sheets
        n = 0
        For m = 1 To 10
            If Len(Sheets(sht.Name).Cells(m, 1)) > 0 Then
                n = 1
            End If
        Next
        If n = 1 Then
                 MsgBox "当前工作薄非空，请重新创建一个空的工作薄，再执行合并工作表命令 ！"
                 End
        End If
    Next
    fpath = ActiveWorkbook.Path
    CurFil = ActiveWorkbook.Name
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    fname = Dir(fpath & "\*.xl*")
    i = i + 1
    Arr(i) = fname
    Do While fname <> ""
        fname = Dir
        If fname = "" Then
            Exit Do
        End If
        i = i + 1
        Arr(i) = fname
    Loop
    For p = 1 To i
        If Arr(p) <> CurFil Then
            Set CurBook = GetObject(fpath & "\" & Arr(p))
            If InStr(Arr(p), "xlsx") > 0 Then
                OtherFil = Left(Arr(1), Len(Arr(1)) - 5)
            Else
                OtherFil = Left(Arr(1), Len(Arr(1)) - 4)
            End If
            For Each osht In CurBook.Sheets
                k = 0
                For j = 1 To 10   ' 探测第一列前10行，如果均为空，则认为此工作表是空的，不合并到新表中
                    If Len(CurBook.Sheets(osht.Name).Cells(j, 1)) > 0 Then
                        k = 1
                    End If
                Next
                If k = 1 Then
                    ActiveShtName = Left(Arr(p), Len(Arr(1)) - 5) & "(" & osht.Name & ")"
                    OArr = CurBook.Sheets(osht.Name).Range(CurBook.Sheets(osht.Name).Cells(1, 1), CurBook.Sheets(osht.Name).Cells(Imax, Jmax))
                    q = 0
                    For Each sht In ActiveWorkbook.Sheets
                        If sht.Name = ActiveShtName Then
                            q = 1
                        End If
                    Next
                    If q = 0 Then
                        Sheets.Add After:=ActiveSheet
                        ActiveSheet.Name = ActiveShtName
                    End If
                    Sheets(ActiveShtName).Range(Sheets(ActiveShtName).Cells(1, 1), Sheets(ActiveShtName).Cells(Imax, Jmax)) = OArr
                End If
            Next
        End If
    Next
    For Each sht In ActiveWorkbook.Sheets
        n = 0
        For m = 1 To 10
            If Len(Sheets(sht.Name).Cells(m, 1)) > 0 Then
                n = 1
            End If
        Next
        If n = 0 Then
                 sht.Delete
        End If
    Next
    Application.ScreenUpdating = True
    ActiveWorkbook.Save
    Application.DisplayAlerts = True
End Sub
Sub 标准VBA按行合并工作表()
    Dim sht As Worksheet
    Dim Hfile, Hpath As String
    Dim i, j, k As Integer
    Dim Imax, Jmax As Integer
    Dim IBegin, IEnd As Integer
    Dim Arr As Variant
    Dim CurBook As Workbook
    Set SFO = CreateObject("Scripting.FileSystemObject")
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Hpath = ActiveWorkbook.Path
    Set CurBook = GetObject(Hpath & "\" & ActiveWorkbook.Name)
    If InStr(ActiveWorkbook.Name, ".xlsx") Then
        Hfile = Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 5) & "(按行合并)" & ".xlsx"
    Else
        Hfile = Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 4) & "(按行合并)" & ".xlsx"
    End If
    Hfile = Hpath & "\" & Hfile
    If SFO.fileExists(Hfile) = False Then
        Application.DisplayAlerts = False
        Workbooks.Add.SaveAs Filename:=Hfile
        Sheets(1).Name = "合并"
        For Each sht In CurBook.Sheets
            Imax = CurBook.Sheets(sht.Name).Cells(STRowMax, 1).End(xlUp).Row                 '暂时按第一行和第一列获取最大范围
            Jmax = CurBook.Sheets(sht.Name).Cells(1, STColMax).End(xlToLeft).Column
            For j = 2 To CheckLine
                If Imax < CurBook.Sheets(sht.Name).Cells(STRowMax, j).End(xlUp).Row Then
                    Imax = CurBook.Sheets(sht.Name).Cells(STRowMax, j).End(xlUp).Row
                End If
            Next
            For i = 2 To CheckLine
                If Jmax < CurBook.Sheets(sht.Name).Cells(i, STColMax).End(xlToLeft).Column Then
                    Jmax = CurBook.Sheets(sht.Name).Cells(i, STColMax).End(xlToLeft).Column
                End If
            Next
            Arr = CurBook.Sheets(sht.Name).Range(CurBook.Sheets(sht.Name).Cells(1, 1), CurBook.Sheets(sht.Name).Cells(Imax, Jmax))
            IBegin = Cells(STRowMax, 1).End(xlUp).Row
            For j = 2 To CheckLine
                If IBegin < Cells(STRowMax, j).End(xlUp).Row Then
                    IBegin = Cells(STRowMax, j).End(xlUp).Row
                End If
            Next
            IEnd = IBegin + Imax - 1
            Range(Cells(IBegin, 1), Cells(IEnd, Jmax)) = Arr
        Next
    Else
        MsgBox Hfile & "已经存在，请删除后重新执行合并命令 ！"
    End If
    Application.ScreenUpdating = True
    ActiveWorkbook.Save
    Application.DisplayAlerts = True
End Sub
Sub 标准VBA按列合并工作表()
    Dim sht As Worksheet
    Dim Hfile, Hpath As String
    Dim i, j, k As Integer
    Dim Imax, Jmax As Integer
    Dim IBegin, IEnd As Integer
    Dim Arr As Variant
    Dim CurBook As Workbook
    Set SFO = CreateObject("Scripting.FileSystemObject")
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Hpath = ActiveWorkbook.Path
    Set CurBook = GetObject(Hpath & "\" & ActiveWorkbook.Name)
    If InStr(ActiveWorkbook.Name, ".xlsx") Then
        Hfile = Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 5) & "(按列合并)" & ".xlsx"
    Else
        Hfile = Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 4) & "(按列合并)" & ".xlsx"
    End If
    Hfile = Hpath & "\" & Hfile
    If SFO.fileExists(Hfile) = False Then
        Application.DisplayAlerts = False
        Workbooks.Add.SaveAs Filename:=Hfile
        Sheets(1).Name = "合并"
        For Each sht In CurBook.Sheets
            Imax = CurBook.Sheets(sht.Name).Cells(STRowMax, 1).End(xlUp).Row
            Jmax = CurBook.Sheets(sht.Name).Cells(1, STColMax).End(xlToLeft).Column                 '暂时按第一行和第一列获取最大范围
            For j = 2 To CheckLine
                If Imax < CurBook.Sheets(sht.Name).Cells(STRowMax, j).End(xlUp).Row Then
                    Imax = CurBook.Sheets(sht.Name).Cells(STRowMax, j).End(xlUp).Row
                End If
            Next
            For i = 2 To CheckLine
                If Jmax < CurBook.Sheets(sht.Name).Cells(i, STColMax).End(xlToLeft).Column Then
                    Jmax = CurBook.Sheets(sht.Name).Cells(i, STColMax).End(xlToLeft).Column
                End If
            Next
            Arr = CurBook.Sheets(sht.Name).Range(CurBook.Sheets(sht.Name).Cells(1, 1), CurBook.Sheets(sht.Name).Cells(Imax, Jmax))
            IBegin = Cells(1, STColMax).End(xlToLeft).Column
            For j = 2 To CheckLine
                If IBegin < Cells(j, STColMax).End(xlToLeft).Column Then
                    IBegin = Cells(j, STColMax).End(xlToLeft).Column
                End If
            Next
            IEnd = IBegin + Jmax - 1
            Range(Cells(1, IBegin), Cells(Imax, IEnd)) = Arr
        Next
    Else
        MsgBox Hfile & "已经存在，请删除后重新执行合并命令 ！"
    End If
    Application.ScreenUpdating = True
    ActiveWorkbook.Save
    Application.DisplayAlerts = True
End Sub
Sub 标准VBA转置()
    Dim i, j, k As Integer
    Dim Imax, Jmax, Zmax As Integer
    Dim Arr As Variant
    Dim Brr(1 To CheckLine, 1 To CheckLine) As Variant
    Imax = Sheets(ActiveSheet.Name).Cells(STRowMax, 1).End(xlUp).Row
    Jmax = Sheets(ActiveSheet.Name).Cells(1, STColMax).End(xlToLeft).Column                '暂时按第一行和第一列获取最大范围
    For j = 2 To CheckLine
        If Imax < Sheets(ActiveSheet.Name).Cells(STRowMax, j).End(xlUp).Row Then
            Imax = Sheets(ActiveSheet.Name).Cells(STRowMax, j).End(xlUp).Row
        End If
    Next
    For i = 2 To CheckLine
        If Jmax < Sheets(ActiveSheet.Name).Cells(i, STColMax).End(xlToLeft).Column Then
            Jmax = Sheets(ActiveSheet.Name).Cells(i, STColMax).End(xlToLeft).Column
        End If
    Next
    If Imax > Jmax Then
        Zmax = Imax
    Else
        Zmax = Jmax
    End If
    Arr = Range(Cells(1, 1), Cells(Zmax, Zmax))
    For i = 1 To Zmax
        For j = 1 To Zmax
            Brr(i, j) = Arr(j, i)
        Next
    Next
    Range(Cells(1, 1), Cells(Zmax, Zmax)) = Brr
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
   Call 标准VBA排序("Sheet2", 1, 1, 0)
End Sub
Sub test4()
  Call 标准VBA填充("Sheet1", 2, 6, 7, "Sheet2", 1, 1, 3)
End Sub
Sub test5()
  Call 标准VBA分数段人数("Sheet1", 3, 2, 5, "Sheet2", 2, 1, 2, 0, 800)
End Sub
Sub test6()
    k = ScoreLine("Sheet1", 300, 2, 14)
    MsgBox k
End Sub
Sub test7()
    Call 标准VBA拆分工作表("教师评A率")
End Sub
