Attribute VB_Name = "SSMain_V10"
' 项目： SSMain
' 版本： V10
' 作者：冯振华
' 单位：山东省平原县第一中学
' 邮箱：fengzhenhua@outlook.com
' 博客：https://fengzhenhua-vip.github.io
' 主页：https://github.com/fengzhenhua-vip
' 版权：2021年12月30日--2022年1月1日
' 日志： 完成第一版，实现每次高三联考的自动统计工作2022/01/01，统计时打开本次考试的成绩结果所在页面，否则是不成立的
'
Public SSLimOne, SSLimTwo As Variant
Sub SSMain_V10()
    Dim SSCfgPath, SSCfgFile As String
    Dim SSBook As Workbook
    Dim SSImax, SSJmax As Integer
    Dim SSLimName1, SSLimName2 As String
    Dim OutFolder As String
    Dim SSlimTemp As Variant
    OutFolder = ActiveWorkbook.Path & "\统计结果"
    SSLimName1 = "一线": SSLimName2 = "二线"
    SSCfgPath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\考勤系统" & "\" & "考勤系统配置"
    SSCfgFile = SSCfgPath & "\对比表模板.xlsx"
    Set SFO = CreateObject("Scripting.FileSystemObject")
' 调取配置文件中的一线和二线到对应数组，注意一线与二线的学校数目是完全一样的，所以不必分开来写
    Set SSBook = GetObject(SSCfgFile)
    SSImax = SSBook.Sheets(SSLimName1).Cells(10000, 1).End(xlUp).Row
    SSJmax = SSBook.Sheets(SSLimName1).Cells(2, 1000).End(xlToLeft).Column
    SSLimOne = SSBook.Sheets(SSLimName1).Range(SSBook.Sheets(SSLimName1).Cells(1, 1), SSBook.Sheets(SSLimName1).Cells(SSImax, SSJmax))
    SSLimTwo = SSBook.Sheets(SSLimName2).Range(SSBook.Sheets(SSLimName2).Cells(1, 1), SSBook.Sheets(SSLimName2).Cells(SSImax, SSJmax))
' 在当前目录下创建文件夹和目录
    If SFO.folderExists(OutFolder) = False Then
        MkDir OutFolder
    End If
' 在当前文件后创建Sheet
'    Call SSSRename("源表")
    ActiveSheet.Name = "源表"
    Call SSSheetAdd("总分")
    Call SSSheetAdd("语文")
    Call SSSheetAdd("数学")
    Call SSSheetAdd("英语")
    Call SSSheetAdd("物理")
    Call SSSheetAdd("化学")
    Call SSSheetAdd("生物")
    Call SSSheetAdd("政治")
    Call SSSheetAdd("历史")
    Call SSSheetAdd("地理")
' 取得参加联考的各个学校
    Call GetSchool
' 写入基准数据
    Dim sht As Worksheet
    Dim shtImax, shtJmax As Integer
    SSlimTemp = SSLimOne
    m = 8: n = 1
SSStart:
    For Each sht In ActiveWorkbook.Sheets
        If InStr(sht.Name, "源表") = 0 Then
            For k = 3 To UBound(SSlimTemp, 2)
                If InStr(SSlimTemp(2, k), sht.Name) > 0 Then
                    shtImax = sht.Cells(90000, 1).End(xlUp).Row
                    sht.Cells(shtImax, m) = Empty
                    sht.Cells(3, m) = SSlimTemp(2, k)
                    For i = 4 To shtImax
                       For j = 3 To UBound(SSlimTemp, 1) - 1
                        If InStr(SSlimTemp(j, 1), sht.Cells(i, 1)) > 0 Then
                            sht.Cells(i, m) = SSlimTemp(j, k)
                            sht.Cells(shtImax, m) = sht.Cells(shtImax, m) + SSlimTemp(j, k)
                        End If
                       Next
                    Next
                End If
            Next
        End If
    Next
    If n = 1 Then
        SSlimTemp = SSLimTwo
        m = 3: n = 2
        GoTo SSStart:
    End If
'
    Call 统计分数线
'
    Dim OutOk As Integer
    Application.DisplayAlerts = False
    For Each sht In ActiveWorkbook.Sheets
        Select Case sht.Name
            Case Is = "总分"
                OutOk = 1
            Case Is = "语文"
                OutOk = 1
            Case Is = "数学"
                OutOk = 1
            Case Is = "英语"
                OutOk = 1
            Case Is = "物理"
                OutOk = 1
            Case Is = "化学"
                OutOk = 1
            Case Is = "生物"
                OutOk = 1
            Case Is = "政治"
                OutOk = 1
            Case Is = "历史"
                OutOk = 1
            Case Is = "地理"
                OutOk = 1
            Case Else
                OutOk = 0
        End Select
        If OutOk = 1 Then
            sht.Select
            Call 上色
            Call 加边框
            sht.Copy
            ActiveWorkbook.SaveAs Filename:=OutFolder & "\" & sht.Name & "联考上线对比表", FileFormat:=xlNormal  '将工作簿另存为EXCEL默认格式
            ActiveWorkbook.Close
        End If
    Next
    Application.DisplayAlerts = True
        MsgBox "联考成绩统计完毕!"
End Sub
Sub SSSRename(ReName)
    Dim sht As Worksheet
    Dim shtOk As Integer
    shtOk = 0
    For Each sht In ActiveWorkbook.Sheets
        If InStr(sht.Name, AddName) > 0 Then
            shtOk = 1
        End If
    Next
    If shtOk = 0 Then
        ActiveSheet.Name = AddName
    End If
End Sub

Sub SSSheetAdd(AddName)
    Dim sht As Worksheet
    Dim shtOk As Integer
    shtOk = 0
    For Each sht In ActiveWorkbook.Sheets
        If InStr(sht.Name, AddName) > 0 Then
            shtOk = 1
        End If
    Next
    If shtOk = 0 Then
        Sheets.Add After:=ActiveSheet
        ActiveSheet.Name = AddName
    End If
    Sheets(AddName).Cells(1, 1) = AddName & "联考上线对比表"
    Sheets(AddName).Cells(2, 3) = "其中二线"
    Sheets(AddName).Cells(3, 4) = "名次"
    Sheets(AddName).Cells(2, 5) = "联考二线"
    Sheets(AddName).Cells(3, 6) = "名次"
    Sheets(AddName).Cells(3, 7) = "二线差"
    Sheets(AddName).Cells(2, 8) = "其中一线"
    Sheets(AddName).Cells(3, 9) = "名次"
    Sheets(AddName).Cells(2, 10) = "联考一线"
    Sheets(AddName).Cells(3, 11) = "名次"
    Sheets(AddName).Cells(3, 12) = "一线差"
    Sheets(AddName).Cells(3, 1) = "学校"
    Sheets(AddName).Cells(3, 2) = "人数"
End Sub

Sub AbsorbSchool(AbSht)
    Dim i, j, k, p, q, Imax, Jmax As Integer
    Dim YBook As Variant
    Imax = Sheets("源表").Cells(90000, 1).End(xlUp).Row
    Jmax = Sheets("源表").Cells(2, 1000).End(xlToLeft).Column
    YBook = Sheets("源表").Range(Sheets("源表").Cells(1, 1), Sheets("源表").Cells(Imax, Jmax))
    Sheets(AbSht).Cells(4, 1) = Sheets("源表").Cells(3, 2)
    k = 4
AbsorbBegin:
    For i = 3 To Imax
        If InStr(YBook(i, 2), Sheets(AbSht).Cells(k, 1)) > 0 And Len(YBook(i, 2)) > 0 Then
            YBook(i, 2) = Empty
        End If
    Next
    q = 3
    Do Until Len(YBook(q, 2)) > 0 Or q = Imax
        q = q + 1
    Loop
    If q > 3 And q < Imax Then
        k = k + 1
        Sheets(AbSht).Cells(k, 1) = YBook(q, 2)
        GoTo AbsorbBegin:
    End If
    k = k + 1
    Sheets(AbSht).Cells(k, 1) = "合计"
End Sub
Sub GetSchool()
    Dim sht As Worksheet
    Call AbsorbSchool("总分")
    For Each sht In ActiveWorkbook.Sheets
        If InStr(sht.Name, "源表") = 0 And InStr(sht.Name, "总分") = 0 Then
            For i = 4 To Sheets("总分").Cells(90000, 1).End(xlUp).Row
                sht.Cells(i, 1) = Sheets("总分").Cells(i, 1)
            Next
        End If
    Next
End Sub
Sub 排名X(Xian, Col)
Dim Imax, Jmax As Integer
Imax = Sheets(Xian).Cells(90000, 1).End(xlUp).Row
Jmax = Sheets(Xian).Cells(1, 200).End(xlToLeft).Column
Dim i, j, k, p, q As Integer
Dim outab(1 To 30, 1 To 2) As Variant
Dim intab As Variant
intab = Sheets(Xian).Range(Sheets(Xian).Cells(4, 1), Sheets(Xian).Cells(Imax, 22))
PaiMingStart:
p = p + 1
For i = 1 To Imax - 4
    If CInt(intab(i, Col)) > CInt(outab(p, 2)) And intab(i, Col) > 0 Then
        outab(p, 1) = intab(i, 1)
        outab(p, 2) = intab(i, Col)
    End If
Next
For i = 1 To Imax - 4
    If outab(p, 1) = intab(i, 1) Then
        intab(i, Col) = 0
    End If
Next
If p < Imax - 4 Then
    GoTo PaiMingStart:
End If
For i = 4 To Imax - 1
    For j = 1 To Imax - 4
        If Sheets(Xian).Cells(i, 1) = outab(j, 1) Then
            Sheets(Xian).Cells(i, Col + 1) = j
        End If
    Next
Next
End Sub

Sub 排名()
    For r = 3 To 21 Step 2
     Call 排名X("一线", r)
     Call 排名X("二线", r)
    Next
End Sub

Sub 分科排序X(SortName, SortHang, SortLie)
    Dim Imax, Jmax As Integer
    Sheets(SortName).Select
    Imax = Sheets(SortName).Cells(90000, 1).End(xlUp).Row
    Jmax = Sheets(SortName).Cells(2, 200).End(xlToLeft).Column
    ActiveWorkbook.Worksheets(SortName).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(SortName).Sort.SortFields.Add2 Key:=Range(Cells(SortHang, SortLie), Cells(Imax, SortLie)) _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(SortName).Sort
        .SetRange Range(Cells(SortHang, 1), Cells(Imax, Jmax))
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub 分科排序XX(FKSource, FKName)
' 各副科以赋分后的成绩统计
    Dim FKLie
    Select Case FKName
        Case Is = "总分"
            FKLie = 6
        Case Is = "语文"
            FKLie = 10
        Case Is = "数学"
            FKLie = 14
        Case Is = "英语"
            FKLie = 18
        Case Is = "物理"
            FKLie = 23
        Case Is = "化学"
            FKLie = 28
        Case Is = "生物"
            FKLie = 33
        Case Is = "政治"
            FKLie = 38
        Case Is = "历史"
            FKLie = 43
        Case Is = "地理"
            FKLie = 48
    End Select
    Call 分科排序X(FKSource, 3, FKLie)
    Call 统计分数线XX(FKSource, FKLie, FKName)
End Sub
Sub 统计分数线()
    Call 分科排序XX("源表", "总分")
    Call 分科排序XX("源表", "语文")
    Call 分科排序XX("源表", "数学")
    Call 分科排序XX("源表", "英语")
    Call 分科排序XX("源表", "物理")
    Call 分科排序XX("源表", "化学")
    Call 分科排序XX("源表", "生物")
    Call 分科排序XX("源表", "政治")
    Call 分科排序XX("源表", "历史")
    Call 分科排序XX("源表", "地理")
End Sub

Sub 统计分数线X(Yuan, YuanLie, MuBiao, MBLie)
Dim CurScore As Single
Dim CurRow, CurUp, CurDown As Integer
Dim Imax, Jmax As Integer
Dim CurImax, CurJmax As Integer
Dim YuanDoc, MuBiaoDoc As Variant
Imax = Sheets(Yuan).Cells(90000, 1).End(xlUp).Row
Jmax = Sheets(Yuan).Cells(2, 200).End(xlToLeft).Column
YuanDoc = Sheets(Yuan).Range(Sheets(Yuan).Cells(1, 1), Sheets(Yuan).Cells(Imax, Jmax))
CurImax = Sheets(MuBiao).Cells(90000, 1).End(xlUp).Row
CurJmax = Sheets(MuBiao).Cells(3, 200).End(xlToLeft).Column
MuBiaoDoc = Sheets(MuBiao).Range(Sheets(MuBiao).Cells(1, 1), Sheets(MuBiao).Cells(Imax, Jmax))
CurRow = Sheets(MuBiao).Cells(13, MBLie) + 2
CurScore = Sheets(Yuan).Cells(CurRow, YuanLie)
CurUp = 0: CurDown = 0
Do While YuanDoc(CurRow - CurUp, YuanLie) = YuanDoc(CurRow, YuanLie)
    CurUp = CurUp + 1
Loop
Do While YuanDoc(CurRow + CurDown, YuanLie) = YuanDoc(CurRow, YuanLie)
    CurDown = CurDown + 1
Loop
If CurUp < CurDown Then
    CurScore = YuanDoc(CurRow - CurUp, YuanLie)
Else
    CurScore = YuanDoc(CurRow, YuanLie)
End If
    Sheets(MuBiao).Cells(3, MBLie + 2) = Sheets(MuBiao).Name & "（" & CurScore & ")"
For i = 4 To CurImax - 1
    MuBiaoDoc(i, MBLie + 2) = 0
    MuBiaoDoc(i, 2) = 0
    For j = 3 To Imax
        If InStr(YuanDoc(j, 2), MuBiaoDoc(i, 1)) > 0 And YuanDoc(j, YuanLie) >= CurScore Then
           MuBiaoDoc(i, MBLie + 2) = MuBiaoDoc(i, MBLie + 2) + 1
        End If
        If InStr(YuanDoc(j, 2), MuBiaoDoc(i, 1)) > 0 Then
            MuBiaoDoc(i, 2) = MuBiaoDoc(i, 2) + 1
        End If
    Next
    Sheets(MuBiao).Cells(i, MBLie + 2) = MuBiaoDoc(i, MBLie + 2)
    Sheets(MuBiao).Cells(i, 2) = MuBiaoDoc(i, 2)
Next
For i = 4 To CurImax - 1
 Sheets(MuBiao).Cells(CurImax, 2) = Sheets(MuBiao).Cells(CurImax, 2) + MuBiaoDoc(i, 2)
 Sheets(MuBiao).Cells(CurImax, MBLie + 2) = Sheets(MuBiao).Cells(CurImax, MBLie + 2) + MuBiaoDoc(i, MBLie + 2)
 Sheets(MuBiao).Cells(i, MBLie + 4) = MuBiaoDoc(i, MBLie + 2) - MuBiaoDoc(i, MBLie)
Next
 Sheets(MuBiao).Cells(CurImax, MBLie + 4) = Sheets(MuBiao).Cells(CurImax, MBLie + 2) - MuBiaoDoc(CurImax, MBLie)
End Sub
Sub 统计分数线XX(Source, SouLie, KeMu)
   Call 统计分数线X(Source, SouLie, KeMu, 3)
   Call 统计分数线X(Source, SouLie, KeMu, 8)
   Call 排名X(KeMu, 3)
   Call 排名X(KeMu, 5)
   Call 排名X(KeMu, 8)
   Call 排名X(KeMu, 10)
End Sub

Sub 加边框()
'
    Range("A1:L1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("C2:D2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("E2:F2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("H2:I2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("J2:K2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Selection.CurrentRegion.Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Rows.AutoFit
    Selection.Columns.AutoFit
    Range("A1:L1").Select
    With Selection.Font
        .Name = "微软雅黑"
        .Size = 14
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
End Sub
Sub 上色()
    Dim i, j, k As Integer
    Imax = Cells(90000, 1).End(xlUp).Row
    i = 3: j = 65535
ColorAdd:
    Range(Cells(3, i), Cells(Imax, i + 1)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = j
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Select Case i
        Case Is = 3
            i = 8: j = 65535
            GoTo ColorAdd:
        Case Is = 8
            i = 5: j = 15773696
            GoTo ColorAdd:
        Case Is = 5
            i = 10: j = 15773696
            GoTo ColorAdd:
    End Select
End Sub

