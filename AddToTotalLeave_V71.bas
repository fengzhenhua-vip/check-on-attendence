Attribute VB_Name = "AddToTotalLeave_V71"
' 项目：AddToTotalLeave
' 版本：V71
' 作者：冯振华
' 日期：2021年7月16日
' 作用：将每周统计的结果汇入总表，并对时间和违规结果排名，由于每周汇总结果需要人工校准，所认单独设置这一程序
' 日志：2021年7月13日-15日实现了第7代考勤系统，完全从顶层设计重新构建，并优化了显示格式，升级内容如下：
'       1.适配COAV70的格式要求
'       2.优化AddToToalLeave模块，规范代码，提升效率
'       3.提升版本号为V70，与COASMain相同，方便二者匹配的识别
' 日志：修复紧急bug,只有在异常教师和异常班主任报表打开时才执行汇入总表，同时优化了部分代码，兼容性得到提升，加入颜色显示结果，升级版本号V71
'

Sub AddToTotalLeave()
    Application.ScreenUpdating = False
    Call COAConfigSet
    If InStr(ActiveWorkbook.Name, NameTeacherUN) > 0 Or InStr(ActiveWorkbook.Name, NameHeadMasterUN) > 0 Then
        Dim ToTeFile, ToTeName, ToHeFile, ToHeName, ToTalDate, TotalFile, TotalName As String
        Dim WriteColumn As Integer
        Dim ToTalSource, TeacherSource, ToTalOut As Variant
        Dim TBook, HMBook As Workbook
        Dim i, j, k, l, m, n, o, p As Integer
        Dim WriteSwitch As Integer
        Dim Imax, Jmax, TBRmax, TBCmax, HMRmax, HMCmax As Integer
        TotalFolder = OutPath & "\" & Format(Now, "yyyy" & "年") & "统计总表"
        ToTeName = NameTeacherUN & Format(Now, "yyyy" & "年") & "总表"
        ToHeName = NameHeadMasterUN & Format(Now, "yyyy" & "年") & "总表"
        ToTeFile = TotalFolder & "\" & ToTeName & ".xlsx"
        ToHeFile = TotalFolder & "\" & ToHeName & ".xlsx"
        Set SFO = CreateObject("Scripting.FileSystemObject")                                                                             '设SFO为文件夹对象变量
        If SFO.FolderExists(TotalFolder) = False Then
           MkDir TotalFolder
        End If
        If SFO.FileExists(ToTeFile) = False And InStr(ActiveWorkbook.Name, NameTeacherUN) > 0 Then
           Call CreatBook(ToTeFile, ToTeName)
        End If
        If SFO.FileExists(ToHeFile) = False And InStr(ActiveWorkbook.Name, NameHeadMasterUN) > 0 Then
           Call CreatBook(ToHeFile, ToHeName)
        End If
        If InStr(ActiveWorkbook.Name, NameTeacherUN) > 0 Then
            TotalFile = ToTeFile: TotalName = ToTeName
          ElseIf InStr(ActiveWorkbook.Name, NameHeadMasterUN) > 0 Then
            TotalFile = ToHeFile: TotalName = ToHeName
        End If
' 异常报表调入数组ToTalSource
        Imax = Cells(RowMax, 3).End(xlUp).Row
        Jmax = Cells(1, ColMax).End(xlToLeft).Column
        ToTalSource = Range(Cells(1, 1), Cells(Imax, Jmax))
' 考勤总表调入数组
        Set TBook = GetObject(TotalFile)
        TeacherSource = TBook.Sheets(TotalName).Range(TBook.Sheets(TotalName).Cells(1, 1), TBook.Sheets(TotalName).Cells(RowMax, ColMax))
' 取得TeacherSource 的最大非空行和列数
        TBRmax = TBook.Sheets(TotalName).Cells(RowMax, 3).End(xlUp).Row
        TBCmax = TBook.Sheets(TotalName).Cells(1, ColMax).End(xlToLeft).Column
' 取得TotalSource 最大日期
        ToTalDate = Mid(ToTalSource(2, 2), 1, Len(ToTalSource(2, 2)) - 3)
        For i = 4 To Imax
           If Len(ToTalSource(i, 2)) > 3 Then
            If CDate(ToTalDate) < CDate(Mid(ToTalSource(i, 2), 1, Len(ToTalSource(i, 2)) - 3)) Then
             ToTalDate = Mid(ToTalSource(i, 2), 1, Len(ToTalSource(i, 2)) - 3)
            End If
           End If
        Next
' 获取TeacherSource 数据写入列号WriteColumn
        WriteColumn = 4
        If TBCmax > 3 Then
            k = 0
            For j = 4 To TBCmax
                If CDate(TeacherSource(1, j)) = CDate(ToTalDate) Then
                    WriteColumn = j: k = 1
                End If
            Next
            If k = 0 Then
                WriteColumn = TBCmax + 1
            End If
        End If
'  提取ToTalSource 的有效数据到SubToTalSource
        k = 0                                                                                               '记录了SubToTalSource 中的有效数据行数
        Dim SubToTalSource(1 To RowMax, 1 To 2) As Variant
        For j = 2 To Cells(RowMax, 1).End(xlUp).Row - 1
            If ToTalSource(j, 1) > 0 Then
                k = k + 1
                SubToTalSource(k, 1) = ToTalSource(j, 1)
                Do Until ToTalSource(j + 1, 1) <> 0
                    j = j + 1
                Loop
                SubToTalSource(k, 2) = ToTalSource(j, Jmax)
            End If
        Next
' 向目标 TeacherSource 加入有效数据
        If k > 0 Then
            m = TBRmax
            TeacherSource(1, WriteColumn) = ToTalDate
            If TBRmax > 1 Then
                For i = 1 To k
                  l = 0
                  For j = 2 To TBRmax
                    If TeacherSource(j, 2) = SubToTalSource(i, 1) Then
                      TeacherSource(j, WriteColumn) = SubToTalSource(i, 2)
                      l = 1
                    End If
                  Next
                  If l = 0 Then
                    m = m + 1
                    TeacherSource(m, WriteColumn) = SubToTalSource(i, 2)
                    TeacherSource(m, 2) = SubToTalSource(i, 1)
                  End If
                Next
            Else
                For i = 1 To k
                    m = m + 1
                    TeacherSource(m, WriteColumn) = SubToTalSource(i, 2)
                    TeacherSource(m, 2) = SubToTalSource(i, 1)
                Next
            End If
        End If
''以上m记录了TeacherSource 中的有效总行数,下一步工作是排序Sort ，对日期列排序
        Dim SortC(1 To RowMax, 1 To 1) As Variant
        Dim SortCMin As Variant
        k = 4                                                                                                    '最小数据列号
'获得已经输入的数据区最大列号n
        If WriteColumn < TBCmax Then
            n = TBCmax
        Else
            n = WriteColumn
        End If
'排序处理,当比较日期时应当先强制类型转换为日期再做比较
        For p = 4 To n
            k = p
            SortCMin = TeacherSource(1, p)
            For j = p To n
                If CDate(SortCMin) > CDate(TeacherSource(1, j)) Then
                    SortCMin = TeacherSource(1, j)
                    k = j
                End If
            Next
            If p < k Then
                For i = 1 To m
                   SortC(i, 1) = TeacherSource(i, p)
                   TeacherSource(i, p) = TeacherSource(i, k)
                   TeacherSource(i, k) = SortC(i, 1)
                Next
            End If
        Next
'对已经输入各列数据求和
        For i = 2 To m                                                                                           'm=TeacherSource总数据行数
           TeacherSource(i, 3) = 0
           For j = 4 To n
                TeacherSource(i, 3) = TeacherSource(i, 3) + TeacherSource(i, j)
           Next
        Next
'根据第3列总数对行排序
        Dim SortR(1 To 1, 1 To 54) As Variant
        Dim SortRMin As Variant
        For p = 2 To m
            k = p
            SortRMin = TeacherSource(p, 3)
            For i = p To m
                If CInt(SortRMin) > CInt(TeacherSource(i, 3)) Then
                    SortRMin = TeacherSource(i, 3)
                    k = i
                End If
            Next
            If p < k Then
                For j = 2 To n
                 SortR(1, j) = TeacherSource(p, j)
                 TeacherSource(p, j) = TeacherSource(k, j)
                 TeacherSource(k, j) = SortR(1, j)
                Next
            End If
        Next
' 第1列追加排行
        For i = 2 To m
            k = i + 1
            Do While CInt(TeacherSource(k, 3)) = CInt(TeacherSource(i, 3))
                k = k + 1
            Loop
            For p = i To k - 1
                TeacherSource(p, 1) = i - 1
            Next
            i = k - 1
        Next
' 数组写入到目标文件
         Workbooks.Open Filename:=TotalFile
         Range(Cells(1, 1), Cells(m, TBCmax + 1)) = TeacherSource
         Cells(1, WriteColumn).NumberFormatLocal = DateFormat
         k = Cells(1, ColMax).End(xlToLeft).Column
         Call COAFormat(Range(Cells(1, 1), Cells(m, k)))
         Range(Cells(1, 1), Cells(1, k)).Font.Bold = True
         Cells.Interior.ColorIndex = 0
         For i = 2 To m
            For j = 4 To k
                If Len(TeacherSource(i, j)) > 0 Then
                    If j Mod 2 = 1 Then
                        Call COAColor(Cells(i, j), 37, 1)
                    Else
                        Call COAColor(Cells(i, j), 36, 1)
                    End If
                End If
            Next
         Next
         Cells(1, 1).Select
         Workbooks(TotalName).Close savechanges:=True
     End If
     Application.ScreenUpdating = True
End Sub
Sub CreatBook(InFile, InName)
    Dim ToTalBook As Workbook
    Application.SheetsInNewWorkbook = 1                                                                 ' 设置1个Sheet
    Set ToTalBook = Workbooks.Add
    Application.DisplayAlerts = False
        ToTalBook.SaveAs Filename:=InFile
        Cells(1, 1) = "排名"
        Cells(1, 2) = "姓名"
        Cells(1, 3) = "总数"
        Sheets(1).Name = InName
        ActiveWorkbook.Close savechanges:=True
    Application.DisplayAlerts = True
    Set ToTalBook = Nothing                                                                              '取消ToTalBook
End Sub
