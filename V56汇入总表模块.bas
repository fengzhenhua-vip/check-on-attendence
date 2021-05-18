Attribute VB_Name = "V56汇入总表模块"
' 汇入总表模块V55
' 作者：冯振华
' 日期：2021年4月29日
' 作用：将每周统计的结果汇入总表，并对时间和违规结果排名，由于每周汇总结果需要人工校准，所认单独设置这一程序

Sub 汇入总表()
    Call 基础变量设置
    Dim ToTeFile As String                                                                                      'TotalTeacherName
    Dim ToTeName As String
    Dim ToHeFile As String                                                                                      'TotalHeadMasterName
    Dim ToHeName As String
    Dim ToTalDate As String
    Dim WriteColumn As Integer
    Dim ToTalSource As Variant
    Dim TeacherSource As Variant
    Dim TBook As Workbook
    Dim HMBook As Workbook
    Dim ToTalOut As Variant
    Dim i, j, k, l, m, n, o, p As Integer
    Dim WriteSwitch As Integer
    Dim Imax, Jmax, TBRmax, TBCmax, HMRmax, HMCmax As Integer
    TotalFolder = OutPath & "\" & Format(Now, "yyyy" & "年") & "统计异常总表"
    ToTeName = NameTeacherUN & Format(Now, "yyyy" & "年") & "总表"
    ToHeName = NameHeadMasterUN & Format(Now, "yyyy" & "年") & "总表"
    ToTeFile = TotalFolder & "\" & ToTeName & ".xlsx"
    ToHeFile = TotalFolder & "\" & ToHeName & ".xlsx"
    Dim SFO As Object
    Set SFO = CreateObject("Scripting.FileSystemObject")                                                        '设SFO为文件夹对象变量

    If SFO.FolderExists(TotalFolder) = False Then
       MkDir TotalFolder
    End If
    If SFO.FileExists(ToTeFile) = False Then
       Call 生成异常总表(ToTeFile, ToTeName)
    End If
    If SFO.FileExists(ToHeFile) = False Then
       Call 生成异常总表(ToHeFile, ToHeName)
    End If
' 异常报表调入数组ToTalSource
    Imax = Range("c65536").End(xlUp).Row
    Jmax = Cells(1, 200).End(xlToLeft).Column
    ToTalSource = Range(Cells(1, 1), Cells(Imax, Jmax))
' 考勤总表调入数组
    If InStr(ToTalSource(2, 3), NameTeacher) > 0 Then
        Set TBook = GetObject(ToTeFile)
        TeacherSource = TBook.Sheets(ToTeName).Range("a1:bb" & 600)
      ElseIf InStr(ToTalSource(2, 3), NameHeadMaster) > 0 Then
        Set TBook = GetObject(ToHeFile)
        TeacherSource = TBook.Sheets(ToHeName).Range("a1:bb" & 600)
    End If
' 取得TeacherSource 的最大非空行和列数
    TBRmax = 1
    Do While TeacherSource(TBRmax, 2) <> 0
        TBRmax = TBRmax + 1
    Loop
    TBRmax = TBRmax - 1
    TBCmax = 1
    Do While TeacherSource(1, TBCmax) <> 0
    TBCmax = TBCmax + 1
    Loop
    TBCmax = TBCmax - 1
 ' 取得TotalSource 最大日期
    ToTalDate = ToTalSource(2, 2)
    For i = 3 To UBound(ToTalSource, 1)
       If CDate(ToTalDate) < CDate(ToTalSource(i, 2)) Then
        ToTalDate = ToTalSource(i, 2)
       End If
    Next
' 获取TeacherSource 数据写入列号WriteColumn
    WriteColumn = 0
    If TBCmax = 3 Then
        WriteColumn = 4
    ElseIf TBCmax > 3 Then
        For j = 4 To TBCmax
            If CDate(TeacherSource(1, j)) = CDate(ToTalDate) Then
                WriteColumn = j
            End If
        Next
        If WriteColumn = 0 Then
            WriteColumn = TBCmax + 1
        End If
    End If
'  提取ToTalSource 的有效数据到SubToTalSource
    l = 0
    k = 0                                                               '记录了SubToTalSource 中的有效数据行数
    Dim SubToTalSource(1 To 600, 1 To 2) As Variant                     '所有教职工数不会超过600，所以暂定为600
    If InStr(ToTalSource(2, 3), NameTeacher) > 0 Then
        For j = 2 To UBound(ToTalSource, 1)
           If ToTalSource(j, 13) > 0 Then                               '教师13列
            k = k + 1
            SubToTalSource(k, 2) = ToTalSource(j, 13)
            l = j
            Do Until ToTalSource(l, 1) <> 0
                l = l - 1
            Loop
            SubToTalSource(k, 1) = ToTalSource(l, 1)
           End If
        Next
    Else
        For j = 2 To UBound(ToTalSource, 1)                             '班主任16列
           If ToTalSource(j, 16) > 0 Then
            k = k + 1
            SubToTalSource(k, 2) = ToTalSource(j, 16)
            l = j
            Do Until ToTalSource(l, 1) <> 0
                l = l - 1
            Loop
            SubToTalSource(k, 1) = ToTalSource(l, 1)
           End If
        Next
    End If
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
    Dim SortC(1 To 600, 1 To 1) As Variant
    Dim SortCMin As Variant
    k = 4                                                                   '最小数据列号
'获得已经输入的数据区最大列号n
    If WriteColumn < TBCmax Then
        n = TBCmax
    Else
        n = WriteColumn
    End If
' 排序处理,当比较日期时应当先强制类型转换为日期再做比较
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
' 对已经输入各列数据求和
    For i = 2 To m                                                           'm记录了TeacherSource中的总数据行数
       TeacherSource(i, 3) = 0
       For j = 4 To n
            TeacherSource(i, 3) = TeacherSource(i, 3) + TeacherSource(i, j)
       Next
    Next
' 根据第3列总数对行排序
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
     If InStr(ToTalSource(2, 3), NameTeacher) > 0 Then
        Workbooks.Open Filename:=ToTeFile
     Else
        Workbooks.Open Filename:=ToHeFile
     End If
     Range("a1:bb" & 600) = TeacherSource
     Call 格式化
     Range(Cells(1, 1), Cells(m, n)).Select
     Call FontSet(NameFont)
     Cells(1, 1).Select
     If InStr(ToTalSource(2, 3), NameTeacher) > 0 Then
        Workbooks(ToTeName).Close savechanges:=True
    Else
        Workbooks(ToHeName).Close savechanges:=True
    End If
End Sub
Sub 生成异常总表(InFile, InName)
    Dim ToTalBook As Workbook
    Application.SheetsInNewWorkbook = 1                                             ' 设置1个Sheet
    Set ToTalBook = Workbooks.Add
    Application.DisplayAlerts = False
        ToTalBook.SaveAs Filename:=InFile
        Sheets(1).name = InName
        Sheets(InName).Range("C2:BB600").NumberFormatLocal = "0;[红色]0"            '设置时间格式
        Sheets(InName).Rows("1:1").NumberFormatLocal = DateFormat                   '设置日期格式
        Cells(1, 1) = "排名"
        Cells(1, 2) = "姓名"
        Cells(1, 3) = "总数"
        ActiveWorkbook.Close savechanges:=True
    Application.DisplayAlerts = True
    Set ToTalBook = Nothing                                                         '取消ToTalBook
End Sub

