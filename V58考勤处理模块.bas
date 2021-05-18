Attribute VB_Name = "V58考勤处理模块"
' 通用考勤系统V58
' 作者：冯振华
' 日期：2021年4月25日--2021年5月13日
' 作用：处理标准化的输入源数组Original
' 说明：将校准表集成到到根据配置文件直接生成为Correct ,这样做的好处在于可以根据自习变化及时调整校准表，而这个集成后多用的时间几乎可以忽略不计，就方便程序来讲采用了这个方案。
' 日志：规范化了全局变量赋值，更加容易控制各辅助表2021/4/27
' 日志：升级了关于假期和任一时间停止考勤的设置，并将配置文件合并到一个文件中，增加了生成报表功能
' 日志：不再保留原始文件，直接将目标文件生成到目标文件夹，然后关闭源文件，主动打开目标文件夹 2021/4/28
' 日志：优化了基础变量设置，使其可以自动根据时间生成对应的文件夹
' 日志：将文件夹生成部分转入通用处理模块，规范代码 2021/4/29
' 日志：校准表和二次校准表加入日期，这样子可以唯一定位到某一行，上一周的情况可以不在下一周受影响2021/4/30
' 日志：统计违规时，加入函数CDate()强制转换成时间格式
' 日志：加入调课表和调课记录2021/5/1
' 日志：优化调课模块2021/5/5
' 日志：根据日期自动选择十月一和五月一日的校准时间、基准时间。系统生成文件及目录，后缀将以文件中的最大时间和最小时间确定，不再以当前发布时间充当后缀
' 日志：统计模块以GoTo优化了代码2021/5/13
' 日志：增加班主任的考勤弹性考核，周一到周五允许1次和正常教师一样早上签到，晚上9：50前可以早退1次，周六周日也允许1次，中午周一到周五允许最多3次和正常教师一样签到，周六周日允许1次
'
'定义全局变量
    Public OutFolder As String
    Public WorkFolder As String
    Public ConfigFolder As String
    Public OutPath As String
    Public OutFileFix As String
    Public NameOriginal As String
    Public NameTeacher As String
    Public NameHeadMaster As String
    Public NameTeacherUN As String
    Public NameHeadMasterUN As String
    Public ConfigPath As String
    Public ConfigFile As String
    Public ConfigBook As Workbook
    Public ConfigSheet1 As String
    Public ConfigSheet2 As String
    Public ConfigSheet3 As String
    Public OriginalSheet1 As String                                                                            '以下4项专为校准表设置
    Public OriginalSheet2 As String
    Public OriginalSheet3 As String
    Public OriginalSheet4 As String
    Public VipSwitch As String
    Public NormalSwitch As String
    Public DateFormat As String
    Public TimeFormat As String
    Public StopSymbol As String
    Public NameFont As String
    Public Correct As Variant
    Public Standard As Variant
    Public Change As Variant
    Public Changed As Variant
    Public CGRmax As Integer
    Public CGCmax As Integer
    Public BeginDate As Date                                                                                    '考勤开始时间
    Public EndDate As Date                                                                                      '考勤结束时间
    Public WuYi As Date
    Public ShiYi As Date
    Public Morning As Integer                                              '正常上班时间违规次数
    Public Afternoon As Integer
    Public Evening As Integer
    Public MorningX As Integer                                             '周末时间违规次数
    Public AfternoonX As Integer
    Public EveningX As Integer
'全局变量赋值
Sub 基础变量设置()
    ConfigPath = "D:\考勤系统"
    ConfigFolder = ConfigPath & "\" & "考勤系统配置"
    OutPath = ConfigPath & "\" & Format(Now, "yyyy" & "年") & "考勤"
    DateFormat = "m""月""d""日"";@"
    TimeFormat = "h:mm;@"
    NameFont = "宋体"
    NameOriginal = "源表"
    NameHeadMasterUN = "异常班主任"
    NameTeacherUN = "异常教师"
    NameHeadMaster = "班主任"
    NameTeacher = "教师"
    ConfigFile = ConfigFolder & "\" & "考勤配置.xlsx"
    ConfigSheet1 = "停止考勤"
    ConfigSheet2 = "请假表"
    OriginalSheet1 = "自习安排"                                                                                 '以下4项专为校准表设置
    OriginalSheet3 = "二次校准"
    OriginalSheet4 = "调课表"
    WuYi = Format(Now, "yyyy") & "/5/1"
    ShiYi = Format(Now, "yyyy") & "/10/1"
    If CDate(WuYi) < CDate(Now) < CDate(ShiYi) Then
        ConfigSheet3 = "五一后基准时间"
        OriginalSheet2 = "五一后校准时间"
    Else
        ConfigSheet3 = "十一后基准时间"
        OriginalSheet2 = "十一后校准时间"
    End If
    VipSwitch = 1                                                                                               '开启vip
    NormalSwitch = 1                                                                                            '0不输出正常报表，1输出
    StopSymbol = "*"
End Sub
Sub 标准通用考勤(Original)
    Dim i, j, k, l, m, n, o, p As Integer
    Dim Source As Variant
    Dim Holiday As Variant
    Dim VipSource As Variant
    Dim Teacher() As Variant
    Dim HeadMaster() As Variant
' 生成校准表时的数组设置
    Dim SelfStudyTable As Variant
    Dim CorrectTable As Variant
    Dim ReCorrectTable As Variant
    Dim CorrectTime As Variant
    Dim SRmax, SCmax, HRmax, HCmax, CRmax, CCmax, STRmax, STCmax, ViRmax, ViCmax, ORmax, OCmax As Integer       '通用数值
    Dim SSRmax, SSCmax, RCTRmax, RCTCmax, CTRmax, CTCmax, CTERmax, CTECmax As Integer                           '专为校准表设置
    Dim SFO As Object
'' 检测配置及输出文件夹是否存在，若不存在，则新建
    Set SFO = CreateObject("Scripting.FileSystemObject")                                                        '设SFO为文件夹对象变量
    If SFO.FolderExists(ConfigPath) = False Then
       MkDir ConfigPath
    End If
    If SFO.FolderExists(ConfigFolder) = False Then
       MkDir ConfigFolder
    End If
    If SFO.FolderExists(OutPath) = False Then
       MkDir OutPath
    End If
'' 获取配置文件
    Set ConfigBook = GetObject(ConfigFile)
''
    ViRmax = ConfigBook.Sheets(ConfigSheet1).Range("a65536").End(xlUp).Row
    ViCmax = ConfigBook.Sheets(ConfigSheet1).Cells(1, 200).End(xlToLeft).Column
    VipSource = ConfigBook.Sheets(ConfigSheet1).Range("a1:d" & ViRmax)
''
    HRmax = ConfigBook.Sheets(ConfigSheet2).Range("a65536").End(xlUp).Row
    HCmax = ConfigBook.Sheets(ConfigSheet2).Cells(1, 200).End(xlToLeft).Column
''''''
    ConfigBook.Sheets(ConfigSheet2).Columns("B:B").NumberFormatLocal = DateFormat
    For i = 1 To HRmax
        ConfigBook.Sheets(ConfigSheet2).Cells(i, 2) = ConfigBook.Sheets(ConfigSheet2).Cells(i, 2).Value
    Next
''''''
    Holiday = ConfigBook.Sheets(ConfigSheet2).Range("a1:f" & HRmax)
''
    STRmax = ConfigBook.Sheets(ConfigSheet3).Range("a65536").End(xlUp).Row
    STCmax = ConfigBook.Sheets(ConfigSheet3).Cells(1, 200).End(xlToLeft).Column
    Standard = ConfigBook.Sheets(ConfigSheet3).Range("a1:e" & STRmax)
''                                                                                                          '以下4项专为校准表设置
    SSRmax = ConfigBook.Sheets(OriginalSheet1).Range("a65536").End(xlUp).Row
    SSCmax = ConfigBook.Sheets(OriginalSheet1).Cells(1, 200).End(xlToLeft).Column
    SelfStudyTable = ConfigBook.Sheets(OriginalSheet1).Range("a1:g" & SSRmax)
''
    CTERmax = ConfigBook.Sheets(OriginalSheet2).Range("a65536").End(xlUp).Row
    CTECmax = ConfigBook.Sheets(OriginalSheet2).Cells(1, 200).End(xlToLeft).Column
    CorrectTime = ConfigBook.Sheets(OriginalSheet2).Range("a1:e" & CTERmax)
''
    RCTRmax = ConfigBook.Sheets(OriginalSheet3).Range("a65536").End(xlUp).Row
    RCTCmax = ConfigBook.Sheets(OriginalSheet3).Cells(1, 200).End(xlToLeft).Column
    ReCorrectTable = ConfigBook.Sheets(OriginalSheet3).Range("a1:i" & RCTRmax)
''
    CGRmax = ConfigBook.Sheets(OriginalSheet4).Range("a65536").End(xlUp).Row                                '调入换课表
    CGCmax = ConfigBook.Sheets(OriginalSheet4).Cells(1, 200).End(xlToLeft).Column
    Change = ConfigBook.Sheets(OriginalSheet4).Range("a1:j" & CGRmax)
''
    ReDim CorrectTable(1 To 2000, 1 To 9) As Variant
    CTRmax = UBound(CorrectTable, 1)
    CTCmax = UBound(CorrectTable, 2)
''
    k = 0
    For i = 2 To SSRmax
        For j = 3 To 7
            If InStr(SelfStudyTable(i, 2), NameTeacher) > 0 Then
''针对教师的设置,检测第1节B和第5节C，影响上午
              If InStr(SelfStudyTable(i, j), "B") + InStr(SelfStudyTable(i, j), "C") > 0 Then
                  k = k + 1
                  CorrectTable(k, 1) = SelfStudyTable(i, 1)
                  CorrectTable(k, 3) = SelfStudyTable(i, 2) & "上午"
                  CorrectTable(k, 4) = SelfStudyTable(1, j)
                  If InStr(SelfStudyTable(i, j), "B") > 0 Then
                        CorrectTable(k, 5) = CorrectTime(2, 2)
                        CorrectTable(k, 6) = CorrectTime(2, 3)
                        CorrectTable(k, 9) = "第1节"
                  End If
                  If InStr(SelfStudyTable(i, j), "C") > 0 Then
                    CorrectTable(k, 7) = CorrectTime(3, 4)
                    CorrectTable(k, 8) = CorrectTime(3, 5)
                    If CorrectTable(k, 9) = 0 Then
                        CorrectTable(k, 9) = "第5节"
                    Else
                        CorrectTable(k, 9) = "第1,5节"
                    End If
                  End If
''上午有第1节或第5节的时候享受下午可以晚到的照顾
                If InStr(SelfStudyTable(i, j), "E") > 0 Then
                     k = k + 1
                     CorrectTable(k, 1) = SelfStudyTable(i, 1)
                     CorrectTable(k, 3) = SelfStudyTable(i, 2) & "下午"
                     CorrectTable(k, 4) = SelfStudyTable(1, j)
                     If InStr(SelfStudyTable(i, j), "D") > 0 Then
                       CorrectTable(k, 8) = CorrectTime(5, 5)
                       CorrectTable(k, 9) = "第6,9节"
                     Else
                       CorrectTable(k, 7) = CorrectTime(4, 4)
                       CorrectTable(k, 9) = "第6节"
                     End If
                Else
                     k = k + 1
                     CorrectTable(k, 1) = SelfStudyTable(i, 1)
                     CorrectTable(k, 3) = SelfStudyTable(i, 2) & "下午"
                     CorrectTable(k, 4) = SelfStudyTable(1, j)
                     CorrectTable(k, 6) = CorrectTime(4, 3)
                     CorrectTable(k, 9) = "★"
                     If InStr(SelfStudyTable(i, j), "D") > 0 Then
                       CorrectTable(k, 8) = CorrectTime(5, 5)
                       CorrectTable(k, 9) = CorrectTable(k, 9) & "第9节"
                     End If
                 End If
               ElseIf InStr(SelfStudyTable(i, j), "D") > 0 Then
                  k = k + 1
                  CorrectTable(k, 1) = SelfStudyTable(i, 1)
                  CorrectTable(k, 3) = SelfStudyTable(i, 2) & "下午"
                  CorrectTable(k, 4) = SelfStudyTable(1, j)
                  CorrectTable(k, 8) = CorrectTime(5, 5)
                  If InStr(SelfStudyTable(i, j), "E") > 0 Then
                    CorrectTable(k, 9) = "第6,9节"
                  Else
                    CorrectTable(k, 9) = "第9节"
                  End If
               ElseIf InStr(SelfStudyTable(i, j), "E") > 0 Then
                  k = k + 1
                  CorrectTable(k, 1) = SelfStudyTable(i, 1)
                  CorrectTable(k, 3) = SelfStudyTable(i, 2) & "下午"
                  CorrectTable(k, 4) = SelfStudyTable(1, j)
                  CorrectTable(k, 7) = CorrectTime(4, 4)
                  CorrectTable(k, 9) = "第6节"
               End If
 ''针对班主任的设置
            Else
                If InStr(SelfStudyTable(i, j), "B") + InStr(SelfStudyTable(i, j), "C") > 0 Then
                  k = k + 1
                  CorrectTable(k, 1) = SelfStudyTable(i, 1)
                  CorrectTable(k, 3) = SelfStudyTable(i, 2) & "上午"
                  CorrectTable(k, 4) = SelfStudyTable(1, j)
                  If InStr(SelfStudyTable(i, j), "B") > 0 Then
                        CorrectTable(k, 9) = "第1节"
                  End If
                  If InStr(SelfStudyTable(i, j), "C") > 0 Then
                    CorrectTable(k, 7) = CorrectTime(3, 4)
                    CorrectTable(k, 8) = CorrectTime(3, 5)
                    If CorrectTable(k, 9) = 0 Then
                        CorrectTable(k, 9) = "第5节"
                    Else
                        CorrectTable(k, 9) = "第1,5节"
                    End If
                  End If
                ElseIf InStr(SelfStudyTable(i, j), "D") > 0 Then
                  k = k + 1
                  CorrectTable(k, 1) = SelfStudyTable(i, 1)
                  CorrectTable(k, 3) = SelfStudyTable(i, 2) & "下午"
                  CorrectTable(k, 4) = SelfStudyTable(1, j)
                  CorrectTable(k, 9) = "第9节"
                End If
            End If
        Next
    Next
'二次校准,n记录匹配数
   o = 0
   For i = 1 To CTRmax
        If CorrectTable(i, 1) <> 0 Then
         o = o + 1
        End If
   Next
   l = o
   For j = 2 To RCTRmax
       n = 0
       For i = 1 To o
           If InStr(CorrectTable(i, 1), ReCorrectTable(j, 1)) > 0 Then
                If InStr(CorrectTable(i, 2), ReCorrectTable(j, 2)) > 0 Then
                    If InStr(CorrectTable(i, 3), ReCorrectTable(j, 3)) > 0 Then
                        If InStr(CorrectTable(i, 4), ReCorrectTable(j, 4)) > 0 Then
                            For k = 5 To 8
                              CorrectTable(i, k) = ReCorrectTable(j, k)
                            Next
                            If InStr(CorrectTable(i, 9), "★") > 0 Then
                                  If InStr(CorrectTable(i, 9), "第9节") > 0 Then
                                      CorrectTable(i, 9) = ReCorrectTable(j, 9) & "第9节"
                                  Else
                                      CorrectTable(i, 9) = ReCorrectTable(j, 9)
                                  End If
                            Else
                                  CorrectTable(i, 9) = ReCorrectTable(j, 9) & CorrectTable(i, 9)
                            End If
                            n = 1
                        End If
                    End If
                End If
           End If
       Next
       If n = 0 Then
            l = l + 1
            For m = 1 To 8
                CorrectTable(l, m) = ReCorrectTable(j, m)
            Next
            CorrectTable(l, 9) = ReCorrectTable(j, 9)
       End If
    Next
''取得CorrectTable 的非空单元格数
    p = 0
    For i = 1 To CTRmax
       If CorrectTable(i, 1) > 0 Then
        p = p + 1
       End If
    Next
''CorrectTable 加入调课表Change
l = p                                                                '获得当前CorrectTable中的总行数l
For i = 2 To CGRmax
' 调出校准
    p = p + 1
    For j = 1 To 4
        CorrectTable(p, j) = Change(i, j)
    Next
    If InStr(Change(i, 5), "B") > 0 Then
        CorrectTable(p, 6) = CorrectTime(2, 2)
    End If
    If InStr(Change(i, 5), "C") > 0 Then
        CorrectTable(p, 7) = CorrectTime(3, 5)
    End If
    If InStr(Change(i, 5), "D") > 0 Then
        CorrectTable(p, 7) = CorrectTime(5, 5)
    End If
    CorrectTable(p, 9) = "↑"
    m = 0
    For k = 1 To l
         If InStr(CorrectTable(p, 1), CorrectTable(k, 1)) > 0 Then
             If InStr(CorrectTable(p, 2), CorrectTable(k, 2)) > 0 Then
                If InStr(CorrectTable(p, 3), CorrectTable(k, 3)) > 0 Then
                   If InStr(CorrectTable(p, 4), CorrectTable(k, 4)) > 0 Then
                    m = k
                   End If
                End If
            End If
        End If
    Next
    If m > 0 Then
        CorrectTable(p, 9) = CorrectTable(p, 9) & CorrectTable(m, 9)
    End If
' 如果调出BC则下午不再享受晚来20分钟待遇
    If InStr(Change(i, 5), "B") + InStr(Change(i, 5), "C") > 0 Then
        p = p + 1
        For j = 1 To 4
            CorrectTable(p, j) = CorrectTable(p - 1, j)
        Next
        CorrectTable(p, 3) = "教师下午"
        CorrectTable(p, 5) = CorrectTime(2, 2)
        m = 0
        For k = 1 To l
             If InStr(CorrectTable(p, 1), CorrectTable(k, 1)) > 0 Then
                 If InStr(CorrectTable(p, 2), CorrectTable(k, 2)) > 0 Then
                    If InStr(CorrectTable(p, 3), CorrectTable(k, 3)) > 0 Then
                       If InStr(CorrectTable(p, 4), CorrectTable(k, 4)) > 0 Then
                        m = k
                       End If
                    End If
                End If
            End If
        Next
        If m > 0 Then
            If InStr(CorrectTable(m, 9), "第9节") > 0 Then
                CorrectTable(m, 9) = "第9节"
            End If
        End If
    End If
' 调入校准
    p = p + 1
    For j = 1 To 4
        CorrectTable(p, j) = Change(i, j + 5)
    Next
    If InStr(Change(i, 10), "B") > 0 Then
        CorrectTable(p, 5) = CorrectTime(2, 2)
    End If
    If InStr(Change(i, 10), "C") > 0 Then
        CorrectTable(p, 8) = CorrectTime(3, 5)
    End If
    If InStr(Change(i, 10), "D") > 0 Then
        CorrectTable(p, 8) = CorrectTime(5, 5)
    End If
    CorrectTable(p, 9) = "↓"
    m = 0
    For k = 1 To l
         If InStr(Change(i, 6), CorrectTable(k, 1)) > 0 Then
             If InStr(Change(i, 7), CorrectTable(k, 2)) > 0 Then
                If InStr(Change(i, 8), CorrectTable(k, 3)) > 0 Then
                   If InStr(Change(i, 9), CorrectTable(k, 4)) > 0 Then
                    m = k
                   End If
                End If
            End If
        End If
    Next
    If m > 0 Then
        CorrectTable(p, 9) = CorrectTable(p, 9) & CorrectTable(m, 9)
    End If
' 如果调入BC则下午享受20分钟晚到待遇
    If InStr(Change(i, 10), "B") + InStr(Change(i, 10), "C") > 0 Then
        m = 0
        For k = 1 To l
             If InStr(Change(i, 6), CorrectTable(k, 1)) > 0 Then
                 If InStr(Change(i, 7), CorrectTable(k, 2)) > 0 Then
                    If InStr(Change(i, 8), CorrectTable(k, 3)) > 0 Then
                       If InStr(Change(i, 9), CorrectTable(k, 4)) > 0 Then
                        m = k
                       End If
                    End If
                End If
            End If
        Next
        If m > 0 Then
            If InStr(CorrectTable(m, 9), "★") > 0 Then
            Else
                p = p + 1
                For j = 1 To 4
                    CorrectTable(p, j) = CorrectTable(p - 1, j)
                Next
                CorrectTable(p, 3) = "教师下午"
                CorrectTable(p, 6) = CorrectTime(2, 2)
                CorrectTable(p, 9) = "★" & CorrectTable(m, 9)
            End If
        Else
            p = p + 1
            For j = 1 To 4
                CorrectTable(p, j) = CorrectTable(p - 1, j)
            Next
            CorrectTable(p, 3) = "教师下午"
            CorrectTable(p, 6) = CorrectTime(2, 2)
            CorrectTable(p, 9) = "★"
        End If
    End If
Next
'生成Changed 用以在合并单元格中记录调课情况
o = 2 * CGRmax
ReDim Changed(1 To o, 1 To 10) As Variant
k = 0
For i = 1 To CGRmax
  k = k + 1
  For j = 1 To 4
    Changed(k, j) = Change(i, j)
  Next
  Changed(k, 5) = Format(Change(i, 2), DateFormat)
  If InStr(Change(i, 5), "B") + InStr(Change(i, 5), "C") + InStr(Change(i, 5), "D") > 0 Then
        If InStr(Change(i, 5), "B") > 0 Then
          Changed(k, 5) = Changed(k, 5) & " 第1节"
        End If
        If InStr(Change(i, 5), "C") > 0 Then
           Changed(k, 5) = Changed(k, 5) & " 第5节"
        End If
        If InStr(Change(i, 5), "D") > 0 Then
          Changed(k, 5) = Changed(k, 5) & " 第9节"
        End If
  Else
        Changed(k, 5) = Changed(k, 5) & " " & Change(i, 5)
  End If
  For j = 6 To 9
    Changed(k, j) = Change(i, j)
  Next
  Changed(k, 10) = Format(Change(i, 7), DateFormat)
  If InStr(Change(i, 10), "B") + InStr(Change(i, 10), "C") + InStr(Change(i, 10), "D") > 0 Then
        If InStr(Change(i, 10), "B") > 0 Then
          Changed(k, 10) = Changed(k, 10) & " 第1节"
        End If
        If InStr(Change(i, 10), "C") > 0 Then
           Changed(k, 10) = Changed(k, 10) & " 第5节"
        End If
        If InStr(Change(i, 10), "D") > 0 Then
          Changed(k, 10) = Changed(k, 10) & " 第9节"
        End If
  Else
        Changed(k, 10) = Changed(k, 10) & " " & Change(i, 10)
  End If
  k = k + 1
  For j = 1 To 5
    Changed(k, j) = Changed(k - 1, j + 5)
    Changed(k, j + 5) = Changed(k - 1, j)
  Next
Next
' 同一个人的调课情况并到一块，只保留一个名字
For i = 2 To 2 * CGRmax
    For j = 2 To 2 * CGRmax
        If i > j Then
            If Changed(i, 1) = Changed(j, 1) Then
                Changed(j, 1) = 0
            End If
        End If
    Next
Next
'' 获得Correct(校准表)
    ReDim Correct(1 To p, 1 To 9) As Variant
    For i = 1 To p
        For j = 1 To 9
            Correct(i, j) = CorrectTable(i, j)
        Next
    Next
    CRmax = UBound(Correct, 1)
    CCmax = UBound(Correct, 2)
''设置vip处理模块
If VipSwitch = 0 Then
    Source = Original
Else
    ORmax = UBound(Original, 1)
    OCmax = UBound(Original, 2)
    ReDim Source(1 To ORmax, 1 To OCmax) As Variant
'  过滤VIP人员
    k = 0
    For i = 1 To ORmax
        l = 0
        If Original(i, 3) <> 0 Then
' 将第i行和VipSource 列表比对，以l记录比对数，如果比对数为1，则表明有匹配项产生
            For j = 2 To ViRmax
                 If InStr(Original(i, 1), VipSource(j, 1)) > 0 Then
                     If InStr(Original(i, 2), VipSource(j, 2)) > 0 Then
                        If InStr(Original(i, 3), VipSource(j, 3)) > 0 Then
                           If InStr(Original(i, 4), VipSource(j, 4)) > 0 Then
                            l = 1
                           End If
                       End If
                      End If
                 ElseIf InStr(VipSource(j, 1), StopSymbol) > 0 Then
                     If InStr(Original(i, 2), VipSource(j, 2)) > 0 Then
                        If InStr(Original(i, 3), VipSource(j, 3)) > 0 Then
                           If InStr(Original(i, 4), VipSource(j, 4)) > 0 Then
                            l = 1
                           End If
                       End If
                      End If
                 End If
            Next
'当l=0，则说明此行不在VipSource 中，则导入到Source
            If l = 0 Then
                k = k + 1
                For m = 1 To 6
                    Source(k, m) = Original(i, m)
                Next
            End If
        End If
     Next
End If
' Source 表头设置
    Source(1, 3) = "班次"
    Source(1, 4) = "自习"
    Source(1, 5) = "签到"
    Source(1, 6) = "签退"
    Source(1, 7) = "上迟"
    Source(1, 8) = "上退"
    Source(1, 9) = "上漏"
    Source(1, 10) = "下迟"
    Source(1, 11) = "下退"
    Source(1, 12) = "下漏"
    Source(1, 13) = "晚迟"
    Source(1, 14) = "晚退"
    Source(1, 15) = "晚漏"
' Source 行数和列数设置
    SRmax = UBound(Original, 1)
    SCmax = UBound(Original, 2)
' 将请假信息加入到Source
    For i = 2 To HRmax
        For j = 2 To SRmax
           If InStr(Source(j, 1), Holiday(i, 1)) > 0 Then
            If InStr(Source(j, 2), Holiday(i, 2)) > 0 Then
                If InStr(Source(j, 3), Holiday(i, 3)) > 0 Then
                    If InStr(Source(j, 4), Holiday(i, 4)) > 0 Then
                       If Holiday(i, 5) > 0 Then
                          Source(j, 5) = Holiday(i, 5)
                       End If
                       If Holiday(i, 6) > 0 Then
                          Source(j, 6) = Holiday(i, 6)
                       End If
                    End If
                  End If
              End If
           ElseIf InStr(Holiday(i, 1), StopSymbol) > 0 Then
            If InStr(Source(j, 2), Holiday(i, 2)) > 0 Then
                If InStr(Source(j, 3), Holiday(i, 3)) > 0 Then
                    If InStr(Source(j, 4), Holiday(i, 4)) > 0 Then
                       If Holiday(i, 5) > 0 Then
                          Source(j, 5) = Holiday(i, 5)
                       End If
                       If Holiday(i, 6) > 0 Then
                          Source(j, 6) = Holiday(i, 6)
                       End If
                    End If
                  End If
              End If
           End If
        Next
    Next
'校准Source中的签到签退数据
    For i = 2 To SRmax
        For j = 1 To CRmax
            If InStr(Source(i, 1), Correct(j, 1)) > 0 Then
              If InStr(Source(i, 2), Correct(j, 2)) > 0 Then
                If InStr(Source(i, 3), Correct(j, 3)) > 0 Then
                    If InStr(Source(i, 4), Correct(j, 4)) > 0 Then
                        If IsNumeric(Source(i, 5)) Then
                            If Source(i, 5) = 0 Then
                            Else
                              Source(i, 5) = Source(i, 5) + Correct(j, 5)
                              Source(i, 5) = Source(i, 5) - Correct(j, 6)
                            End If
                        End If
                        If IsNumeric(Source(i, 6)) Then
                            If Source(i, 6) = 0 Then
                            Else
                              Source(i, 6) = Source(i, 6) + Correct(j, 7)
                              Source(i, 6) = Source(i, 6) - Correct(j, 8)
                            End If
                        End If
                    End If
                End If
              End If
            End If
        Next
    Next
'据校准后的Source生成统计数据
    For i = 2 To SRmax
     If InStr(Source(i, 3), Standard(2, 1)) > 0 Then
        If IsNumeric(Source(i, 5)) Then
            If Source(i, 5) = 0 Then
               Source(i, 9) = Source(i, 9) + 1
            ElseIf CDate(Source(i, 5)) >= CDate(Standard(2, 3)) Then
               Source(i, 7) = 1
            End If
        End If
        If IsNumeric(Source(i, 6)) Then
            If Source(i, 6) = 0 Then
               Source(i, 9) = Source(i, 9) + 1
            ElseIf CDate(Source(i, 6)) < CDate(Standard(2, 4)) Then
               Source(i, 8) = 1
            End If
        End If
     ElseIf InStr(Source(i, 3), Standard(3, 1)) > 0 Then
        If IsNumeric(Source(i, 5)) Then
            If Source(i, 5) = 0 Then
               Source(i, 12) = Source(i, 12) + 1
            ElseIf CDate(Source(i, 5)) >= CDate(Standard(3, 3)) Then
               Source(i, 10) = 1
            End If
        End If
        If IsNumeric(Source(i, 6)) Then
            If Source(i, 6) = 0 Then
               Source(i, 12) = Source(i, 12) + 1
            ElseIf CDate(Source(i, 6)) < CDate(Standard(3, 4)) Then
               Source(i, 11) = 1
            End If
        End If
     ElseIf InStr(Source(i, 3), Standard(4, 1)) > 0 Then
        If IsNumeric(Source(i, 5)) Then
            If Source(i, 5) = 0 Then
               Source(i, 9) = Source(i, 9) + 1
            ElseIf CDate(Source(i, 5)) >= CDate(Standard(4, 3)) Then
               Source(i, 7) = 1
            End If
        End If
        If IsNumeric(Source(i, 6)) Then
            If Source(i, 6) = 0 Then
               Source(i, 9) = Source(i, 9) + 1
            ElseIf CDate(Source(i, 6)) < CDate(Standard(4, 4)) Then
               Source(i, 8) = 1
            End If
        End If
     ElseIf InStr(Source(i, 3), Standard(5, 1)) > 0 Then
        If IsNumeric(Source(i, 5)) Then
            If Source(i, 5) = 0 Then
               Source(i, 12) = Source(i, 12) + 1
            ElseIf CDate(Source(i, 5)) >= CDate(Standard(5, 3)) Then
               Source(i, 10) = 1
            End If
        End If
' 不考核班主任下午签退，不可删除
'        If IsNumeric(Source(i, 6)) Then
'           If Source(i, 6) = 0 Then
'              Source(i, 12) = Source(i, 12) + 1
'           ElseIf CDate(Source(i, 6)) < CDate(Standard(5, 4)) Then
'              Source(i, 11) = 1
'           End If
'        End If
     ElseIf InStr(Source(i, 3), Standard(6, 1)) > 0 Then
'不考核班主任晚上签到，不可删除
'        If IsNumeric(Source(i, 5)) Then
'           If Source(i, 5) = 0 Then
'            Source(i, 15) = Source(i, 12) + 1
'            ElseIf CDate(Source(i, 5)) >= CDate(Standard(6, 3)) Then
'            Source(i, 13) = 1
'         End If
'        End If
        If IsNumeric(Source(i, 6)) Then
            If Source(i, 6) = 0 Then
               Source(i, 15) = Source(i, 12) + 1
            ElseIf CDate(Source(i, 6)) < CDate(Standard(6, 4)) Then
               Source(i, 14) = 1
            End If
        End If
     End If
    Next
' 生成教师数组
    ReDim Teacher(1 To SRmax, 1 To 13)
    k = 1
    For j = 1 To 12
        Teacher(1, j) = Source(1, j)
    Next
        Teacher(1, 13) = "总数"
    For i = 1 To SRmax
        If InStr(Source(i, 3), NameTeacher) > 0 Then
            k = k + 1
            For j = 1 To 12
                Teacher(k, j) = Source(i, j)
            Next
        End If
    Next
' 生成班主任数组
    ReDim HeadMaster(1 To SRmax, 1 To 16)
    k = 1
    For j = 1 To 15
        HeadMaster(1, j) = Source(1, j)
    Next
        HeadMaster(1, 16) = "总数"
    For i = 1 To SRmax
        If InStr(Source(i, 3), NameHeadMaster) > 0 Then
            k = k + 1
            For j = 1 To 15
                HeadMaster(k, j) = Source(i, j)
            Next
        End If
    Next
' 获取统计时间范围
    BeginDate = Source(2, 2)
    EndDate = Source(2, 2)
    For i = 2 To SRmax
        If Source(i, 2) > 0 Then
            If BeginDate > Source(i, 2) Then
                BeginDate = Source(i, 2)
            End If
            If EndDate < Source(i, 2) Then
                EndDate = Source(i, 2)
            End If
        End If
    Next
'  设置输出文件夹
    OutFileFix = "（" & Format(BeginDate, "yyyy" & "年" & "m" & "月" & "d" & "日") & "-" & Format(EndDate, "yyyy" & "年" & "m" & "月" & "d" & "日") & "）"
    OutFolder = OutPath & "\" & Format(EndDate, "m" & "月" & "d" & "日") & "正式上报"
    If SFO.FolderExists(OutFolder) = False Then
       MkDir OutFolder
    End If
' 汇总Teacher
    Call 汇总统计处理(Teacher)
' 汇总HeadMaster
    Call 汇总统计处理(HeadMaster)
    Application.DisplayAlerts = False
    Workbooks.Close                                                 '关闭所有工作薄
    Application.DisplayAlerts = True
    Application.Quit                                                '退出Excel
    Shell "explorer.exe " & OutFolder, vbNormalFocus                '打开目标文件夹
End Sub
Sub 汇总统计处理(THDATA)
    Dim i, j, k, l, m, n, o, p, q As Integer
    Dim SubSRmax As Integer
    Dim SubSCmax As Integer
    Dim DataRmax As Integer
    Dim SubSource As Variant
    Dim Abnormal As Variant
    SubSource = THDATA
    SubSRmax = UBound(SubSource, 1)
    SubSCmax = UBound(SubSource, 2)
    DataRmax = 0
    For i = 1 To SubSRmax
         If SubSource(i, 1) > 0 Then
           DataRmax = DataRmax + 1
         End If
    Next
    If InStr(SubSource(2, 3), NameHeadMaster) > 0 Then
      n = 15
    Else
      n = 12
    End If
    o = n + 1
    ReDim Abnormal(1 To DataRmax, 1 To o)
    j = 2
    For m = 1 To o
        Abnormal(1, m) = SubSource(1, m)
    Next
    p = 1
    For i = 3 To DataRmax
       If InStr(SubSource(i - 1, 1), SubSource(i, 1)) > 0 Then
            If i = DataRmax Then
                GoTo OTOGi
            End If
       Else
OTOGi:       If i = DataRmax Then
                l = i
            Else
                l = i - 1
            End If
            For k = j To l - 1
                For m = 7 To n
                    SubSource(l, m) = SubSource(l, m) + SubSource(k, m)
                    SubSource(k, m) = ""
                Next
            Next
            For m = 7 To n
                SubSource(l, o) = SubSource(l, o) + SubSource(l, m)
            Next
            If SubSource(l, o) > 0 Then                                             '写出异常的记录
              For k = j To l
                 p = p + 1
                 For m = 1 To o
                   Abnormal(p, m) = SubSource(k, m)
                 Next
              Next
            End If
            j = i
       End If
    Next
' 处理异常班主任考勤弹性规则 adding,Morning 等记录班主任到教师正常上班的次数，是可以选择性去除的次数2021/5/13
If InStr(Abnormal(2, 3), NameHeadMaster) > 0 Then
    k = 1
    j = 0
    Morning = 0
    Afternoon = 0
    MorningX = 0
    AfternoonX = 0
    Evening = 0
    EveningX = 0
    For i = 1 To UBound(Abnormal, 1)
        If Abnormal(i, 1) > 0 Then
            j = j + 1
        End If
    Next
    For i = 3 To j
        If InStr(Abnormal(i - 1, 1), Abnormal(i, 1)) > 0 Then
            k = k + 1
            If i = j Then
                GoTo OTOGii
            End If
        Else
OTOGii:     If i < j Then
                m = i - 1
            Else
                m = j
            End If
            For l = i - k To m
                If InStr(Abnormal(l, 4), "六") + InStr(Abnormal(l, 4), "日") > 0 Then
                    If InStr(Abnormal(l, 3), "上午") > 0 Then
                       If IsNumeric(Abnormal(l, 5)) Then
                         If CDate(Standard(4, 3)) <= CDate(Abnormal(l, 5)) Then
                            If CDate(Abnormal(l, 5)) < CDate(Standard(2, 3)) Then
                                     MorningX = MorningX + 1
                            End If
                         End If
                       End If
                    End If
                    If InStr(Abnormal(l, 3), "下午") > 0 Then
                        If IsNumeric(Abnormal(l, 5)) Then
                            If CDate(Standard(5, 3)) < CDate(Abnormal(l, 5)) Then
                                If CDate(Abnormal(l, 5)) < CDate(Standard(3, 3)) Then
                                    AfternoonX = AfternoonX + 1
                                End If
                            End If
                        End If
                    End If
                    If InStr(Abnormal(l, 3), "晚上") > 0 Then
                        If IsNumeric(Abnormal(l, 6)) Then
                            If CDate(Abnormal(l, 6)) < CDate(Standard(6, 4)) Then
                                EveningX = EveningX + 1
                            End If
                        End If
                    End If
                Else
                    If InStr(Abnormal(l, 3), "上午") > 0 Then
                       If IsNumeric(Abnormal(l, 5)) Then
                         If CDate(Standard(4, 3)) <= CDate(Abnormal(l, 5)) Then
                            If CDate(Abnormal(l, 5)) < CDate(Standard(2, 3)) Then
                                     Morning = Morning + 1
                            End If
                         End If
                       End If
                    End If
                    If InStr(Abnormal(l, 3), "下午") > 0 Then
                        If IsNumeric(Abnormal(l, 5)) Then
                            If CDate(Standard(5, 3)) < CDate(Abnormal(l, 5)) Then
                                If CDate(Abnormal(l, 5)) < CDate(Standard(3, 3)) Then
                                    Afternoon = Afternoon + 1
                                End If
                            End If
                        End If
                    End If
                    If InStr(Abnormal(l, 3), "晚上") > 0 Then
                        If IsNumeric(Abnormal(l, 6)) Then
                            If 0 < CDate(Abnormal(l, 6)) Then
                                If CDate(Abnormal(l, 6)) < CDate(Standard(6, 4)) Then
                                    Evening = Evening + 1
                                End If
                            End If
                        End If
                    End If
                End If
            Next
            If Morning <= 1 Then
             Abnormal(m, o) = Abnormal(m, o) - Morning
            Else
             Abnormal(m, o) = Abnormal(m, o) - 1
            End If
            If MorningX <= 1 Then
             Abnormal(m, o) = Abnormal(m, o) - MorningX
            Else
             Abnormal(m, o) = Abnormal(m, o) - 1
            End If
            If Afternoon <= 3 Then
             Abnormal(m, o) = Abnormal(m, o) - Afternoon
            Else
             Abnormal(m, o) = Abnormal(m, o) - 3
            End If
            If AfternoonX <= 1 Then
             Abnormal(m, o) = Abnormal(m, o) - AfternoonX
            Else
             Abnormal(m, o) = Abnormal(m, o) - 1
            End If
            If Evening <= 1 Then
             Abnormal(m, o) = Abnormal(m, o) - Evening
            Else
             Abnormal(m, o) = Abnormal(m, o) - 1
            End If
            If EveningX <= 1 Then
             Abnormal(m, o) = Abnormal(m, o) - EveningX
            Else
             Abnormal(m, o) = Abnormal(m, o) - 1
            End If
            k = 1
            Morning = 0
            Afternoon = 0
            MorningX = 0
            AfternoonX = 0
            Evening = 0
            EveningX = 0
        End If
    Next
End If
' 根据汇总结果生成报表,Normalswitch=1 生成异常班主任和异常教师表,Normalswitch=1 生成全部报表
    If InStr(SubSource(2, 3), NameHeadMaster) > 0 Then
        If NormalSwitch <> 0 Then
            Call OutToBook(SubSource, SubSRmax, SubSCmax, NameHeadMaster)
        End If
            Call OutToBook(Abnormal, DataRmax, o, NameHeadMasterUN)
    ElseIf InStr(SubSource(2, 3), NameTeacher) > 0 Then
         If NormalSwitch <> 0 Then
             Call OutToBook(SubSource, SubSRmax, SubSCmax, NameTeacher)
         End If
             Call OutToBook(Abnormal, DataRmax, o, NameTeacherUN)
    End If
End Sub

'''' V50考勤系统预置颜色
''
Sub 正常色()

'正常色为浅绿色底+深绿色文字
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10092441
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .Color = -16751104
        .TintAndShade = 0
    End With
End Sub
Sub 预警色()
'
' 黄色底+棕色文字
'
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10092543
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = -0.499984740745262
    End With
End Sub

Sub 违规色()
'
' 红色底+深红色文字
'
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 13408767
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .Color = -16777024
        .TintAndShade = 0
    End With
End Sub

Sub 漏签色()
'
' 深红色底+黄色字
'

    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .Color = -16711681
        .TintAndShade = 0
    End With
End Sub
Sub 备注色()
'
' 深绿色底+黄色字
'
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .Color = -16711681
        .TintAndShade = 0
    End With
End Sub
Sub 申请早退色()
'
' 天蓝色,考虑下一版更新
'
   With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
End Sub
Sub 申请晚到色()
'
' 紫色，考虑下一版更新
'
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10498160
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
''''
Sub WriteColorTo(THSource)
   Dim THColor As Variant
   Dim ColorRmax As Integer
   Dim THSRmax As Integer
   Dim THSCmax As Integer
   Dim i, j, k As Integer
   THColor = THSource
   THSRmax = UBound(THColor, 1)
   THSCmax = UBound(THColor, 2)
   ColorRmax = 0
   For i = 1 To THSRmax
        If THColor(i, 1) > 0 Then
          ColorRmax = ColorRmax + 1
        End If
   Next
''
   For i = 2 To ColorRmax
'''' 教师上午
    If InStr(THColor(i, 3), "教师上午") > 0 Then
'''''签到
      Cells(i, 5).Select
      If IsNumeric(THColor(i, 5)) Then
        If THColor(i, 5) < Standard(2, 2) Then
''''''''''''''
         If THColor(i, 5) = 0 Then
            Call 漏签色
         Else
            Call 正常色
         End If
''''''''''''''
        ElseIf THColor(i, 5) < Standard(2, 3) Then
         Call 预警色
        Else
         Call 违规色
        End If
      Else
        Call 备注色
      End If
'''''签退
      Cells(i, 6).Select
      If IsNumeric(THColor(i, 6)) Then
       If THColor(i, 6) < Standard(2, 4) Then
         If THColor(i, 6) = 0 Then
          Call 漏签色
         Else
          Call 违规色
         End If
       ElseIf THColor(i, 6) < Standard(2, 5) Then
         Call 预警色
       Else
         Call 正常色
       End If
      Else
       Call 备注色
      End If
''''教师下午
    ElseIf InStr(THColor(i, 3), "教师下午") > 0 Then
'''''签到
      Cells(i, 5).Select
      If IsNumeric(THColor(i, 5)) Then
       If THColor(i, 5) < Standard(3, 2) Then
''''''''''''''
        If THColor(i, 5) = 0 Then
         Call 漏签色
        Else
         Call 正常色
        End If
''''''''''''''
       ElseIf THColor(i, 5) < Standard(3, 3) Then
        Call 预警色
       Else
        Call 违规色
       End If
     Else
       Call 备注色
     End If
'''''签退
      Cells(i, 6).Select
      If IsNumeric(THColor(i, 6)) Then
       If THColor(i, 6) < Standard(3, 4) Then
        If THColor(i, 6) = 0 Then
            Call 漏签色
        Else
            Call 违规色
        End If
       Else
        Call 正常色
       End If
      Else
       Call 备注色
      End If
''''班主任上午
    ElseIf InStr(THColor(i, 3), "班主任上午") > 0 Then
'''''签到
      Cells(i, 5).Select
      If IsNumeric(THColor(i, 5)) Then
       If THColor(i, 5) < Standard(4, 2) Then
''''''''''''''
        If THColor(i, 5) = 0 Then
         Call 漏签色
        Else
         Call 正常色
        End If
''''''''''''''
       ElseIf THColor(i, 5) < Standard(4, 3) Then
        Call 预警色
       Else
        Call 违规色
       End If
      Else
       Call 备注色
      End If
'''''签退
      Cells(i, 6).Select
      If IsNumeric(THColor(i, 6)) Then
       If THColor(i, 6) < Standard(4, 4) Then
        If THColor(i, 6) = 0 Then
         Call 漏签色
        Else
         Call 违规色
        End If
       Else
         Call 正常色
       End If
      Else
       Call 备注色
      End If
''''班主任下午只统计签到
    ElseIf InStr(THColor(i, 3), "班主任下午") > 0 Then
      Cells(i, 5).Select
      If IsNumeric(THColor(i, 5)) Then
       If THColor(i, 5) < Standard(5, 2) Then
''''''''''''''
        If THColor(i, 5) = 0 Then
         Call 漏签色
        Else
         Call 正常色
        End If
''''''''''''''
       ElseIf THColor(i, 5) < Standard(5, 3) Then
        Call 预警色
       Else
        Call 违规色
       End If
      Else
       Call 备注色
      End If
''''班主任晚上只统计签退
    ElseIf InStr(THColor(i, 3), "班主任晚上") > 0 Then
      Cells(i, 6).Select
      If IsNumeric(THColor(i, 6)) Then
       If THColor(i, 6) < Standard(6, 4) Then
        If THColor(i, 6) = 0 Then
            Call 漏签色
        Else
            Call 违规色
        End If
       Else
        Call 正常色
       End If
     Else
       Call 备注色
     End If
    End If
   Next
' 对统计区上色
   For i = 3 To ColorRmax
     If InStr(THColor(i - 1, 1), THColor(i, 1)) > 0 Then
        If i = ColorRmax Then
            GoTo OTOGiii
        End If
     Else
OTOGiii:
        If i = ColorRmax Then
            k = i
        Else
            k = i - 1
        End If
      For j = 7 To THSCmax
          Cells(k, j).Select
          If THColor(k, j) > 0 Then
           Call 违规色
          Else
           Call 正常色
          End If
       Next
     End If
   Next
End Sub
Sub RecoverSource(RECS, RECR, RECC)
'四个参量：数组，写入时的行，写入时的列
   Dim RecSource As Variant
   Dim RecRmax As Integer
   Dim RecSRmax As Integer
   Dim RecSCmax As Integer
   Dim RecCRmax As Integer
   Dim RecCCmax As Integer
   Dim i As Integer
   Dim j As Integer
   Dim k As Integer
   Dim l As Integer
   RecSource = RECS
   RecSRmax = UBound(RecSource, 1)
   RecSCmax = UBound(RecSource, 2)
   RecCRmax = UBound(Correct, 1)
   RecCCmax = UBound(Correct, 2)
   RecRmax = 0
   For i = 1 To RecSRmax
        If RecSource(i, 1) > 0 Then
          RecRmax = RecRmax + 1
        End If
   Next
' 根据校准表将时间恢复
    For i = 2 To RecRmax
        k = 0
        For j = 1 To RecCRmax
            If InStr(RecSource(i, 1), Correct(j, 1)) > 0 Then
               If InStr(RecSource(i, 2), Correct(j, 2)) > 0 Then
                If InStr(RecSource(i, 3), Correct(j, 3)) > 0 Then
                    If InStr(RecSource(i, 4), Correct(j, 4)) > 0 Then
                        If IsNumeric(RecSource(i, 5)) Then
                            If RecSource(i, 5) = 0 Then
                            Else
                              RecSource(i, 5) = RecSource(i, 5) - Correct(j, 5)
                              RecSource(i, 5) = RecSource(i, 5) + Correct(j, 6)
                            End If
                        End If
                        If IsNumeric(RecSource(i, 6)) Then
                           If RecSource(i, 6) = 0 Then
                            Else
                              RecSource(i, 6) = RecSource(i, 6) - Correct(j, 7)
                              RecSource(i, 6) = RecSource(i, 6) + Correct(j, 8)
                            End If
                        End If
                        k = j                                   '自习调整时，在Correct中会有多条记录，则以最后一条为准，若最后一条无日期匹配，则以前面无日期的固定值为准
                    End If
                End If
              End If
            End If
        Next
        If k = 0 Then
            RecSource(i, 4) = ""
        Else
            RecSource(i, 4) = Correct(k, 9)                     ' 将第8列自习标记写入到原周次第4列标记列
        End If
        If RecSource(i, 5) = 0 Then
            If InStr(RecSource(i, 3), "班主任晚上") > 0 Then
            Else
                RecSource(i, 5) = "漏签"
            End If
        Else
            If InStr(RecSource(i, 3), "班主任晚上") > 0 Then
                RecSource(i, 5) = ""
            End If
        End If
        If RecSource(i, 6) = 0 Then
            If InStr(RecSource(i, 3), "班主任下午") > 0 Then
            Else
                RecSource(i, 6) = "漏签"
            End If
        Else
            If InStr(RecSource(i, 3), "班主任下午") > 0 Then
                RecSource(i, 6) = ""
            End If
        End If
    Next
' 将处理结果写入Sheet
    Range(Cells(1, 1), Cells(RECR, RECC)) = RecSource
' 合并统计区和姓名区
    j = 2
    For i = 3 To RecRmax
        m = 0
        For l = 2 To 2 * CGRmax
            If InStr(RecSource(i - 1, 1), Changed(l, 1)) > 0 Then
               If InStr(RecSource(i - 1, 2), Changed(l, 2)) > 0 Then
                    m = l
                End If
            End If
        Next
        If InStr(RecSource(i - 1, 1), RecSource(i, 1)) > 0 Then
          If i = RecRmax Then
             GoTo OTOGv
          End If
        Else
OTOGv:      If i = RecRmax Then
                k = i - 1
            Else
                k = i - 2
            End If
            Range(Cells(j, 1), Cells(k + 1, 1)).Select
            Call 合并选中单元格
            If m > 0 Then
                Cells(j, 7) = "调出:" & Changed(m, 1) & Chr(10) & Changed(m, 5) & Chr(10) & "调入:" & Changed(m, 6) & Chr(10) & Changed(m, 10)
            End If
            Range(Cells(j, 7), Cells(k, 9)).Select
            Call 合并选中单元格
            Range(Cells(j, 10), Cells(k, 12)).Select
            Call 合并选中单元格
            If InStr(Cells(2, 3), NameHeadMaster) > 0 Then
                Range(Cells(j, 13), Cells(k, 15)).Select
                Call 合并选中单元格
            End If
            j = i
        End If
    Next
' 合并时间区
    l = 2
    For i = 3 To RecRmax
        If InStr(RecSource(i - 1, 2), RecSource(i, 2)) > 0 Then
          If i = RecRmax Then
            GoTo OTOGiv
          End If
        Else
OTOGiv:
            If i = RecRmax Then
                k = i
            Else
                k = i - 1
            End If
            Range(Cells(l, 2), Cells(k, 2)).Select
            Call 合并选中单元格
            l = i
        End If
    Next
End Sub
'
Sub 合并选中单元格()
'
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Application.DisplayAlerts = False
    Selection.Merge
    Application.DisplayAlerts = True
End Sub
Sub 格式化()
'
    Range("A1").Select
    Selection.CurrentRegion.Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
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
    Range("A1").Select
End Sub
Sub FontSet(NF)
    With Selection.Font
        .name = NF
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
End Sub
''''创建目标文件
Sub OutToBook(InSource, InRmax, InCmax, InName)
    Dim OutBook As Workbook
    Application.SheetsInNewWorkbook = 1                                                      ' 设置1个Sheet
    Set OutBook = Workbooks.Add
    Application.DisplayAlerts = False
        OutBook.SaveAs Filename:=OutFolder & "\" & InName & OutFileFix
        Sheets(1).name = InName & OutFileFix
        Sheets(InName & OutFileFix).Columns("E:F").NumberFormatLocal = TimeFormat            '设置时间格式
        Sheets(InName & OutFileFix).Columns("B:B").NumberFormatLocal = DateFormat            '设置日期格式
        Range(Cells(1, 1), Cells(InRmax, InCmax)).Select
        Call FontSet(NameFont)                                                               '设置字体格式
        Call WriteColorTo(InSource)
        Call RecoverSource(InSource, InRmax, InCmax)
        Call 格式化
    ActiveWorkbook.Close savechanges:=True
    Application.DisplayAlerts = True
    Set OutBook = Nothing                                                                    '取消OutBook
End Sub
