Attribute VB_Name = "V68考勤处理模块"
' 通用考勤系统V68
' 作者：冯振华
' 日期：2021年4月25日--2021年7月2日
' 作用：处理标准化的输入源数组Original
' 说明：将校准表集成到到根据配置文件直接生成为Correct ,这样做的好处在于可以根据自习变化及时调整校准表，而这个集成后多用的时间几乎可以忽略不计，就方便程序来讲采用了这个方案。
' 注意：由于这个时间节点的修正，如果在周五生成请假表时有人还没有上交周五的假条，则需要将周五的假条手工加入到学校假条和总表
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
' 日志：修改班主任弹性考核规则，去除调课信息添加bug,去除追踪第6节时间bug2021/5/17
' 日志：修改特有变量为公有变量，这样在不同模块配合工作时非常方便  2021/5/23
' 日志：增强请假表的编写格式，及时自动译成v59格式                 2021/5/23
' 日志：修复换课表bug 2021/5/24
' 日志：增加学校考勤请假表的自动输入，及年度请假的总表自动加入功能
' 日志：修复学校考勤请假表bug 2021/6/2 升级版本号为V62
' 日志：修复班主任弹性考勤bug 2021/6/3 升级版本号为V63
' 日志：修复假条日期校准到一周，目前级部要求周五上交考勤，而学校要求周一上交上一周的结果，但是一般而言周五这个假条工作已经结束，问题不大 2021/6/3 升级版本号V64
' 日志：修复请假汇入总表bug，同时加入教师班主任报表的首行冻结，在电脑上查看结果更加直观     2021/6/4
' 日志：重写了换课程序，简化了调课标记，仅在备注框标注即可。同时调课更加合理，程序可读性增强 2021/6/4 升级版本号V65
' 日志：更改判断时间时IsNumberic 为IsDate ,这样可以避开源表输入时的时间转换，更加合理。尽管V65已经可用，但是本次升级提高了兼容性及效率，升级版本号为V66
' 日志：修复班主任弹性考核，由于不统计下午签退入晚上签到而导致的上色问题V66
' 日志：由于使用CDate转换了时间，而这会产生10的-17方的误差，所以修改配置文件标准对比信息以59秒分界，对于系统获得的00秒时间具备严格衡量能力V66
' 日志：增加常量最大行RowMax=10000和最大列ColMax=200,删除GroupName,优化了配置文件取得方式及GetHoliday模块2021/6/13 ，升级版本号V67
' 日志：因将配置文件升级为配置全校教师联系方式的智能文件，所以资料规模增长较大，于是将Col调整为1000以适应此种情况，并留有足够空间。同时，由于配置文件的升级造成了对按组加入
'       请假信息时判断的bug,此版修复了此bug,追加了对TeGrStep 和 TeGrZhiWu 两列列号的自动判断。升级版本号为V68 2021/7/2
'
'
'' 定义全局变量
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
    Public OriginalSheet5 As String
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
    Public Morning As Integer                                                                                   '正常上班时间违规次数
    Public Afternoon As Integer
    Public Evening As Integer
    Public MorningX As Integer                                                                                  '周末时间违规次数
    Public MorningXX As Integer
    Public AfternoonX As Integer
    Public EveningX As Integer
    Public Holiday As Variant
    Public HolidayRow As Integer
    Public DateX As Date
    Public DateY As Date
    Public DateZ As Date
    Public WeekX As Integer
    Public WeekY As Integer
    Public WeekZ As Integer
    Public GroupRow As Integer
    Public GroupColum As Integer
    Public PreHoliday As Variant
    Public TeacherGroup As Variant
    Public Source As Variant
    Public VipSource As Variant
    Public Teacher() As Variant
    Public HeadMaster() As Variant
    Public SelfStudyTable As Variant
    Public CorrectTable As Variant
    Public ReCorrectTable As Variant
    Public CorrectTime As Variant
    Public SRmax, SCmax, HRmax, HCmax, TGRmax, TGCmax, CRmax, CCmax, STRmax, STCmax, ViRmax, ViCmax, ORmax, OCmax As Integer      '通用数值
    Public SSRmax, SSCmax, RCTRmax, RCTCmax, CTRmax, CTCmax, CTERmax, CTECmax As Integer                                          '专为校准表设置
    Public SubSRmax As Integer
    Public SubSCmax As Integer
    Public DataRmax As Integer
    Public SubSource As Variant
    Public Abnormal As Variant
    Public THColor As Variant
    Public ColorRmax As Integer
    Public THSRmax As Integer
    Public THSCmax As Integer
    Public BCok As Integer
    Public PreLeave As Variant                     ' 准备，用来生成向学校上交的请假考勤表，尚未合并
    Public Leave As Variant                        ' 将Leave同一个人的信息叠加为一块，然后生成格式化后的请假考勤表
    Public LeaveName As Variant
    Public TTsize As Variant
    Public OutLeavePath As String
    Public LeaveBook As Workbook
    Public ChangedRow As Integer
    Public Const RowMax As Integer = 10000
    Public Const ColMax As Integer = 1000
    Public TeGrStep, TeGrZhiWu As Integer
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
'    ConfigFile = ConfigFolder & "\" & "考勤配置.xlsx"                                                          '开启普通，不能自动校准的配置文件
    ConfigFile = ConfigFolder & "\" & "考勤配置.xlsm"                                                           '开启增强，自动校准功能的配置文件
    ConfigSheet1 = "停止考勤"
    ConfigSheet2 = "请假表"
    OriginalSheet1 = "自习安排"                                                                                 '以下4项专为校准表设置
    OriginalSheet3 = "二次校准"
    OriginalSheet4 = "调课表"
    OriginalSheet5 = "教师分组"
    WuYi = Format(Now, "yyyy") & "/5/1"
    ShiYi = Format(Now, "yyyy") & "/10/1"
    If CDate(WuYi) < CDate(Now) < CDate(ShiYi) Then
        ConfigSheet3 = "51ST"
        OriginalSheet2 = "51CT"
    Else
        ConfigSheet3 = "10ST"
        OriginalSheet2 = "10CT"
    End If
    VipSwitch = 1                                                                                               '开启vip
    NormalSwitch = 1                                                                                            '0不输出正常报表，1输出
    StopSymbol = "*"
End Sub
'' 获得请假表
Sub GetHoliday()
    Dim GBeginDate As Date
    Dim GEndDate As Date
    HRmax = ConfigBook.Sheets(ConfigSheet2).Cells(RowMax, 1).End(xlUp).Row
    HCmax = ConfigBook.Sheets(ConfigSheet2).Cells(1, ColMax).End(xlToLeft).Column
    PreHoliday = ConfigBook.Sheets(ConfigSheet2).Range(ConfigBook.Sheets(ConfigSheet2).Cells(1, 1), ConfigBook.Sheets(ConfigSheet2).Cells(HRmax, HCmax))
    TGRmax = ConfigBook.Sheets(OriginalSheet5).Cells(RowMax, 1).End(xlUp).Row
    TGCmax = ConfigBook.Sheets(OriginalSheet5).Cells(2, ColMax).End(xlToLeft).Column
    TeacherGroup = ConfigBook.Sheets(OriginalSheet5).Range(ConfigBook.Sheets(OriginalSheet5).Cells(1, 1), ConfigBook.Sheets(OriginalSheet5).Cells(TGRmax, TGCmax))
    TeGrStep = 2: TeGrZhiWu = 2
    Do Until TeacherGroup(1, TeGrStep + 1) > 0
        TeGrStep = TeGrStep + 1
    Loop
    Do Until TeacherGroup(2, TeGrZhiWu + 1) = "职务"
        TeGrZhiWu = TeGrZhiWu + 1
    Loop
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 校准考勤时间为标准的周一到周日，因为年级需要周五提交上周五到本周四的考勤，而学校要求提交周一到周日的考勤报表              '
'                                                                                                                           '
' 注意：由于这个时间节点的修正，如果在周五生成请假表时有人还没有上交周五的假条，则需要将周五的假条手工加入到学校假条和总表  '
'                                                                                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    GEndDate = EndDate
    Do While Weekday(GEndDate, 2) < 7
        GEndDate = GEndDate + 1
    Loop
    GBeginDate = GEndDate - 6
   
''根据请假表生成上交的请假表leave
    LeaveName = "高二文理部" & Format(GBeginDate, "m" & "月" & "d" & "日") & "-" & Format(GEndDate, "m" & "月" & "d" & "日") & "考勤"
    ReDim PreLeave(1 To 6000, 1 To 7) As Variant
    PreLeave(1, 1) = Format(GBeginDate, "m" & "月" & "d" & "日") & "-" & Format(GEndDate, "m" & "月" & "d" & "日") & "考勤情况"
    PreLeave(2, 1) = "年级/科室"
    PreLeave(2, 2) = "时间"
    PreLeave(2, 3) = "姓名"
    PreLeave(2, 4) = "事由"
    k = 2
    For i = 2 To HRmax
        If IsDate(PreHoliday(i, 3)) And IsDate(PreHoliday(i, 4)) Then           '先判断为日期后再执行请假操作
            If CDate(GBeginDate) <= CDate(PreHoliday(i, 4)) And CDate(PreHoliday(i, 3)) <= CDate(GEndDate) And InStr(PreHoliday(i, 1), "*") = 0 Then
                k = k + 1
                PreLeave(k, 1) = "高二"
'' 取得请假的时间间隔
                If CDate(PreHoliday(i, 3)) <= CDate(GBeginDate) Then           '取得起始时间
                    PreLeave(k, 2) = CDate(GBeginDate)
                Else
                    PreLeave(k, 2) = CDate(PreHoliday(i, 3))
                End If
                If CDate(GEndDate) <= CDate(PreHoliday(i, 4)) Then             '取得终止时间
                    PreLeave(k, 3) = CDate(GEndDate)
                Else
                    PreLeave(k, 3) = CDate(PreHoliday(i, 4))
                End If
'' 生成预处理请假表（上交学校）
                PreLeave(k, 6) = PreHoliday(i, 1)
                If PreHoliday(i, 5) > 0 Then
                    PreLeave(k, 7) = PreHoliday(i, 5)
                    PreLeave(k, 4) = PreLeave(k, 3) - PreLeave(k, 2) + 1
                Else
                    PreLeave(k, 4) = 0.5 * (PreLeave(k, 3) - PreLeave(k, 2) + 1)
                    If PreHoliday(i, 6) > 0 Or PreHoliday(i, 7) > 0 Then
                        PreLeave(k, 5) = "上午"
                    End If
                    If PreHoliday(i, 8) > 0 Or PreHoliday(i, 9) > 0 Then
                        PreLeave(k, 5) = "下午"
                    End If
                    If PreHoliday(i, 10) > 0 Or PreHoliday(i, 11) > 0 Then
                        PreLeave(k, 5) = "晚上"
                    End If
                    If PreHoliday(i, 6) > 0 Then
                        PreLeave(k, 7) = PreHoliday(i, 6)
                    ElseIf PreHoliday(i, 7) > 0 Then
                        PreLeave(k, 7) = PreHoliday(i, 7)
                    ElseIf PreHoliday(i, 8) > 0 Then
                        PreLeave(k, 7) = PreHoliday(i, 8)
                    ElseIf PreHoliday(i, 9) > 0 Then
                        PreLeave(k, 7) = PreHoliday(i, 9)
                    ElseIf PreHoliday(i, 10) > 0 Then
                        PreLeave(k, 7) = PreHoliday(i, 10)
                    ElseIf PreHoliday(i, 11) > 0 Then
                        PreLeave(k, 7) = PreHoliday(i, 11)
                    End If
                End If
            End If
        End If
    Next
''''同一个人的请假信息如果在连续的几行中，则是同一条请假信息分开来写的，所以应当合并在一块
    ReDim Leave(1 To UBound(PreLeave, 1), 1 To 4) As Variant
    For i = 1 To 2
        For j = 1 To 4
            Leave(i, j) = PreLeave(i, j)
        Next
    Next
    l = 2
    For i = 3 To UBound(PreLeave, 1) - 1
        If PreLeave(i, 6) > 0 And InStr(PreLeave(i, 6), "组") = 0 Then
            k = i + 1
            Do While PreLeave(i, 6) = PreLeave(k, 6)
                k = k + 1
            Loop
            If k > i + 1 Then
             l = l + 1
             Leave(l, 1) = PreLeave(i, 1)
             Leave(l, 2) = Format(PreLeave(i, 2), "m" & "月" & "d" & "日") & PreLeave(i, 5) & "-" & Format(PreLeave(k - 1, 3), "m" & "月" & "d" & "日") & PreLeave(k - 1, 5)
             Leave(l, 3) = PreLeave(l, 6)
             For n = i To k - 1
                Leave(l, 4) = Leave(l, 4) + PreLeave(n, 4)
             Next
             Leave(l, 4) = PreLeave(k - 1, 7) & "共" & Leave(l, 4) & "天"
             i = k - 1
            Else
             l = l + 1
             Leave(l, 1) = PreLeave(i, 1)
             If CDate(PreLeave(i, 2)) = CDate(PreLeave(i, 3)) Then
                Leave(l, 2) = Format(PreLeave(i, 2), "m" & "月" & "d" & "日") & PreLeave(i, 5)
             Else
                Leave(l, 2) = Format(PreLeave(i, 2), "m" & "月" & "d" & "日") & "-" & Format(PreLeave(i, 3), "m" & "月" & "d" & "日")
             End If
             Leave(l, 3) = PreLeave(i, 6)
             Leave(l, 4) = PreLeave(i, 7) & "共" & PreLeave(i, 4) & "天"
            End If
       End If
  Next
'' 生成请假的校准表Holiday
    ReDim Holiday(1 To 3000, 1 To 6)
    Holiday(1, 1) = "姓名"
    Holiday(1, 2) = "时间"
    Holiday(1, 3) = "班次"
    Holiday(1, 4) = "星期"
    Holiday(1, 5) = "签到"
    Holiday(1, 6) = "签退"
    p = 1
    For i = 2 To HRmax
        If PreHoliday(i, 1) > 0 Then
    ''''' 取得教师分组和对应行数 debug
            If InStr(PreHoliday(i, 1), "组") > 0 Or InStr(PreHoliday(i, 1), "班主任") > 0 Or InStr(PreHoliday(i, 1), "督导室") > 0 Then
                For j = 1 To TGCmax Step TeGrStep
                    If TeacherGroup(1, j) = PreHoliday(i, 1) Then
                        GroupColum = j
                    End If
                Next
                GroupRow = ConfigBook.Sheets(OriginalSheet5).Cells(RowMax, GroupColum).End(xlUp).Row
            End If
    '''''
            If IsDate(PreHoliday(i, 3)) Then
            ' 取得有效时间
                If CDate(BeginDate) <= CDate(PreHoliday(i, 4)) Then
                   If CDate(BeginDate) <= CDate(PreHoliday(i, 3)) Then
                    DateX = PreHoliday(i, 3)
                   Else
                    DateX = BeginDate
                   End If
                   If CDate(EndDate) <= CDate(PreHoliday(i, 4)) Then
                    DateY = EndDate
                   Else
                    DateY = PreHoliday(i, 4)
                   End If
            ' 转换不同模式
    '''''''
            If InStr(PreHoliday(i, 1), "组") > 0 Or InStr(PreHoliday(i, 1), "班主任") > 0 Or InStr(PreHoliday(i, 1), "督导室") > 0 Then
                For a = 3 To GroupRow
                   If PreHoliday(i, 5) <> 0 Then
                   DateZ = CDate(DateX)
                    Do While CDate(DateZ) <= CDate(DateY)
                         p = p + 1
                         Holiday(p, 1) = TeacherGroup(a, GroupColum)
                         Holiday(p, 2) = DateZ
                         Holiday(p, 3) = TeacherGroup(a, GroupColum + TeGrZhiWu)
                         Holiday(p, 5) = PreHoliday(i, 5)
                         Holiday(p, 6) = PreHoliday(i, 5)
                         DateZ = DateZ + 1
                    Loop
                   Else
                        For k = 6 To 10 Step 2
                            DateZ = CDate(DateX)
                            If PreHoliday(i, k) > 0 Or PreHoliday(i, k + 1) > 0 Then
                             Do While CDate(DateZ) <= CDate(DateY)
                                  p = p + 1
                                  Holiday(p, 1) = TeacherGroup(a, GroupColum)
                                  Holiday(p, 2) = DateZ
                                  Holiday(p, 3) = TeacherGroup(a, GroupColum + TeGrZhiWu) & PreHoliday(1, k)
                                  Holiday(p, 5) = PreHoliday(i, k)
                                  Holiday(p, 6) = PreHoliday(i, k + 1)
                                  DateZ = DateZ + 1
                             Loop
                            End If
                         Next
                    End If
                Next
            Else
                   If PreHoliday(i, 5) <> 0 Then
                    Do While CDate(DateX) <= CDate(DateY)
                         p = p + 1
                         Holiday(p, 1) = PreHoliday(i, 1)
                         Holiday(p, 2) = DateX
                         Holiday(p, 3) = PreHoliday(i, 2)
                         Holiday(p, 5) = PreHoliday(i, 5)
                         Holiday(p, 6) = PreHoliday(i, 5)
                         DateX = DateX + 1
                    Loop
                   Else
                        For k = 6 To 10 Step 2
                            DateZ = CDate(DateX)
                            If PreHoliday(i, k) > 0 Or PreHoliday(i, k + 1) > 0 Then
                             Do While CDate(DateZ) <= CDate(DateY)
                                  p = p + 1
                                  Holiday(p, 1) = PreHoliday(i, 1)
                                  Holiday(p, 2) = DateZ
                                  Holiday(p, 3) = PreHoliday(i, 2) & PreHoliday(1, k)
                                  Holiday(p, 5) = PreHoliday(i, k)
                                  Holiday(p, 6) = PreHoliday(i, k + 1)
                                  DateZ = DateZ + 1
                             Loop
                            End If
                         Next
                    End If
            End If
                    
'''''
                End If
            Else
' 取得星期的起始
                For j = 1 To 7
                 If PreHoliday(i, 3) = Choose(j, "一", "二", "三", "四", "五", "六", "日") Then
                    WeekX = j
                 End If
                Next
                For j = 1 To 7
                 If PreHoliday(i, 4) = Choose(j, "一", "二", "三", "四", "五", "六", "日") Then
                    WeekY = j
                 End If
                Next
    ' 逐条加入到结果中
''''
                If InStr(PreHoliday(i, 1), "组") > 0 Or InStr(PreHoliday(i, 1), "班主任") > 0 Or InStr(PreHoliday(i, 1), "督导室") > 0 Then
                    For a = 3 To GroupRow
                        For k = 6 To 10 Step 2
                            WeekZ = WeekX
                            If PreHoliday(i, k) > 0 Or PreHoliday(i, k + 1) > 0 Then
                                Do While WeekZ <= WeekY
                                    p = p + 1
                                    Holiday(p, 1) = TeacherGroup(a, GroupColum)
                                    Holiday(p, 3) = TeacherGroup(a, GroupColum + TeGrZhiWu) & PreHoliday(1, k)
                                    Select Case WeekZ
                                        Case Is = 1
                                            Holiday(p, 4) = "一"
                                        Case Is = 2
                                            Holiday(p, 4) = "二"
                                        Case Is = 3
                                            Holiday(p, 4) = "三"
                                        Case Is = 4
                                            Holiday(p, 4) = "四"
                                        Case Is = 5
                                            Holiday(p, 4) = "五"
                                        Case Is = 6
                                            Holiday(p, 4) = "六"
                                        Case Is = 7
                                            Holiday(p, 4) = "日"
                                    End Select
                                    Holiday(p, 5) = PreHoliday(i, k)
                                    Holiday(p, 6) = PreHoliday(i, k + 1)
                                    WeekZ = WeekZ + 1
                                Loop
                            End If
                         Next
                     Next
                Else
                     For k = 6 To 10 Step 2
                       WeekZ = WeekX
                       If PreHoliday(i, k) > 0 Or PreHoliday(i, k + 1) > 0 Then
                           Do While WeekZ <= WeekY
                               p = p + 1
                               Holiday(p, 1) = PreHoliday(i, 1)
                               Holiday(p, 3) = PreHoliday(i, 2) & PreHoliday(1, k)
                               Select Case WeekZ
                                   Case Is = 1
                                       Holiday(p, 4) = "一"
                                   Case Is = 2
                                       Holiday(p, 4) = "二"
                                   Case Is = 3
                                       Holiday(p, 4) = "三"
                                   Case Is = 4
                                       Holiday(p, 4) = "四"
                                   Case Is = 5
                                       Holiday(p, 4) = "五"
                                   Case Is = 6
                                       Holiday(p, 4) = "六"
                                   Case Is = 7
                                       Holiday(p, 4) = "日"
                               End Select
                               Holiday(p, 5) = PreHoliday(i, k)
                               Holiday(p, 6) = PreHoliday(i, k + 1)
                               WeekZ = WeekZ + 1
                           Loop
                       End If
                    Next
                End If
    ''''
            End If
        End If
    Next
    HRmax = p                                                                                                   '获得Holiday的非空行数和列数
    HCmax = UBound(Holiday, 2)
End Sub
Sub 标准通用考勤(Original)
    Dim i, j, k, l, m, n, o, p As Integer
    Dim SelfStudyTemp As Variant
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
    ViRmax = ConfigBook.Sheets(ConfigSheet1).Cells(RowMax, 1).End(xlUp).Row
    ViCmax = ConfigBook.Sheets(ConfigSheet1).Cells(1, ColMax).End(xlToLeft).Column
    VipSource = ConfigBook.Sheets(ConfigSheet1).Range(ConfigBook.Sheets(ConfigSheet1).Cells(1, 1), ConfigBook.Sheets(ConfigSheet1).Cells(ViRmax, ViCmax))

'' 获得统计时间范围
    BeginDate = CDate(Original(2, 2))
    EndDate = CDate(Original(2, 2))
    For i = 2 To UBound(Original, 1)
        If Original(i, 2) > 0 Then
            If CDate(BeginDate) > CDate(Original(i, 2)) Then
                BeginDate = CDate(Original(i, 2))
            End If
            If CDate(EndDate) < CDate(Original(i, 2)) Then
                EndDate = CDate(Original(i, 2))
            End If
        End If
    Next
'' 获得请假表
    Call GetHoliday
''
    STRmax = ConfigBook.Sheets(ConfigSheet3).Cells(RowMax, 1).End(xlUp).Row
    STCmax = ConfigBook.Sheets(ConfigSheet3).Cells(1, ColMax).End(xlToLeft).Column
    Standard = ConfigBook.Sheets(ConfigSheet3).Range(ConfigBook.Sheets(ConfigSheet3).Cells(1, 1), ConfigBook.Sheets(ConfigSheet3).Cells(STRmax, STCmax))
''                                                                                                          '以下4项专为校准表设置
    SSRmax = ConfigBook.Sheets(OriginalSheet1).Cells(RowMax, 1).End(xlUp).Row
    SSCmax = ConfigBook.Sheets(OriginalSheet1).Cells(1, ColMax).End(xlToLeft).Column
    SelfStudyTable = ConfigBook.Sheets(OriginalSheet1).Range(ConfigBook.Sheets(OriginalSheet1).Cells(1, 1), ConfigBook.Sheets(OriginalSheet1).Cells(SSRmax, SSCmax))
''
    CTERmax = ConfigBook.Sheets(OriginalSheet2).Cells(RowMax, 1).End(xlUp).Row
    CTECmax = ConfigBook.Sheets(OriginalSheet2).Cells(1, ColMax).End(xlToLeft).Column
    CorrectTime = ConfigBook.Sheets(OriginalSheet2).Range(ConfigBook.Sheets(OriginalSheet2).Cells(1, 1), ConfigBook.Sheets(OriginalSheet2).Cells(CTERmax, CTECmax))
''
    RCTRmax = ConfigBook.Sheets(OriginalSheet3).Cells(RowMax, 1).End(xlUp).Row
    RCTCmax = ConfigBook.Sheets(OriginalSheet3).Cells(1, ColMax).End(xlToLeft).Column
    ReCorrectTable = ConfigBook.Sheets(OriginalSheet3).Range(ConfigBook.Sheets(OriginalSheet3).Cells(1, 1), ConfigBook.Sheets(OriginalSheet3).Cells(RCTRmax, RCTCmax))
''
    CGRmax = ConfigBook.Sheets(OriginalSheet4).Cells(RowMax, 1).End(xlUp).Row                               '调入换课表
    CGCmax = ConfigBook.Sheets(OriginalSheet4).Cells(1, ColMax).End(xlToLeft).Column
    Change = ConfigBook.Sheets(OriginalSheet4).Range(ConfigBook.Sheets(OriginalSheet4).Cells(1, 1), ConfigBook.Sheets(OriginalSheet4).Cells(CGRmax, CGCmax))
''
    ReDim CorrectTable(1 To 6000, 1 To 9) As Variant
    CTRmax = UBound(CorrectTable, 1)
    CTCmax = UBound(CorrectTable, 2)
''''''''''''''''''
' 校准表根据自习表生成，所以调课会影响到自习，因此需首先校准自习表
    Dim SSTB As Variant
    Dim SSTC As Variant
    Dim SSTD As Variant
    Dim SSTE As Variant
    SSTB = "B"
    SSTC = "C"
    SSTD = "D"
    SSTE = "E"
    For i = 2 To CGRmax                                                                                 '调课的二人都调出原来的自习，如果没有自习，则执行结果无效
        For l = 1 To 6 Step 5
            If InStr(Change(i, l + 4), "B") > 0 Or InStr(Change(i, l + 4), "C") > 0 Or InStr(Change(i, l + 4), "D") > 0 Or InStr(Change(i, l + 4), "E") > 0 Then
                If CDate(BeginDate) <= CDate(Change(i, l + 1)) And CDate(Change(i, l + 1)) <= CDate(EndDate) Then
                  For j = 2 To SSRmax
                      If Change(i, l) = SelfStudyTable(j, 1) Then
                          SelfStudyTemp = ""
                          k = Weekday(Change(i, l + 1), 2) + 2
                          If InStr(SelfStudyTable(j, k), Change(i, l + 4)) > 0 Then
                              SelfStudyTemp = SelfStudyTable(j, k)
                              SelfStudyTable(j, k) = ""
                              Select Case Change(i, l + 4)                                              'A并不在自习编码之列，所以使用A排除要去除的自习
                                   Case Is = "B"
                                      SSTB = "A"
                                   Case Is = "C"
                                      SSTC = "A"
                                   Case Is = "D"
                                      SSTD = "A"
                                   Case Is = "E"
                                      SSTE = "A"
                                   Case Else
                              End Select
                              If InStr(SelfStudyTemp, SSTB) > 0 Then
                                  SelfStudyTable(j, k) = SelfStudyTable(j, k) & SSTB
                              End If
                              If InStr(SelfStudyTemp, SSTC) > 0 Then
                                  SelfStudyTable(j, k) = SelfStudyTable(j, k) & SSTC
                              End If
                              If InStr(SelfStudyTemp, SSTD) > 0 Then
                                  SelfStudyTable(j, k) = SelfStudyTable(j, k) & SSTD
                              End If
                              If InStr(SelfStudyTemp, SSTE) > 0 Then
                                  SelfStudyTable(j, k) = SelfStudyTable(j, k) & SSTE
                              End If
                              SSTB = "B"
                              SSTC = "C"
                              SSTD = "D"
                              SSTE = "E"
                          End If
                      End If
                  Next
                End If
            End If
        Next
    Next
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'调课分析：如果同一个人自己的一个自习和另一个自习对换，这点不用进行调课，所以不必追加调课记录 '
'          如果同一个人自己的一节自习和另一节公共自习对换，则不必调入结果，只需调出对应自习即 '
'          可。所以，在下面的调入对方自习中，不考虑同一个人的换课情况。                       '
'                                                                                             '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    For i = 2 To CGRmax                                                                                         ' 调入对方的课,如果调入自习与已有自习重复，则提示冲突，而调课失败
        If InStr(Change(i, 1), Change(i, 6)) = 0 Then                                                           ' 调课的双方不可以是同一个人
            If InStr(Change(i, 10), "B") > 0 Or InStr(Change(i, 10), "C") > 0 Or InStr(Change(i, 10), "D") > 0 Or InStr(Change(i, 10), "E") > 0 Then
                If CDate(BeginDate) <= CDate(Change(i, 7)) And CDate(Change(i, 7)) <= CDate(EndDate) Then
                    For j = 2 To SSRmax
                        If Change(i, 1) = SelfStudyTable(j, 1) Then
                            k = Weekday(Change(i, 7), 2) + 2
                            If InStr(SelfStudyTable(j, k), Change(i, 10)) > 0 Then
                                MsgBox Change(i, 1) & "与" & Change(i, 6) & "调课因冲突失败"
                            Else
                             SelfStudyTable(j, k) = SelfStudyTable(j, k) & Change(i, 10)
                            End If
                        End If
                    Next
                End If
            End If
            If InStr(Change(i, 5), "B") > 0 Or InStr(Change(i, 5), "C") > 0 Or InStr(Change(i, 5), "D") > 0 Or InStr(Change(i, 5), "E") > 0 Then
                If CDate(BeginDate) <= CDate(Change(i, 2)) And CDate(Change(i, 2)) <= CDate(EndDate) Then
                    For j = 2 To SSRmax
                        If Change(i, 6) = SelfStudyTable(j, 1) Then
                            k = Weekday(Change(i, 2), 2) + 2
                            If InStr(SelfStudyTable(j, k), Change(i, 5)) > 0 Then
                                MsgBox Change(i, 6) & "与" & Change(i, 1) & "调课因冲突失败"
                            Else
                             SelfStudyTable(j, k) = SelfStudyTable(j, k) & Change(i, 5)
                            End If
                        End If
                    Next
                End If
            End If
        End If
    Next
' 调课情况需要在最后的结果中记录，所以用Changed 来记录，提前生成，留到后期写入表格后处理
o = 2 * CGRmax
ReDim Changed(1 To o, 1 To 10) As Variant
k = 0
For i = 2 To CGRmax
    If CDate(BeginDate) <= CDate(Change(i, 2)) Or CDate(BeginDate) <= CDate(Change(i, 7)) Then                  ' 只有换课表中时间超过统计时间时才写入Changed
        k = k + 1
        For j = 1 To 4
          Changed(k, j) = Change(i, j)
        Next
        Changed(k, 5) = Format(Change(i, 2), DateFormat)
        If InStr(Change(i, 5), "B") + InStr(Change(i, 5), "C") + InStr(Change(i, 5), "D") + InStr(Change(i, 5), "E") > 0 Then
              If InStr(Change(i, 5), "B") > 0 Then
                Changed(k, 5) = Changed(k, 5) & " 第1节"
              End If
              If InStr(Change(i, 5), "C") > 0 Then
                 Changed(k, 5) = Changed(k, 5) & " 第5节"
              End If
              If InStr(Change(i, 5), "E") > 0 Then
                 Changed(k, 5) = Changed(k, 5) & " 第6节"
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
        If InStr(Change(i, 10), "B") + InStr(Change(i, 10), "C") + InStr(Change(i, 10), "D") + InStr(Change(i, 10), "E") > 0 Then
              If InStr(Change(i, 10), "B") > 0 Then
                Changed(k, 10) = Changed(k, 10) & " 第1节"
              End If
              If InStr(Change(i, 10), "C") > 0 Then
                 Changed(k, 10) = Changed(k, 10) & " 第5节"
              End If
              If InStr(Change(i, 10), "E") > 0 Then
                 Changed(k, 10) = Changed(k, 10) & " 第6节"
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
    End If
Next
ChangedRow = k

'''''''''''''''''''''''''''''''''''''''''''''''''''''
'换课记录：以上生成Changed表，为最后注明换课做准备  '
'                                                   '
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'' 将自习表转化为校准表，由于前面已经做了自习校准，所以之后不必再追加校准表多余记录，可以提高效率
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
                     CorrectTable(k, 5) = CorrectTime(6, 2)
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
                  CorrectTable(k, 5) = CorrectTime(6, 2)
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
                If CDate(CorrectTable(i, 2)) = CDate(ReCorrectTable(j, 2)) Or ReCorrectTable(j, 2) = 0 Then
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
    l = p
'' 获得Correct(校准表)，这一步是只取非空记录，可以减小统计时的对比数量
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
    For i = 2 To ORmax
        l = 0
        If Original(i, 3) <> 0 Then
' 将第i行和VipSource 列表比对，以l记录比对数，如果比对数为1，则表明有匹配项产生
            For j = 2 To ViRmax
                 If InStr(Original(i, 1), VipSource(j, 1)) > 0 Then
                     If CDate(Original(i, 2)) = CDate(VipSource(j, 2)) Or VipSource(j, 2) = 0 Then
                        If InStr(Original(i, 3), VipSource(j, 3)) > 0 Then
                           If InStr(Original(i, 4), VipSource(j, 4)) > 0 Then
                            l = 1
                           End If
                       End If
                      End If
                 ElseIf InStr(VipSource(j, 1), StopSymbol) > 0 Then
                     If CDate(Original(i, 2)) = CDate(VipSource(j, 2)) Or VipSource(j, 2) = 0 Then
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
    Source(1, 1) = "姓名"
    Source(1, 2) = "日期"
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
            If CDate(Source(j, 2)) = CDate(Holiday(i, 2)) Or Holiday(i, 2) = 0 Then
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
            If CDate(Source(j, 2)) = CDate(Holiday(i, 2)) Or Holiday(i, 2) = 0 Then
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
              If CDate(Source(i, 2)) = CDate(Correct(j, 2)) Or Correct(j, 2) = 0 Then
                If InStr(Source(i, 3), Correct(j, 3)) > 0 Then
                    If InStr(Source(i, 4), Correct(j, 4)) > 0 Then
                        If IsDate(Source(i, 5)) Then
                            Source(i, 5) = Source(i, 5) + Correct(j, 5)
                            Source(i, 5) = Source(i, 5) - Correct(j, 6)
                        End If
                        If IsDate(Source(i, 6)) Then
                            Source(i, 6) = Source(i, 6) + Correct(j, 7)
                            Source(i, 6) = Source(i, 6) - Correct(j, 8)
                        End If
                    End If
                End If
              End If
            End If
        Next
    Next
'据校准后的Source生成统计数据
    For i = 2 To SRmax
     If InStr(Source(i, 3), Standard(2, 1)) > 0 Then                        '教师上午
        If Source(i, 5) = 0 Then
               Source(i, 9) = Source(i, 9) + 1
        ElseIf IsDate(Source(i, 5)) Then
            If CDate(Source(i, 5)) >= CDate(Standard(2, 3)) Then
               Source(i, 7) = 1
            End If
        End If
        If Source(i, 6) = 0 Then
                Source(i, 9) = Source(i, 9) + 1
        ElseIf IsDate(Source(i, 6)) Then
            If CDate(Source(i, 6)) < CDate(Standard(2, 4)) Then
               Source(i, 8) = 1
            End If
        End If
     ElseIf InStr(Source(i, 3), Standard(3, 1)) > 0 Then                    '教师下午
        If Source(i, 5) = 0 Then
               Source(i, 12) = Source(i, 12) + 1
        ElseIf IsDate(Source(i, 5)) Then
            If CDate(Source(i, 5)) >= CDate(Standard(3, 3)) Then
               Source(i, 10) = 1
            End If
        End If
        If Source(i, 6) = 0 Then
           Source(i, 12) = Source(i, 12) + 1
        ElseIf IsDate(Source(i, 6)) Then
            If CDate(Source(i, 6)) < CDate(Standard(3, 4)) Then
               Source(i, 11) = 1
            End If
        End If
     ElseIf InStr(Source(i, 3), Standard(4, 1)) > 0 Then                    '班主任上午
        If Source(i, 5) = 0 Then
               Source(i, 9) = Source(i, 9) + 1
        ElseIf IsDate(Source(i, 5)) Then
            If CDate(Source(i, 5)) >= CDate(CDate(Standard(4, 3))) Then
               Source(i, 7) = 1
            End If
        End If
        If Source(i, 6) = 0 Then
               Source(i, 9) = Source(i, 9) + 1
        ElseIf IsDate(Source(i, 6)) Then
            If CDate(Source(i, 6)) < CDate(Standard(4, 4)) Then
               Source(i, 8) = 1
            End If
        End If
     ElseIf InStr(Source(i, 3), Standard(5, 1)) > 0 Then                   '班主任下午
        If Source(i, 5) = 0 Then
           Source(i, 12) = Source(i, 12) + 1
        ElseIf IsDate(Source(i, 5)) Then
            If CDate(Source(i, 5)) >= CDate(Standard(5, 3)) Then
               Source(i, 10) = 1
            End If
        End If
' 不考核班主任下午签退，不可删除
'        If Source(i, 6) = 0 Then
'              Source(i, 12) = Source(i, 12) + 1
'        ElseIf Isdate(Source(i, 6)) Then
'           If CDate(Source(i, 6)) < CDate(Standard(5, 4)) Then
'              Source(i, 11) = 1
'           End If
'        End If
     ElseIf InStr(Source(i, 3), Standard(6, 1)) > 0 Then                   '班主任晚上
'不考核班主任晚上签到，不可删除
'        If Source(i, 5) = 0 Then
'            Source(i, 15) = Source(i, 12) + 1
'        ElseIf Isdate(Source(i, 5)) Then
'         If CDate(Source(i, 5)) >= CDate(Standard(6, 3)) Then
'            Source(i, 13) = 1
'         End If
'        End If
        If Source(i, 6) = 0 Then
               Source(i, 15) = Source(i, 12) + 1
        ElseIf IsDate(Source(i, 6)) Then
            If CDate(Source(i, 6)) < CDate(Standard(6, 4)) Then
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
'  设置输出文件夹
    OutFileFix = "（" & Format(BeginDate, "yyyy" & "年" & "m" & "月" & "d" & "日") & "-" & Format(EndDate, "yyyy" & "年" & "m" & "月" & "d" & "日") & "）"
    OutFolder = OutPath & "\" & Format(EndDate, "m" & "月" & "d" & "日") & "正式上报"
    If SFO.FolderExists(OutFolder) = False Then
       MkDir OutFolder
    Else
        i = 1                                                         ' 如果正式上报生成，则以编号的方式新增文件夹
        Do While SFO.FolderExists(OutFolder & i) = True
            i = i + 1
        Loop
        OutFolder = OutFolder & i
        MkDir OutFolder
    End If
' 汇总Teacher
    Call 汇总统计处理(Teacher)
' 汇总HeadMaster
    Call 汇总统计处理(HeadMaster)
' 导出请假考勤周表
    Call OutToLeave(Leave, UBound(Leave, 1), UBound(Leave, 2), OutFolder, LeaveName)
' 请假考勤情况汇入总请假表
    Call AddToTotalLeave
    Application.DisplayAlerts = False
    Workbooks.Close                                                 '关闭所有工作薄
    Application.DisplayAlerts = True
    Application.Quit                                                '退出Excel
    Shell "explorer.exe " & OutFolder, vbNormalFocus                '打开目标文件夹
End Sub
Sub 汇总统计处理(THDATA)
    Dim i, j, k, l, m, n, o, p, q, r, s As Integer
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
    
' 处理异常班主任考勤弹性规则 ,Morning 等记录班主任到教师正常上班的次数，是可以选择性去除的次数2021/5/13

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'开始：此弹性规则有可能随年级改变，所以以双线标出此部分，必要时修改此部分代码 '
'                                                                             '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If InStr(Abnormal(2, 3), NameHeadMaster) > 0 Then
    k = 1
    j = 0
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
OTOGii:
            Morning = 0
            Afternoon = 0
            MorningX = 0
            MorningXX = 0
            AfternoonX = 0
            Evening = 0
            EveningX = 0
            If i < j Then
                m = i - 1
            Else
                m = j
            End If
            For l = i - k To m
                If InStr(Abnormal(l, 4), "六") + InStr(Abnormal(l, 4), "日") > 0 Then                   '处理周六和周日的情况
                    If InStr(Abnormal(l, 3), "上午") > 0 Then
                        If Abnormal(l, 5) = 0 Then
                          MorningX = MorningX + 1
                        ElseIf IsDate(Abnormal(l, 5)) Then
                          If CDate(Standard(4, 3)) <= CDate(Abnormal(l, 5)) Then
                            MorningX = MorningX + 1
                          End If
                        End If
                        If Abnormal(l, 6) = 0 Then
                              MorningXX = MorningXX + 1
                        End If
                    End If
                    If InStr(Abnormal(l, 3), "下午") > 0 Then
                        If CDate(Abnormal(l, 5)) = 0 Then
                            AfternoonX = AfternoonX + 1
                        ElseIf IsDate(Abnormal(l, 5)) Then
                            If CDate(Standard(5, 3)) <= CDate(Abnormal(l, 5)) Then
                                AfternoonX = AfternoonX + 1
                            End If
                        End If
                    End If
                    If InStr(Abnormal(l, 3), "晚上") > 0 Then                                           '在晚上签退为空时再去检测下午签退和晚上签到
                        If Abnormal(l, 6) = 0 Then
 '                           If IsDate(Abnormal(l - 1, 6))  Then                                        '对于周未如果要求晚上必有一次签到，则取消注释
 '                            EveningX = EveningX + 1
 '                           ElseIf IsDate(Abnormal(l, 5))  Then
 '                            EveningX = EveningX + 1
 '                           End If
                            EveningX = EveningX + 1                                                     '对于周未如果要求晚上必有一次签到，则注释掉此行
                        ElseIf IsDate(Abnormal(l, 6)) Then
                            If CDate(Abnormal(l, 6)) < CDate(Standard(6, 4)) Then
                                EveningX = EveningX + 1
                            End If
                        End If
                     End If
                Else                                                                                   '处理周一到周五的情况
                    If InStr(Abnormal(l, 3), "上午") > 0 Then                                          '处理上午签到
                       If IsDate(Abnormal(l, 5)) Then
                          If CDate(Standard(4, 3)) <= CDate(Abnormal(l, 5)) Then                       '上午有第1节的情况要做单独考虑
                                r = 0
                                For s = 1 To UBound(CorrectTable, 1)
                                    If InStr(Abnormal(l, 1), CorrectTable(s, 1)) > 0 And CorrectTable(s, 1) <> 0 Then
                                        If CDate(Abnormal(l, 2)) = CDate(CorrectTable(s, 2)) Or CorrectTable(s, 2) = 0 Then
                                            If InStr(Abnormal(l, 3), CorrectTable(s, 3)) > 0 Then
                                                If InStr(Abnormal(l, 4), CorrectTable(s, 4)) > 0 Then
                                                    If InStr(CorrectTable(s, 9), "第1") > 0 Then
                                                        r = s
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                Next
                             If r > 0 Then
                                Abnormal(l, 5) = Abnormal(l, 5) + CorrectTime(2, 2)
                             End If
                             If CDate(Abnormal(l, 5)) < CDate(Standard(2, 3)) Then
                                      Morning = Morning + 1
                             End If
                             If r > 0 Then
                                Abnormal(l, 5) = Abnormal(l, 5) - CorrectTime(2, 2)
                             End If
                          End If
                       End If
                    End If
                    If InStr(Abnormal(l, 3), "下午") > 0 Then                                           '处理下午签到,匹配上午自习，以确定是否享受优惠
                        If IsDate(Abnormal(l, 5)) And Abnormal(l, 5) > 0 Then
                            If CDate(Standard(5, 3)) <= CDate(Abnormal(l, 5)) Then
                                r = 0
                                For s = 1 To UBound(CorrectTable, 1)
                                    If InStr(Abnormal(l, 1), CorrectTable(s, 1)) > 0 And CorrectTable(s, 1) <> 0 Then
                                        If CDate(Abnormal(l, 2)) = CDate(CorrectTable(s, 2)) Or CorrectTable(s, 2) = 0 Then
                                            If InStr("班主任上午", CorrectTable(s, 3)) > 0 Then
                                                If InStr(Abnormal(l, 4), CorrectTable(s, 4)) > 0 Then
                                                    If InStr(CorrectTable(s, 9), "第1") > 0 Or InStr(CorrectTable(s, 9), "第5") > 0 Then
                                                        r = s
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                Next
                                If r > 0 Then   '有上午第1节或第5节的，班主任享受照顾,前提是输入源你必须按升序排列
                                    Abnormal(l, 5) = Abnormal(l, 5) - CorrectTime(4, 3)
                                End If
                                If CDate(Abnormal(l, 5)) < CDate(Standard(3, 3)) Then
                                    Afternoon = Afternoon + 1
                                End If
                                If r > 0 Then
                                    Abnormal(l, 5) = Abnormal(l, 5) + CorrectTime(4, 3)
                                End If
                            End If
                        End If
                    End If
                    If InStr(Abnormal(l, 3), "晚上") > 0 Then                                   '处理晚上签退
                        If IsDate(Abnormal(l, 6)) And Abnormal(l, 6) > 0 Then
                            If CDate(Abnormal(l, 6)) < CDate(Standard(6, 4)) Then
                                Evening = Evening + 1
                            End If
                        ElseIf Abnormal(l, 6) = 0 Then
                            If IsDate(Abnormal(l - 1, 6)) And Abnormal(l - 1, 6) > 0 Then
                             Evening = Evening + 1
                            ElseIf IsDate(Abnormal(l, 5)) And Abnormal(l, 5) > 0 Then
                             Evening = Evening + 1
                            End If
                        End If
                    End If
                End If
            Next
' 正常上班时间, 各量记录可以去除的量
            If Morning <= 1 Then
             Abnormal(m, o) = Abnormal(m, o) - Morning
            Else
             Abnormal(m, o) = Abnormal(m, o) - 1
            End If
            If Afternoon <= 3 Then
             Abnormal(m, o) = Abnormal(m, o) - Afternoon
            Else
             Abnormal(m, o) = Abnormal(m, o) - 3
            End If
            If Evening <= 1 Then
             Abnormal(m, o) = Abnormal(m, o) - Evening
            Else
             Abnormal(m, o) = Abnormal(m, o) - 1
            End If
' 周六周日，各量记录了违规数
            Abnormal(m, o) = Abnormal(m, o) - MorningXX                     '去除对上午漏签的统计数
            If MorningX <> 0 Then                                           '周末上午签到最多2次违规
              Abnormal(m, o) = Abnormal(m, o) - 1
            End If
            If AfternoonX <> 0 Then
             Abnormal(m, o) = Abnormal(m, o) - 1
            End If
            If EveningX <> 0 Then
             Abnormal(m, o) = Abnormal(m, o) - 1
            End If
            k = 1
        End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'结束：此弹性规则有可能随年级改变，所以以双线标出此部分，必要时修改此部分代码 '
'                                                                             '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
                If THColor(i, 5) = 0 Then                                                          ' 上色修改
                   Call 漏签色
                ElseIf IsDate(THColor(i, 5)) Then
                    If CDate(THColor(i, 5)) < CDate(Standard(2, 2)) Then
                        Call 正常色
                    ElseIf CDate(THColor(i, 5)) < CDate(Standard(2, 3)) Then
                     Call 预警色
                    Else
                     Call 违规色
                    End If
                Else
                  Call 备注色
                End If
'''''签退
              Cells(i, 6).Select
              If THColor(i, 6) = 0 Then
                  Call 漏签色
              ElseIf IsDate(THColor(i, 6)) Then
                   If CDate(THColor(i, 6)) < CDate(Standard(2, 4)) Then
                      Call 违规色
                   ElseIf CDate(THColor(i, 6)) < CDate(Standard(2, 5)) Then
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
              If THColor(i, 5) = 0 Then
                 Call 漏签色
              ElseIf IsDate(THColor(i, 5)) Then
                    If CDate(THColor(i, 5)) < CDate(Standard(3, 2)) Then
                      Call 正常色
                    ElseIf CDate(THColor(i, 5)) < CDate(Standard(3, 3)) Then
                     Call 预警色
                    Else
                     Call 违规色
                    End If
             Else
              Call 备注色
             End If
'''''签退
              Cells(i, 6).Select
              If THColor(i, 6) = 0 Then
                    Call 漏签色
              ElseIf IsDate(THColor(i, 6)) Then
                    If CDate(THColor(i, 6)) < CDate(Standard(3, 4)) Then
                         Call 违规色
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
              If THColor(i, 5) = 0 Then
                 Call 漏签色
              ElseIf IsDate(THColor(i, 5)) Then
                 If CDate(THColor(i, 5)) < CDate(Standard(4, 2)) Then
                   Call 正常色
                 ElseIf CDate(THColor(i, 5)) < CDate(Standard(4, 3)) Then
                  Call 预警色
                 Else
                  Call 违规色
                 End If
              Else
               Call 备注色
              End If
'''''签退
                Cells(i, 6).Select
                If THColor(i, 6) = 0 Then
                 Call 漏签色
                ElseIf IsDate(THColor(i, 6)) Then
                    If CDate(THColor(i, 6)) < CDate(Standard(4, 4)) Then
                      Call 违规色
                    Else
                      Call 正常色
                    End If
                Else
                 Call 备注色
                End If
''''班主任下午只统计签到
    ElseIf InStr(THColor(i, 3), "班主任下午") > 0 Then
              Cells(i, 5).Select
              If THColor(i, 5) = 0 Then
                    Call 漏签色
              ElseIf IsDate(THColor(i, 5)) Then
                    If CDate(THColor(i, 5)) < CDate(Standard(5, 2)) Then
                      Call 正常色
                    ElseIf CDate(THColor(i, 5)) < CDate(Standard(5, 3)) Then
                     Call 预警色
                    Else
                     Call 违规色
                    End If
              Else
               Call 备注色
              End If
              Cells(i, 6).Select                                                '由于班主任不统计下午签退，但是有弹性考核，所以只标注颜色
              If THColor(i, 6) = 0 Then
              ElseIf IsDate(THColor(i, 6)) Then
                Call 正常色                                                     '只要签就认为正常
              Else
                Call 备注色
              End If
''''班主任晚上只统计签退
    ElseIf InStr(THColor(i, 3), "班主任晚上") > 0 Then
            Cells(i, 5).Select                                                '由于班主任不统计晚上签到，但是有弹性考核，所以只标注颜色
            If THColor(i, 5) = 0 Then
            ElseIf IsDate(THColor(i, 5)) Then
              Call 正常色                                                     '只要签就认为正常
            Else
              Call 备注色
            End If
            Cells(i, 6).Select
            If THColor(i, 6) = 0 Then
                   Call 漏签色
             ElseIf IsDate(THColor(i, 6)) Then
                If CDate(THColor(i, 6)) < CDate(Standard(6, 4)) Then
                     Call 违规色
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
   Dim i, j, k, l, m, n As Integer
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
 '              If InStr(RecSource(i, 2), Correct(j, 2)) > 0 Then
               If CDate(RecSource(i, 2)) = CDate(Correct(j, 2)) Or Correct(j, 2) = 0 Then
                If InStr(RecSource(i, 3), Correct(j, 3)) > 0 Then
                    If InStr(RecSource(i, 4), Correct(j, 4)) > 0 Then
                        If IsDate(RecSource(i, 5)) Then
                              RecSource(i, 5) = RecSource(i, 5) - Correct(j, 5)
                              RecSource(i, 5) = RecSource(i, 5) + Correct(j, 6)
                        End If
                        If IsDate(RecSource(i, 6)) Then
                              RecSource(i, 6) = RecSource(i, 6) - Correct(j, 7)
                              RecSource(i, 6) = RecSource(i, 6) + Correct(j, 8)
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
'                RecSource(i, 5) = ""                           '凡是班主任晚上签到的，不再覆盖原始签到标记,但是不参与考勤
            End If
        End If
        If RecSource(i, 6) = 0 Then
            If InStr(RecSource(i, 3), "班主任下午") > 0 Then
            Else
                RecSource(i, 6) = "漏签"
            End If
        Else
            If InStr(RecSource(i, 3), "班主任下午") > 0 Then
'               RecSource(i, 6) = ""                            '凡是班主作下午签到的，不再覆盖原始签到标记，但是也不参与考勤
            End If
        End If
    Next
' 将处理结果写入Sheet
    Range(Cells(1, 1), Cells(RECR, RECC)) = RecSource
' 合并统计区和姓名区
    j = 2
    m = 0
    For i = 3 To RecRmax
        If InStr(RecSource(i - 1, 1), RecSource(i, 1)) > 0 Then
          If i = RecRmax Then
             GoTo OTOGv
          End If
        Else
OTOGv:
            If i = RecRmax Then
                k = i - 1
            Else
                k = i - 2
            End If
            For l = 2 To ChangedRow     'adding
                For n = j To k + 1
                    If Changed(l, 1) > 0 And InStr(RecSource(n, 1), Changed(l, 1)) > 0 Then
                       If CDate(RecSource(n, 2)) = CDate(Changed(l, 2)) Then
                             m = l
                        End If
                    End If
                Next
            Next
            Range(Cells(j, 1), Cells(k + 1, 1)).Select
            Call 合并选中单元格
            If m > 0 Then
                Cells(j, 7) = "调出:" & Changed(m, 1) & Chr(10) & Changed(m, 5) & Chr(10) & "调入:" & Changed(m, 6) & Chr(10) & Changed(m, 10)
                m = 0
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
'' 冻结首行，方便在电脑上对照查看
        With ActiveWindow
            .SplitColumn = 0
            .SplitRow = 1
        End With
        ActiveWindow.FreezePanes = True
    ActiveWorkbook.Close savechanges:=True
    Application.DisplayAlerts = True
    Set OutBook = Nothing                                                                    '取消OutBook
End Sub
''' 创建考勤请假表
Sub OutToLeave(InSource, InRmax, InCmax, OutLeavePath, InName)
    Dim OutBook As Workbook
    Application.SheetsInNewWorkbook = 1                                                      ' 设置1个Sheet
    Set OutBook = Workbooks.Add
    ActiveWindow.FreezePanes = False                                                         ' 禁止冻结窗口
    Application.DisplayAlerts = False
        OutBook.SaveAs Filename:=OutLeavePath & "\" & InName & ".xlsx"
        Sheets(1).name = InName
        Range(Cells(1, 1), Cells(InRmax, InCmax)) = InSource                                 '调入请假信息
        Range(Cells(1, 1), Cells(1, InCmax)).Select
        Call 合并选中单元格
        Range(Cells(1, 1), Cells(InRmax, InCmax)).Select
        Call FontSet(NameFont)                                                               '设置字体格式
        Call 格式化
        Range(Cells(1, 1), Cells(1, InCmax)).Select
        TitleSize (20)
        Range(Cells(2, 1), Cells(2, InCmax)).Select
        TitleSize (14)
''''
        Cells(1, 1).Select                                                                   ' 自动设置行高和列宽
        Selection.CurrentRegion.Select
        Selection.Rows.AutoFit
        Selection.Columns.AutoFit
        Cells(1, 1).Select
    ActiveWorkbook.Close savechanges:=True
    Application.DisplayAlerts = True
    Set OutBook = Nothing                                                                    '取消OutBook
End Sub
Sub TitleSize(TTsize)
'
    With Selection.Font
        .name = "宋体"
        .Size = TTsize
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
End Sub
Sub AddToTotalLeave()
    Dim LeaveFolder As String
    Dim ToLeaveName As String
    Dim LeaveFile  As String
    Dim LeaveOld As Variant
    Dim LORow As Integer
    Dim LORowN As Integer
    Dim i, j, k, l, m As Integer
    ActiveWindow.FreezePanes = False                                                                            ' 禁止冻结窗口
    LeaveFolder = OutPath & "\" & Format(Now, "yyyy" & "年") & "统计总表"
    ToLeaveName = Format(Now, "yyyy" & "年") & "请假总表"
    LeaveFile = LeaveFolder & "\" & ToLeaveName & ".xlsx"
    Dim SFO As Object
    Set SFO = CreateObject("Scripting.FileSystemObject")                                                        '设SFO为文件夹对象变量
    If SFO.FolderExists(LeaveFolder) = False Then
       MkDir LeaveFolder
    End If
    If SFO.FileExists(LeaveFile) = False Then
       Call OutToLeave(Leave, UBound(Leave, 1), UBound(Leave, 2), LeaveFolder, ToLeaveName)
    End If
    Workbooks.Open Filename:=LeaveFile
    LeaveOld = Sheets(ToLeaveName).Range("a1:d" & RowMax)
    LORow = 0
    LORowN = 0
    For i = 1 To UBound(LeaveOld, 1)
        If LeaveOld(i, 1) > 0 Then
            LORow = LORow + 1
        End If
    Next
    For i = 1 To UBound(Leave, 1)
        If Leave(i, 1) > 0 Then
            LORowN = LORowN + 1
        End If
    Next
    k = LORow
    For i = 3 To LORowN
        m = 0
        For j = 3 To LORow                                                                                      ' 前两行是标题，所认空的，因此不能从1开始
            If IsDate(LeaveOld(j, 2)) Then                                                                      ' 如为日期格式，则需要格式化后再对比
               LeaveOld(j, 2) = Format(LeaveOld(j, 2), "m" & "月" & "d" & "日")
            End If
            If InStr(Leave(i, 3), LeaveOld(j, 3)) > 0 And InStr(Leave(i, 2), LeaveOld(j, 2)) > 0 Then
                   m = 1
            End If
        Next
        If m = 0 Then
            k = k + 1
            For l = 1 To 4
                LeaveOld(k, l) = Leave(i, l)
            Next
        End If
    Next
    Sheets(ToLeaveName).Range("a1:d" & RowMax) = LeaveOld
    Range(Cells(1, 1), Cells(UBound(LeaveOld, 1), UBound(LeaveOld, 2))).Select
    Call FontSet(NameFont)                                                                                      ' 设置字体格式
    Call 格式化
    Sheets(ToLeaveName).Range(Cells(1, 1), Cells(1, 4)).Select
    TitleSize (20)
    Sheets(ToLeaveName).Range(Cells(2, 1), Cells(2, 4)).Select
    TitleSize (14)
''''
    Cells(1, 1).Select                                                                                         ' 自动设置行高和列宽
    Selection.CurrentRegion.Select
    Selection.Rows.AutoFit
    Selection.Columns.AutoFit
    Cells(1, 1).Select
    Workbooks(ToLeaveName).Close savechanges:=True
End Sub
