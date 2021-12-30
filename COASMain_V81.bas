Attribute VB_Name = "COASMain_V81"
' 项目：COASMain
' 版本：V81
' 作者：冯振华
' 单位：山东省平原县第一中学
' 邮箱：fengzhenhua@outlook.com
' 博客：https://fengzhenhua-vip.github.io
' 主页：https://github.com/fengzhenhua-vip
' 版权：2021年7月13日--2021年9月23日
' 日志：考勤系统第7版，是从顶层设计实现的，综合V68之前的工作，升级了一些技术，最终实现全年级考核5秒内完成。其特点包括：
'       1. 引入职务号：教师1，班主任2
'       2. 引入控制号：Source 数组中，第20列之后表示控制号，21签到色，22签退色，23职务号，24调课信息，25自习情况
'       3. 引入班次号：1上午，2下午，3晚上
'       4. 引入对时表：DuiShiA(含日期),DuiShiB(含周)
'       5. 引入输出号：根据职务号输出对应的数组Source1,Source2
'       6. 引入考勤情况号：1正常，2警告，3违规，4漏签，5信息备注（请假等）
'       7. 最大限度的减少匹配次数，最大限度减少单元格选定（只有1处）
'       8. 升级了颜色处理技术和单元格合并技术
'       9. 由于从顶层设计，所以结构清析，易于拓展，同时达到全年级考核5秒内完成，其实打开和关闭文件占据了大多数时间
'       10. 由于重新构建及性能的大幅度提升，所以直接升级版本号V70
' 日志：优化了少量代码，提升稳定性及性能，升级版本号V71
' 日志：配置文件实现根据停止考勤日期去除相关考勤记录功能，主要修改了GetSource模块
'       将输出的列都设置成变量，暂时不变动顺序，考虑到输出结果重新设计的要求，所以设置了这项稍显麻烦的操作，使用这些变量的
'       数组包括：Source,Source1,Source2,DuiShiA,DuiShiB,InSource,OutSource及GenerateBook中的上色部分，升级版本号V72
' 日志：精简了GetHoliday,因为不需要职务的判断，所以去除了这部分判断，升级版本号V73
' 日志：修改了配置文件考勤标准时间Sheet，去除了10CT和51CT的设置，改为由10ST和51ST直接计算，配置更加清析，同时对V73进行了改造
'       在生成对时表时，增加因调课D而导致的下午可早退优惠39分钟的判断。升级版本号V74
' 日志：加入配置文件版本匹配检测，此版之后配置文件将修改为：配置文件V75.xlsm 的形式,以配合COASMain ,升级版本V75
' 日志：优化了考勤核心代码，更加方便集中维护，增加班主任弹性考勤并以蓝色标出，升级版本V76
' 日志：修改了判断教师和班主任以职务为准，不再以排班判断2021/9/3
' 日志：修复自习表生成对时表B的bug,2021/9/9 ，升级版本号V77
' 日志：修复自习表及二次校准表，班主任弹性考核bug,本来可以再进一步完善，但是考虑到考研可能占据我更多的时间，所以初步实现后不再计划升级改造。此次修复改动较多，升级版本号V78
' 日志：更改了考勤方式，使之更加易于配置，考核更加准确，升级版本号为V80
' 日志：修复请假上报表的自动生成2021/9/23
' 日志：修复考勤五一和十一的自动切换bug,将换课表限定于一周内暂时可以不必拓展功能，升级版本号V81
Public Const COAVersion As Integer = 81
Public Const RowMax As Integer = 10000
Public Const ColMax As Integer = 1000
Public Const SubColMax As Integer = 25
Public ConfigPath, ConfigFolder, ConfigFile As String
Public OutPath, OutFolder, OutFileFix As String
Public DateFormat, TimeFormat As String
Public NameFont, NameOriginal, NameTeacherUN, NameHeadMasterUN As String
Public ConfigBook As Workbook
Public ConfigSheet1, ConfigSheet2, ConfigSheet3 As String
Public OriginalSheet1, OriginalSheet2, OriginalSheet3, OriginalSheet4, OriginalSheet5 As String
Public WuYi, ShiYi As Date
Public StopSymbol As String
Public VipSource As Variant
Public ViRmax, ViCmax As Integer
Public Source(1 To RowMax, 1 To SubColMax) As Variant
Public Source1(1 To RowMax, 1 To SubColMax) As Variant
Public Source2(1 To RowMax, 1 To SubColMax) As Variant
Public DateMin, DateMax As Date
Public SRowMax, S1RowMax, S2RowMax As Integer
Public Change As Variant
Public CGRmax, CGCmax As Integer
Public SelfStudyTable As Variant
Public SSRmax, SSCmax As Integer
Public DSRAmax, DSRBmax As Integer
Public ReCorrectTable As Variant
Public RCTRmax, RCTCmax As Integer
Public Holiday As Variant
Public HolidayA(1 To RowMax, 1 To SubColMax) As Variant
Public HolidayB(1 To RowMax, 1 To SubColMax) As Variant
Public HRmax, HCmax, HARowMax, HBRowMax As Integer
Public TeacherGroup As Variant
Public TGRmax, TGCmax, GroupRow, GroupColum As Integer
Public TeGrStep, TeGrZhiWu As Integer
Public PreLeave, Leave, NameLeave As Variant
Public DateX, DateY, DateZ As Date
Public WeekX, WeekY, WeekZ As Integer
Public Standard As Variant
Public STRmax, STCmax As Integer
Public SFO As Object
Public COAXingMing, COARiQi, COAZhou, COABanCi, COAZiXi, COAQianDao, COAQianTui, COAQianDaoSe, COAQianTuiSe, COAZhiWu As Integer
Public COAShangChi, COAShangTui, COAShangLou, COAXiaChi, COAXiaTui, COAXiaLou, COAWanChi, COAWanTui, COAWanLou, COAHuanKe As Integer
Public BZSNum, BZXNum, BZWNum, BZSNumX, BZXNumX, BZWNumX As Integer
Public BZBeginNum, BZBeginNumX As Integer
Public Nianji As String
    
'
Sub COASMain()
    Application.ScreenUpdating = False
    Call COAConfigSet
    Call GetSource
    Call GetLeaveBook
    Call GetQingJia
    Call COAHolidayADD
    Call COARecorrectADD
'    Call COAChangeADD     ' 由于不计划拓展本程序，故将程序限定在一周内的考勤，于是不再处理复杂的适用于所有情况的换课
    Call COASelfStudyMOD   ' 将换课限定在一周内，所以将换课表直接调整自习表，于是问题变得容易处理，暂时采纳
    Call COASelfStudyADD
    Call COAChangeADD
    Call COANormalEXE
    Call COARecorrectBAC
    Call COAGenerateEXE
    Application.ScreenUpdating = True
End Sub

Sub COAConfigSet()
 '   ConfigPath = "D:\考勤系统"
    ConfigPath = CreateObject("WScript.Shell").specialfolders("Desktop") & "\考勤系统"                          '默认设置为桌面
    ConfigFolder = ConfigPath & "\" & "考勤系统配置"
    ConfigFile = ConfigFolder & "\" & "考勤配置V" & COAVersion & ".xlsm"                                        '开启增强，自动校准功能的配置文件
    Set SFO = CreateObject("Scripting.FileSystemObject")                                                        '设SFO为文件夹对象变量
    If SFO.FileExists(ConfigFile) = False Then
        MsgBox "配置文件版本与当前系统不符！请使用：" & ConfigFolder & "\考勤配置V" & COAVersion & ".xlsm"
        End
    Else
        OutPath = ConfigPath & "\" & Format(Now, "yyyy" & "年") & "考勤"
        DateFormat = "m""月""d""日"";@"
        TimeFormat = "h:mm;@"
        NameFont = "宋体"
        NameHeadMasterUN = "异常班主任"
        NameTeacherUN = "异常教师"
        ConfigSheet1 = "停止考勤"
        ConfigSheet2 = "请假表"
        OriginalSheet1 = "自习安排"                                                                                 '以下4项专为校准表设置
        OriginalSheet3 = "二次校准"
        OriginalSheet4 = "调课表"
        OriginalSheet5 = "教师分组"
        WuYi = Format(Now, "yyyy") & "/5/1"
        ShiYi = Format(Now, "yyyy") & "/10/1"
        COAXingMing = 1: COARiQi = 2: COAZhou = 4: COABanCi = 3: COAQianDao = 5: COAQianTui = 6
        COAShangChi = 7: COAShangTui = COAShangChi + 1: COAShangLou = COAShangChi + 2                               '确保统计区上午、下午、晚上连续
        COAXiaChi = 10: COAXiaTui = COAXiaChi + 1: COAXiaLou = COAXiaChi + 2
        COAWanChi = 13: COAWanTui = COAWanChi + 1: COAWanLou = COAWanChi + 2
        COAQianDaoSe = 21: COAQianTuiSe = 22: COAZhiWu = 23: COAHuanKe = 24: COAZiXi = 25
        If WuYi < Now And Now < ShiYi Then
            ConfigSheet3 = "51ST"
        Else
            ConfigSheet3 = "10ST"
        End If
        StopSymbol = "*"
        If SFO.FolderExists(ConfigPath) = False Then
           MkDir ConfigPath
        End If
        If SFO.FolderExists(ConfigFolder) = False Then
           MkDir ConfigFolder
        End If
        If SFO.FolderExists(OutPath) = False Then
           MkDir OutPath
        End If
    ' 统一调入配置表
        Set ConfigBook = GetObject(ConfigFile)
        ViRmax = ConfigBook.Sheets(ConfigSheet1).Cells(RowMax, 1).End(xlUp).Row
        ViCmax = ConfigBook.Sheets(ConfigSheet1).Cells(1, ColMax).End(xlToLeft).Column
        VipSource = ConfigBook.Sheets(ConfigSheet1).Range(ConfigBook.Sheets(ConfigSheet1).Cells(1, 1), ConfigBook.Sheets(ConfigSheet1).Cells(ViRmax, ViCmax))
    ' 获取对时表A
        CGRmax = ConfigBook.Sheets(OriginalSheet4).Cells(RowMax, 1).End(xlUp).Row                                                                    '调入换课表
        CGCmax = ConfigBook.Sheets(OriginalSheet4).Cells(1, ColMax).End(xlToLeft).Column
        Change = ConfigBook.Sheets(OriginalSheet4).Range(ConfigBook.Sheets(OriginalSheet4).Cells(1, 1), ConfigBook.Sheets(OriginalSheet4).Cells(CGRmax, CGCmax))
        SSRmax = ConfigBook.Sheets(OriginalSheet1).Cells(RowMax, 1).End(xlUp).Row
        SSCmax = ConfigBook.Sheets(OriginalSheet1).Cells(1, ColMax).End(xlToLeft).Column
        SelfStudyTable = ConfigBook.Sheets(OriginalSheet1).Range(ConfigBook.Sheets(OriginalSheet1).Cells(1, 1), ConfigBook.Sheets(OriginalSheet1).Cells(SSRmax, SSCmax))
        HRmax = ConfigBook.Sheets(ConfigSheet2).Cells(RowMax, 1).End(xlUp).Row
        HCmax = ConfigBook.Sheets(ConfigSheet2).Cells(1, ColMax).End(xlToLeft).Column
        Holiday = ConfigBook.Sheets(ConfigSheet2).Range(ConfigBook.Sheets(ConfigSheet2).Cells(1, 1), ConfigBook.Sheets(ConfigSheet2).Cells(HRmax, HCmax))
        RCTRmax = ConfigBook.Sheets(OriginalSheet3).Cells(RowMax, 1).End(xlUp).Row
        RCTCmax = ConfigBook.Sheets(OriginalSheet3).Cells(1, ColMax).End(xlToLeft).Column
        ReCorrectTable = ConfigBook.Sheets(OriginalSheet3).Range(ConfigBook.Sheets(OriginalSheet3).Cells(1, 1), ConfigBook.Sheets(OriginalSheet3).Cells(RCTRmax, RCTCmax))
        TGRmax = ConfigBook.Sheets(OriginalSheet5).Cells(RowMax, 1).End(xlUp).Row
        TGCmax = ConfigBook.Sheets(OriginalSheet5).Cells(2, ColMax).End(xlToLeft).Column
        TeacherGroup = ConfigBook.Sheets(OriginalSheet5).Range(ConfigBook.Sheets(OriginalSheet5).Cells(1, 1), ConfigBook.Sheets(OriginalSheet5).Cells(TGRmax, TGCmax))
        TeGrStep = 2: TeGrZhiWu = 2
        Do Until TeacherGroup(1, TeGrStep + 1) > 0
            TeGrStep = TeGrStep + 1
        Loop
        STRmax = ConfigBook.Sheets(ConfigSheet3).Cells(RowMax, 1).End(xlUp).Row
        STCmax = ConfigBook.Sheets(ConfigSheet3).Cells(1, ColMax).End(xlToLeft).Column
        Standard = ConfigBook.Sheets(ConfigSheet3).Range(ConfigBook.Sheets(ConfigSheet3).Cells(1, 1), ConfigBook.Sheets(ConfigSheet3).Cells(STRmax, STCmax))
    End If
    Nianji = "高三文理部"
End Sub

Sub GetSource()
' 取得原始表行数和列数
    DeRmax = Sheets(1).Cells(RowMax, 1).End(xlUp).Row
    DeCmax = 13
' 生成标准通用考勤数组格式
    Dim DeLiSource As Variant
    Dim i, j, k, m As Integer
    DeLiSource = Sheets(1).Range(Cells(1, 3), Cells(DeRmax, DeCmax)).Value
    j = 1
    Source(j, COAXingMing) = "姓名": Source(j, COARiQi) = "日期(周)": Source(j, COABanCi) = "班次"
    Source(j, COAZhou) = "自习": Source(j, COAQianDao) = "签到": Source(j, COAQianTui) = "签退"
    Source(j, COAShangChi) = "上迟": Source(j, COAShangTui) = "上退": Source(j, COAShangLou) = "上漏"
    Source(j, COAXiaChi) = "下迟": Source(j, COAXiaTui) = "下退": Source(j, COAXiaLou) = "下漏"
    Source(j, COAWanChi) = "晚迟": Source(j, COAWanTui) = "晚退": Source(j, COAWanLou) = "晚漏"
    DateMin = CDate(DeLiSource(3, 4)): DateMax = DateMin
    For i = 3 To DeRmax                                                                                      'DeLi数据是前两行合并过，所以第2行是空的不必引入
        m = 0
        For k = 2 To ViRmax
            If Len(VipSource(k, 2)) > 0 And Len(VipSource(k, 3)) > 0 Then
                If CDate(VipSource(k, 2)) <= CDate(DeLiSource(i, 4)) And CDate(DeLiSource(i, 4)) <= VipSource(k, 3) Then
                    If Len(VipSource(k, 4)) > 0 Or Len(VipSource(k, 5)) > 0 Or Len(VipSource(k, 6)) > 0 Then
                        For j = 4 To 6
                            If Len(VipSource(k, j)) > 0 Then
                                If InStr(DeLiSource(i, 7), VipSource(1, j)) > 0 Then
                                    If InStr(VBA.Trim(Mid(DeLiSource(i, 1), 1, 1)) & VBA.Trim(Mid(DeLiSource(i, 1), 2)), VipSource(k, 1)) > 0 Then
                                        m = k
                                    ElseIf InStr(DeLiSource(i, 7), VipSource(k, 1)) > 0 Then
                                        m = k
                                    ElseIf InStr(VipSource(k, 1), "*") Then
                                        m = 1
                                    End If
                                End If
                            End If
                        Next
                    Else
                        If InStr(VBA.Trim(Mid(DeLiSource(i, 1), 1, 1)) & VBA.Trim(Mid(DeLiSource(i, 1), 2)), VipSource(k, 1)) > 0 Then
                            m = k
                        ElseIf InStr(DeLiSource(i, 7), VipSource(k, 1)) > 0 Then
                            m = k
                        ElseIf InStr(VipSource(k, 1), "*") Then
                            m = 1
                        End If
                    End If
                End If
            Else
                If Len(VipSource(k, 4)) > 0 Or Len(VipSource(k, 5)) > 0 Or Len(VipSource(k, 6)) > 0 Then
                    For j = 4 To 6
                        If Len(VipSource(k, j)) > 0 Then
                            If InStr(DeLiSource(i, 7), VipSource(1, j)) > 0 Then
                                If InStr(VBA.Trim(Mid(DeLiSource(i, 1), 1, 1)) & VBA.Trim(Mid(DeLiSource(i, 1), 2)), VipSource(k, 1)) > 0 Then
                                    m = k
                                ElseIf InStr(DeLiSource(i, 7), VipSource(k, 1)) > 0 Then
                                    m = k
                                ElseIf InStr(VipSource(k, 1), "*") Then
                                    m = 1
                                End If
                            End If
                        End If
                    Next
                Else
                    If InStr(VBA.Trim(Mid(DeLiSource(i, 1), 1, 1)) & VBA.Trim(Mid(DeLiSource(i, 1), 2)), VipSource(k, 1)) > 0 Then
                        m = k
                    ElseIf InStr(DeLiSource(i, 7), VipSource(k, 1)) > 0 Then
                        m = k
                    End If
                End If
            End If
        Next
        If m = 0 And (InStr(DeLiSource(i, 3), "班主任") > 0 Or (InStr(DeLiSource(i, 3), "班主任") = 0 And InStr(DeLiSource(i, 5), "六") = 0 And InStr(DeLiSource(i, 5), "日") = 0)) Then
            j = j + 1
            Source(j, COAXingMing) = VBA.Trim(Mid(DeLiSource(i, 1), 1, 1)) & VBA.Trim(Mid(DeLiSource(i, 1), 2))   '简单和去除二字姓名中间的空格,不再使用统一的空格消除
            Source(j, COARiQi) = CDate(DeLiSource(i, 4))
            If DateMin > Source(j, COARiQi) Then
                DateMin = Source(j, COARiQi)
            End If
            If DateMax < Source(j, COARiQi) Then
                DateMax = Source(j, COARiQi)
            End If
            If InStr(DeLiSource(i, 7), "上午") > 0 Then
                Source(j, COABanCi) = 1
            ElseIf InStr(DeLiSource(i, 7), "下午") > 0 Then
                Source(j, COABanCi) = 2
            ElseIf InStr(DeLiSource(i, 7), "晚上") > 0 Then
                Source(j, COABanCi) = 3
            End If
            If InStr(DeLiSource(i, 3), "班主任") > 0 Then
                Source(j, COAZhiWu) = 2
            ElseIf Len(DeLiSource(i, 7)) > 0 Then   ' 对于中间加入的人员，由于没有考勤排班，所以为空，此不应当考核
                Source(j, COAZhiWu) = 1
            End If
            Select Case DeLiSource(i, 5)
                Case Is = "一"
                    Source(j, COAZhou) = 1
                Case Is = "二"
                    Source(j, COAZhou) = 2
                Case Is = "三"
                    Source(j, COAZhou) = 3
                Case Is = "四"
                    Source(j, COAZhou) = 4
                Case Is = "五"
                    Source(j, COAZhou) = 5
                Case Is = "六"
                    Source(j, COAZhou) = 6
                Case Is = "日"
                    Source(j, COAZhou) = 7
            End Select
            If DeLiSource(i, 10) > 0 Then
                Source(j, COAQianDao) = CDate(DeLiSource(i, 10))
            End If
            If DeLiSource(i, 11) > 0 Then
                Source(j, COAQianTui) = CDate(DeLiSource(i, 11))
            End If
        End If
    Next
    SRowMax = j
' 设置班主任弹性考核起始点debug
    If Source(2, COAZhou) = 6 Or Source(2, COAZhou) = 6 Then
       BZBeginNum = 1: BZBeginNumX = Source(2, COAZhou)
    Else
       BZBeginNum = Source(2, COAZhou): BZBeginNumX = 6
    End If
' 生成目标Excel文件的后缀、输出文件夹
    OutFileFix = "（" & Format(DateMin, "yyyy" & "年" & "m" & "月" & "d" & "日") & "-" & Format(DateMax, "yyyy" & "年" & "m" & "月" & "d" & "日") & "）"
    OutFolder = OutPath & "\" & Format(DateMax, "m" & "月" & "d" & "日") & "正式上报"
    If SFO.FolderExists(OutFolder) = False Then
       MkDir OutFolder
    Else
        i = 1
        Do While SFO.FolderExists(OutFolder & i) = True
            i = i + 1
        Loop
        OutFolder = OutFolder & i
        MkDir OutFolder
    End If
End Sub
Sub GetQingJia() '本程序用来解析请假列表
' 生成请假的校准表HolidayA,HolidayB，二者均无标题，从第1行起既是有效数据，用来生成考勤统计信息
    HARowMax = 0: HBRowMax = 0
    For i = 2 To HRmax
' 取得教师分组和对应行数
       If InStr(Holiday(i, 1), "组") > 0 Or InStr(Holiday(i, 1), "班主任") > 0 Or InStr(Holiday(i, 1), "督导室") > 0 Then
           For j = 1 To TGCmax Step TeGrStep
               If TeacherGroup(1, j) = Holiday(i, 1) Then
                   GroupColum = j
               End If
           Next
           GroupRow = ConfigBook.Sheets(OriginalSheet5).Cells(RowMax, GroupColum).End(xlUp).Row
       End If
'获取日期格式请表HolidayA
       If IsDate(Holiday(i, 2)) Then
               If InStr(Holiday(i, 1), "组") > 0 Or InStr(Holiday(i, 1), "班主任") > 0 Or InStr(Holiday(i, 1), "督导室") > 0 Then
                   For a = 3 To GroupRow
                         If Len(Holiday(i, 4)) > 0 Then
                             DateZ = CDate(Holiday(i, 2))
                             Do While DateZ <= CDate(Holiday(i, 3))
                                If DateMin <= DateZ And DateZ <= DateMax Then
                                    HARowMax = HARowMax + 1
                                    HolidayA(HARowMax, COAXingMing) = TeacherGroup(a, GroupColum)
                                    HolidayA(HARowMax, COARiQi) = DateZ
                                    HolidayA(HARowMax, COAQianDao) = Holiday(i, 4)
                                    HolidayA(HARowMax, COAQianTui) = Holiday(i, 4)
                                End If
                                DateZ = DateZ + 1
                             Loop
                         Else
                             For k = 5 To 9 Step 2
                                 DateZ = CDate(Holiday(i, 2))
                                 If Len(Holiday(i, k)) > 0 Or Len(Holiday(i, k + 1)) > 0 Then
                                    Do While DateZ <= CDate(Holiday(i, 3))
                                        If DateMin <= DateZ And DateZ <= DateMax Then
                                            HARowMax = HARowMax + 1
                                            HolidayA(HARowMax, COAXingMing) = TeacherGroup(a, GroupColum)
                                            HolidayA(HARowMax, COARiQi) = DateZ
                                            If InStr(Holiday(1, k), "上午") > 0 Then
                                                HolidayA(HARowMax, COABanCi) = 1
                                            ElseIf InStr(Holiday(1, k), "下午") > 0 Then
                                                HolidayA(HARowMax, COABanCi) = 2
                                            ElseIf InStr(Holiday(1, k), "晚上") > 0 Then
                                                HolidayA(HARowMax, COABanCi) = 3
                                            Else
                                                MsgBox "班次出错"
                                            End If
                                            If Len(Holiday(i, k)) > 0 Then HolidayA(HARowMax, 5) = Holiday(i, k)
                                            If Len(Holiday(i, k + 1)) > 0 Then HolidayA(HARowMax, 6) = Holiday(i, k + 1)
                                         End If
                                         DateZ = DateZ + 1
                                    Loop
                                 End If
                              Next
                         End If
                   Next
                 Else
                   If Len(Holiday(i, 4)) > 0 Then
                      DateZ = CDate(Holiday(i, 2))
                      Do While DateZ <= CDate(Holiday(i, 3))
                           If DateMin <= DateZ And DateZ <= DateMax Then
                                HARowMax = HARowMax + 1
                                HolidayA(HARowMax, COAXingMing) = Holiday(i, 1)
                                HolidayA(HARowMax, COARiQi) = DateZ
                                HolidayA(HARowMax, COAQianDao) = Holiday(i, 4)
                                HolidayA(HARowMax, COAQianTui) = Holiday(i, 4)
                           End If
                           DateZ = DateZ + 1
                      Loop
                   Else
                      For k = 5 To 9 Step 2
                          DateZ = CDate(Holiday(i, 2))
                          If Len(Holiday(i, k)) > 0 Or Len(Holiday(i, k + 1)) > 0 Then
                            Do While CDate(DateZ) <= CDate(Holiday(i, 3))
                                 If DateMin <= DateZ And DateZ <= DateMax Then
                                    HARowMax = HARowMax + 1
                                    HolidayA(HARowMax, COAXingMing) = Holiday(i, 1)
                                    HolidayA(HARowMax, COARiQi) = DateZ
                                    If InStr(Holiday(1, k), "上午") > 0 Then
                                       HolidayA(HARowMax, COABanCi) = 1
                                    ElseIf InStr(Holiday(1, k), "下午") > 0 Then
                                       HolidayA(HARowMax, COABanCi) = 2
                                    ElseIf InStr(Holiday(1, k), "晚上") > 0 Then
                                       HolidayA(HARowMax, COABanCi) = 3
                                    Else
                                       MsgBox "班次出错"
                                    End If
                                    If Len(Holiday(i, k)) > 0 Then HolidayA(HARowMax, 5) = Holiday(i, k)
                                    If Len(Holiday(i, k + 1)) > 0 Then HolidayA(HARowMax, 6) = Holiday(i, k + 1)
                                 End If
                                 DateZ = DateZ + 1
                            Loop
                          End If
                       Next
                 End If
           End If
    Else
' 获取非日期格式请假表HolidayB
           For j = 1 To 7
            If Holiday(i, 2) = Choose(j, "一", "二", "三", "四", "五", "六", "日") Then
               WeekX = j
            End If
           Next
           For j = 1 To 7
            If Holiday(i, 3) = Choose(j, "一", "二", "三", "四", "五", "六", "日") Then
               WeekY = j
            End If
           Next
           If InStr(Holiday(i, 1), "组") > 0 Or InStr(Holiday(i, 1), "班主任") > 0 Or InStr(Holiday(i, 1), "督导室") > 0 Then
               For a = 3 To GroupRow
                    If Len(Holiday(i, 4)) > 0 Then
                        WeekZ = WeekX
                        Do While WeekZ <= WeekY
                            HBRowMax = HBRowMax + 1
                            HolidayB(HBRowMax, COAXingMing) = TeacherGroup(a, GroupColum)
                            HolidayB(HBRowMax, COAQianDao) = Holiday(i, 4)
                            HolidayB(HBRowMax, COAQianTui) = Holiday(i, 4)
                            WeekZ = WeekZ + 1
                        Loop
                    Else
                        For k = 5 To 9 Step 2
                            WeekZ = WeekX
                            If Len(Holiday(i, k)) > 0 Or Len(Holiday(i, k + 1)) > 0 Then
                                Do While WeekZ <= WeekY
                                    HBRowMax = HBRowMax + 1
                                    HolidayB(HBRowMax, COAXingMing) = TeacherGroup(a, GroupColum)
                                    If InStr(Holiday(1, k), "上午") > 0 Then
                                        HolidayB(HBRowMax, COABanCi) = 1
                                     ElseIf InStr(Holiday(1, k), "下午") > 0 Then
                                        HolidayB(HBRowMax, COABanCi) = 2
                                     ElseIf InStr(Holiday(1, k), "晚上") > 0 Then
                                        HolidayB(HBRowMax, COABanCi) = 3
                                     Else
                                        MsgBox "班次出错"
                                     End If
                                     HolidayB(HBRowMax, COAZhou) = WeekZ
                                    If Len(Holiday(i, k)) > 0 Then HolidayB(HBRowMax, COAQianDao) = Holiday(i, k)
                                    If Len(Holiday(i, k + 1)) > 0 Then HolidayB(HBRowMax, COAQianTui) = Holiday(i, k + 1)
                                    WeekZ = WeekZ + 1
                                Loop
                            End If
                         Next
                    End If
                Next
           Else
                If Len(Holiday(i, 4)) > 0 Then
                    WeekZ = WeekX
                    Do While WeekZ <= WeekY
                        HBRowMax = HBRowMax + 1
                        HolidayB(HBRowMax, COAXingMing) = Holiday(i, 1)
                        HolidayB(HBRowMax, COAQianDao) = Holiday(i, 4)
                        HolidayB(HBRowMax, COAQianTui) = Holiday(i, 4)
                        WeekZ = WeekZ + 1
                    Loop
                Else
                     For k = 5 To 9 Step 2
                       WeekZ = WeekX
                       If Len(Holiday(i, k)) > 0 Or Len(Holiday(i, k + 1)) > 0 Then
                           Do While WeekZ <= WeekY
                               HBRowMax = HBRowMax + 1
                               HolidayB(HBRowMax, COAXingMing) = Holiday(i, 1)
                               If InStr(Holiday(1, k), "上午") > 0 Then
                                   HolidayB(HBRowMax, COABanCi) = 1
                               ElseIf InStr(Holiday(1, k), "下午") > 0 Then
                                   HolidayB(HBRowMax, COABanCi) = 2
                               ElseIf InStr(Holiday(1, k), "晚上") > 0 Then
                                   HolidayB(HBRowMax, COABanCi) = 3
                               Else
                                   HolidayB(HBRowMax, COABanCi) = Holiday(1, k)
                               End If
                               HolidayB(HBRowMax, COAZhou) = WeekZ
                               If Len(Holiday(i, k)) > 0 Then HolidayB(HBRowMax, COAQianDao) = Holiday(i, k)
                               If Len(Holiday(i, k + 1)) > 0 Then HolidayB(HBRowMax, COAQianTui) = Holiday(i, k + 1)
                               WeekZ = WeekZ + 1
                           Loop
                       End If
                    Next
               End If
           End If
       End If
    Next
End Sub
Sub COAHolidayADD() '添加请假信息
    Dim i, j, m As Integer
    For i = 2 To SRowMax
        m = 0
        For j = 1 To HARowMax
            If Len(HolidayA(j, COAXingMing)) > 0 And (InStr(Source(i, COAXingMing), HolidayA(j, COAXingMing)) > 0 Or InStr(HolidayA(j, COAXingMing), "*") > 0) And InStr(Source(i, COARiQi), HolidayA(j, COARiQi)) > 0 And InStr(Source(i, COABanCi), HolidayA(j, COABanCi)) > 0 Then
                 If Len(Source(i, COAQianDao)) = 0 And Len(HolidayA(j, COAQianDao)) > 0 Then Source(i, COAQianDao) = HolidayA(j, COAQianDao): m = 1: Source(i, COAQianDaoSe) = 5
                 If Len(Source(i, COAQianTui)) = 0 And Len(HolidayA(j, COAQianTui)) > 0 Then Source(i, COAQianTui) = HolidayA(j, COAQianTui): m = 1: Source(i, COAQianTuiSe) = 5
            End If
        Next
        If m = 0 Then
            For j = 1 To HBRowMax
                If Len(HolidayB(j, COAXingMing)) > 0 And (InStr(Source(i, COAXingMing), HolidayB(j, COAXingMing)) > 0 Or InStr(HolidayB(j, COAXingMing), "*") > 0) And InStr(Source(i, COABanCi), HolidayB(j, COABanCi)) > 0 And InStr(Source(i, COAZhou), HolidayB(j, COAZhou)) > 0 Then
                    If Len(Source(i, COAQianDao)) = 0 And Len(HolidayB(j, COAQianDao)) > 0 Then Source(i, COAQianDao) = HolidayB(j, COAQianDao): m = 1: Source(i, COAQianDaoSe) = 5
                    If Len(Source(i, COAQianTui)) = 0 And Len(HolidayB(j, COAQianTui)) > 0 Then Source(i, COAQianTui) = HolidayB(j, COAQianTui): m = 1: Source(i, COAQianTuiSe) = 5
                End If
            Next
        End If
    Next
End Sub
Sub COARecorrectADD() '二次校准时间
    For i = 2 To RCTRmax
        For j = 2 To SRowMax
          If Len(ReCorrectTable(i, 2)) > 0 And IsDate(ReCorrectTable(i, 2)) And Len(ReCorrectTable(i, 3)) > 0 And IsDate(ReCorrectTable(i, 3)) Then
            If CDate(ReCorrectTable(i, 2)) <= CDate(Source(j, COARiQi)) And CDate(Source(j, COARiQi)) <= CDate(ReCorrectTable(i, 3)) Then
                If InStr(ReCorrectTable(i, 1), Source(j, COAXingMing)) > 0 And InStr(ReCorrectTable(i, 4), Source(j, COABanCi)) > 0 And InStr(ReCorrectTable(i, 5), Source(j, COAZhou)) > 0 Then
                    If IsDate(Source(j, COAQianDao)) Then
                        Source(j, COAQianDao) = CDate(Source(j, COAQianDao)) + CDate(ReCorrectTable(i, 6)) - CDate(ReCorrectTable(i, 7))
                        Source(j, COAQianDao) = CDate(Source(j, COAQianDao))
                    End If
                    If IsDate(Source(j, COAQianTui)) Then
                        Source(j, COAQianTui) = CDate(Source(j, COAQianTui)) + CDate(ReCorrectTable(i, 8)) - CDate(ReCorrectTable(i, 9))
                        Source(j, COAQianTui) = CDate(Source(j, COAQianTui))
                    End If
                End If
            End If
          Else
            If InStr(ReCorrectTable(i, 1), Source(j, COAXingMing)) > 0 And InStr(ReCorrectTable(i, 4), Source(j, COABanCi)) > 0 And InStr(ReCorrectTable(i, 5), Source(j, COAZhou)) > 0 Then
                If IsDate(Source(j, COAQianDao)) Then
                    Source(j, COAQianDao) = CDate(Source(j, COAQianDao)) + CDate(ReCorrectTable(i, 6)) - CDate(ReCorrectTable(i, 7))
                    Source(j, COAQianDao) = CDate(Source(j, COAQianDao))
                End If
                If IsDate(Source(j, COAQianTui)) Then
                    Source(j, COAQianTui) = CDate(Source(j, COAQianTui)) + CDate(ReCorrectTable(i, 8)) - CDate(ReCorrectTable(i, 9))
                    Source(j, COAQianTui) = CDate(Source(j, COAQianTui))
                End If
            End If
          End If
        Next
    Next
End Sub
Sub COARecorrectBAC() '恢复二次校准时间
    For i = 2 To RCTRmax
        For j = 2 To SRowMax
          If Len(ReCorrectTable(i, 2)) > 0 And IsDate(ReCorrectTable(i, 2)) And Len(ReCorrectTable(i, 3)) > 0 And IsDate(ReCorrectTable(i, 3)) Then
            If CDate(ReCorrectTable(i, 2)) <= CDate(Source(j, COARiQi)) And CDate(Source(j, COARiQi)) <= CDate(ReCorrectTable(i, 3)) Then
                If InStr(ReCorrectTable(i, 1), Source(j, COAXingMing)) > 0 And InStr(ReCorrectTable(i, 4), Source(j, COABanCi)) > 0 And InStr(ReCorrectTable(i, 5), Source(j, COAZhou)) > 0 Then
                    If IsDate(Source(j, COAQianDao)) Then
                        Source(j, COAQianDao) = CDate(Source(j, COAQianDao)) - CDate(ReCorrectTable(i, 6)) + CDate(ReCorrectTable(i, 7))
                        Source(j, COAQianDao) = CDate(Source(j, COAQianDao))
                    End If
                    If IsDate(Source(j, COAQianTui)) Then
                        Source(j, COAQianTui) = CDate(Source(j, COAQianTui)) - CDate(ReCorrectTable(i, 8)) + CDate(ReCorrectTable(i, 9))
                        Source(j, COAQianTui) = CDate(Source(j, COAQianTui))
                    End If
                End If
            End If
          Else
            If InStr(ReCorrectTable(i, 1), Source(j, COAXingMing)) > 0 And InStr(ReCorrectTable(i, 4), Source(j, COABanCi)) > 0 And InStr(ReCorrectTable(i, 5), Source(j, COAZhou)) > 0 Then
                If IsDate(Source(j, COAQianDao)) Then
                    Source(j, COAQianDao) = CDate(Source(j, COAQianDao)) - CDate(ReCorrectTable(i, 6)) + CDate(ReCorrectTable(i, 7))
                    Source(j, COAQianDao) = CDate(Source(j, COAQianDao))
                End If
                If IsDate(Source(j, COAQianTui)) Then
                    Source(j, COAQianTui) = CDate(Source(j, COAQianTui)) - CDate(ReCorrectTable(i, 8)) + CDate(ReCorrectTable(i, 9))
                    Source(j, COAQianTui) = CDate(Source(j, COAQianTui))
                End If
            End If
          End If
        Next
    Next
End Sub

Sub COAChangeADD()
    Dim i, j As Integer
    For i = 2 To CGRmax
        For j = 2 To SRowMax
            If DateMin <= CDate(Change(i, 3)) And CDate(Change(i, 3)) <= DateMax Then
                If InStr(Source(j, COAXingMing), Change(i, 7)) > 0 And InStr(Source(j, COARiQi), Change(i, 3)) > 0 And InStr(Source(j, COAZhou), Change(i, 5)) > 0 Then
                    If InStr(Source(j, COABanCi), Change(i, 4)) > 0 Then
                        Source(j, COAHuanKe) = "调入" & Change(i, 1) & Change(i, 6) & "调出" & Change(i, 7) & Change(i, 12)
                    End If
                End If
            End If
            If DateMin <= CDate(Change(i, 9)) And CDate(Change(i, 9)) <= DateMax Then
                If InStr(Source(j, COAXingMing), Change(i, 1)) > 0 And InStr(Source(j, COARiQi), Change(i, 9)) > 0 And InStr(Source(j, COAZhou), Change(i, 11)) > 0 Then
                    Call COAChangeKernel(i, 12, j)
                    If InStr(Source(j, COABanCi), Change(i, 10)) > 0 Then
                        Source(j, COAHuanKe) = "调入" & Change(i, 7) & Change(i, 12) & "调出" & Change(i, 1) & Change(i, 6)
                    End If
                End If
            End If
        Next
    Next
End Sub
Sub COAChangeKernel(KerRowI, KerColI, KerRowJ)
    Dim KerZhiWu, i, m As Integer
    KerZhiWu = (Source(KerRowJ, COAZhiWu) - 1) * 5
    If InStr(Change(KerRowI, KerColI), "B") > 0 Then
        If Source(KerRowJ, COABanCi) = 1 Then
            Source(KerRowJ, COAZiXi) = Source(KerRowJ, COAZiXi) & "B"
            If Len(Source(KerRowJ, COAQianDaoSe)) = 0 And Len(Source(KerRowJ, COAQianDao)) > 0 And IsDate(Source(KerRowJ, COAQianDao)) Then
                If CDate(Source(KerRowJ, COAQianDao)) < CDate(Standard(5, 3 + KerZhiWu)) Then
                    Source(KerRowJ, COAQianDaoSe) = 1
                ElseIf CDate(Source(KerRowJ, COAQianDao)) <= CDate(Standard(5, 4 + KerZhiWu)) Then
                    Source(KerRowJ, COAQianDaoSe) = 2
                Else
                    Source(KerRowJ, COAQianDaoSe) = 3
                    Source(KerRowJ, COAShangChi) = 1
                End If
            End If
        Else
            For i = 1 To SSRmax      '自习表中不含D时才允许优惠
                If InStr(SelfStudyTable(i, 1), Source(KerRowJ, COAXingMing)) > 0 And InStr(SelfStudyTable(i, Source(KerRowJ, COAZhou) + 2), "D") = 0 Then
                    If Source(KerRowJ, COABanCi) = 2 Then
                        Source(KerRowJ, COAZiXi) = Source(KerRowJ, COAZiXi) & "★"
                        If Len(Source(KerRowJ, COAQianDaoSe)) = 0 And Len(Source(KerRowJ, COAQianDao)) > 0 And IsDate(Source(KerRowJ, COAQianDao)) Then
                            If CDate(Source(KerRowJ, COAQianDao)) < CDate(Standard(7, 3 + KerZhiWu)) Then
                                Source(KerRowJ, COAQianDaoSe) = 1
                            ElseIf CDate(Source(j, COAQianDao)) <= CDate(Standard(7, 4 + KerZhiWu)) Then
                                Source(KerRowJ, COAQianDaoSe) = 2
                            Else
                                Source(KerRowJ, COAQianDaoSe) = 3
                                Source(KerRowJ, COAXiaChi) = 1
                            End If
                        End If
                    End If
                End If
            Next
        End If
    ElseIf InStr(Change(KerRowI, KerColI), "C") > 0 Then
        If Source(KerRowJ, COABanCi) = 1 Then
            Source(KerRowJ, COAZiXi) = Source(KerRowJ, COAZiXi) & "C"
            If Len(Source(KerRowJ, COAQianTuiSe)) = 0 And Len(Source(KerRowJ, COAQianTui)) > 0 And IsDate(Source(KerRowJ, COAQianTui)) Then
                If CDate(Source(KerRowJ, COAQianTui)) > CDate(Standard(6, 6 + KerZhiWu)) Then
                    Source(KerRowJ, COAQianTuiSe) = 1
                ElseIf CDate(Source(j, COAQianTui)) >= CDate(Standard(6, 5 + KerZhiWu)) Then
                    Source(KerRowJ, COAQianTuiSe) = 2
                Else
                    Source(KerRowJ, COAQianTuiSe) = 3
                    Source(KerRowJ, COAShangTui) = 1
                End If
            End If
        Else
            For i = 1 To SSRmax      '自习表中不含D时才允许优惠
                If InStr(SelfStudyTable(i, 1), Source(KerRowJ, COAXingMing)) > 0 And InStr(SelfStudyTable(i, Source(KerRowJ, COAZhou) + 2), "D") = 0 Then
                    If Source(KerRowJ, COABanCi) = 2 Then
                        Source(KerRowJ, COAZiXi) = Source(KerRowJ, COAZiXi) & "★"
                        If Len(Source(KerRowJ, COAQianDaoSe)) = 0 And Len(Source(KerRowJ, COAQianDao)) > 0 And IsDate(Source(KerRowJ, COAQianDao)) Then
                            If CDate(Source(KerRowJ, COAQianDao)) < CDate(Standard(7, 3 + KerZhiWu)) Then
                                Source(KerRowJ, COAQianDaoSe) = 1
                            ElseIf CDate(Source(KerRowJ, COAQianDao)) <= CDate(Standard(7, 4 + KerZhiWu)) Then
                                Source(KerRowJ, COAQianDaoSe) = 2
                            Else
                                Source(KerRowJ, COAQianDaoSe) = 3
                                Source(KerRowJ, COAXiaChi) = 1
                            End If
                        End If
                    End If
                End If
            Next
        End If
    ElseIf InStr(Change(KerRowI, KerColI), "D") > 0 Then
        If Source(KerRowJ, COABanCi) = 2 Then
            Source(KerRowJ, COAZiXi) = Source(KerRowJ, COAZiXi) & "D"
            If Len(Source(KerRowJ, COAQianDaoSe)) = 0 And Len(Source(KerRowJ, COAQianDao)) > 0 And IsDate(Source(KerRowJ, COAQianDao)) Then
                If CDate(Source(KerRowJ, COAQianDao)) < CDate(Standard(8, 3 + KerZhiWu)) Then
                    Source(KerRowJ, COAQianDaoSe) = 1
                ElseIf CDate(Source(KerRowJ, COAQianDao)) <= CDate(Standard(8, 4 + KerZhiWu)) Then
                    Source(KerRowJ, COAQianDaoSe) = 2
                Else
                    Source(KerRowJ, COAQianDaoSe) = 3
                    Source(KerRowJ, COAXiaChi) = 1
                End If
            End If
            For i = 1 To SSRmax      '自习表中不含D时才允许优惠
                If InStr(SelfStudyTable(i, 1), Source(KerRowJ, COAXingMing)) > 0 And InStr(SelfStudyTable(i, Source(KerRowJ, COAZhou) + 2), "E") = 0 Then
                    Source(KerRowJ, COAZiXi) = Source(KerRowJ, COAZiXi) & "★"
                    If Len(Source(KerRowJ, COAQianTuiSe)) = 0 And Len(Source(KerRowJ, COAQianTui)) > 0 And IsDate(Source(KerRowJ, COAQianTui)) Then
                        If CDate(Source(KerRowJ, COAQianTui)) > CDate(Standard(9, 6 + KerZhiWu)) Then
                            Source(KerRowJ, COAQianTuiSe) = 1
                        ElseIf CDate(Source(KerRowJ, COAQianTui)) >= CDate(Standard(9, 5 + KerZhiWu)) Then
                            Source(KerRowJ, COAQianTuiSe) = 2
                        Else
                            Source(KerRowJ, COAQianTuiSe) = 3
                            Source(KerRowJ, COAXiaTui) = 1
                        End If
                    End If
                End If
            Next
        End If
    ElseIf InStr(Change(KerRowI, KerColI), "E") > 0 Then
        If Source(KerRowJ, COABanCi) = 2 Then
            If Len(Source(KerRowJ, COAQianTuiSe)) = 0 And Len(Source(KerRowJ, COAQianTui)) > 0 And IsDate(Source(KerRowJ, COAQianTui)) Then
                If CDate(Source(KerRowJ, COAQianTui)) > CDate(Standard(10, 6 + KerZhiWu)) Then
                    Source(KerRowJ, COAQianTuiSe) = 1
                ElseIf CDate(Source(KerRowJ, COAQianTui)) >= CDate(Standard(10, 5 + KerZhiWu)) Then
                    Source(KerRowJ, COAQianTuiSe) = 2
                Else
                    Source(KerRowJ, COAQianTuiSe) = 3
                    Source(KerRowJ, COAXiaTui) = 1
                End If
            End If
        End If
    End If
End Sub

Sub COASelfStudyMOD()
' 修正自习表，因为部分人员有可能换课，这会涉及到自习的的考核
    Dim i, j As Integer
    Dim SSitem As String
' 调出
    For i = 2 To CGRmax
        For j = 2 To SSRmax
'
           SSitem = Empty
           If DateMin <= CDate(Change(i, 3)) And CDate(Change(i, 3)) <= DateMax Then
            If InStr(SelfStudyTable(j, 1), Change(i, 1)) > 0 Then
                For k = 3 To 7
                    If InStr(SelfStudyTable(1, k), Change(i, 5)) > 0 Then
                        If InStr(SelfStudyTable(j, k), Change(i, 6)) > 0 Then
                            If InStr(SelfStudyTable(j, k), "B") > 0 And InStr(Change(i, 6), "B") = 0 Then
                                SSitem = SSitem & "B"
                            End If
                            If InStr(SelfStudyTable(j, k), "C") > 0 And InStr(Change(i, 6), "C") = 0 Then
                                SSitem = SSitem & "C"
                            End If
                            If InStr(SelfStudyTable(j, k), "D") > 0 And InStr(Change(i, 6), "D") = 0 Then
                                SSitem = SSitem & "D"
                            End If
                            If InStr(SelfStudyTable(j, k), "E") > 0 And InStr(Change(i, 6), "E") = 0 Then
                                SSitem = SSitem & "E"
                            End If
                            SelfStudyTable(j, k) = SSitem
                        End If
                    End If
                Next
            End If
           End If
'
           SSitem = Empty
           If DateMin <= CDate(Change(i, 9)) And CDate(Change(i, 9)) <= DateMax Then
            If InStr(SelfStudyTable(j, 1), Change(i, 7)) > 0 Then
                For k = 3 To 7
                    If InStr(SelfStudyTable(1, k), Change(i, 11)) > 0 Then
                        If InStr(SelfStudyTable(j, k), Change(i, 12)) > 0 Then
                            If InStr(SelfStudyTable(j, k), "B") > 0 And InStr(Change(i, 12), "B") = 0 Then
                                SSitem = SSitem & "B"
                            End If
                            If InStr(SelfStudyTable(j, k), "C") > 0 And InStr(Change(i, 12), "C") = 0 Then
                                SSitem = SSitem & "C"
                            End If
                            If InStr(SelfStudyTable(j, k), "D") > 0 And InStr(Change(i, 12), "D") = 0 Then
                                SSitem = SSitem & "D"
                            End If
                            If InStr(SelfStudyTable(j, k), "E") > 0 And InStr(Change(i, 12), "E") = 0 Then
                                SSitem = SSitem & "E"
                            End If
                            SelfStudyTable(j, k) = SSitem
                        End If
                    End If
                Next
            End If
           End If
'
        Next
    Next
' 调入
    For i = 2 To CGRmax
        For j = 2 To SSRmax
'
            If DateMin <= CDate(Change(i, 3)) And CDate(Change(i, 3)) <= DateMax Then
                If InStr(SelfStudyTable(j, 1), Change(i, 7)) > 0 Then
                    For k = 3 To 7
                        If InStr(SelfStudyTable(1, k), Change(i, 5)) > 0 Then
                            If InStr(SelfStudyTable(j, k), Change(i, 6)) > 0 Then
                                MsgBox SelfStudyTable(j, 1) & "已经存在" & Change(i, 6) & "调课无效！"
                            Else
                                SelfStudyTable(j, k) = SelfStudyTable(j, k) & Change(i, 6)
                            End If
                        End If
                    Next
                End If
            End If
'
            If DateMin <= CDate(Change(i, 9)) And CDate(Change(i, 9)) <= DateMax Then
                If InStr(SelfStudyTable(j, 1), Change(i, 1)) > 0 Then
                    For k = 3 To 7
                        If InStr(SelfStudyTable(1, k), Change(i, 11)) > 0 Then
                            If InStr(SelfStudyTable(j, k), Change(i, 12)) > 0 Then
                                MsgBox SelfStudyTable(j, 1) & "已经存在" & Change(i, 12) & "调课无效！"
                            Else
                                SelfStudyTable(j, k) = SelfStudyTable(j, k) & Change(i, 12)
                            End If
                        End If
                    Next
                End If
            End If
'
        Next
    Next
End Sub

Sub COASelfStudyADD()
    Dim KerZhiWu As Integer
    For i = 2 To SSRmax
        For j = 2 To SRowMax
            KerZhiWu = (Source(j, COAZhiWu) - 1) * 5
            If InStr(Source(j, COAXingMing), SelfStudyTable(i, 1)) > 0 And Source(j, COAZhou) < 6 Then
                If InStr(SelfStudyTable(i, Source(j, COAZhou) + 2), "B") > 0 Then  '处理B项
                    If InStr(Source(j, COABanCi), 1) > 0 Then
                        Source(j, COAZiXi) = Source(j, COAZiXi) & "B"
                        If Len(Source(j, COAQianDaoSe)) = 0 And Len(Source(j, COAQianDao)) > 0 And IsDate(Source(j, COAQianDao)) Then
                            If CDate(Source(j, COAQianDao)) < CDate(Standard(5, 3 + KerZhiWu)) Then
                                Source(j, COAQianDaoSe) = 1
                            ElseIf CDate(Source(j, COAQianDao)) <= CDate(Standard(5, 4 + KerZhiWu)) Then
                                Source(j, COAQianDaoSe) = 2
                            Else
                                Source(j, COAQianDaoSe) = 3
                                Source(j, COAShangChi) = 1
                            End If
                        End If
                    ElseIf InStr(Source(j, COABanCi), 2) > 0 And InStr(SelfStudyTable(i, Source(j, COAZhou) + 2), "D") = 0 Then
                        Source(j, COAZiXi) = "★"
                        If Len(Source(j, COAQianDaoSe)) = 0 And Len(Source(j, COAQianDao)) > 0 And IsDate(Source(j, COAQianDao)) Then
                            If CDate(Source(j, COAQianDao)) < CDate(Standard(7, 3 + KerZhiWu)) Then
                                Source(j, COAQianDaoSe) = 1
                            ElseIf CDate(Source(j, COAQianDao)) <= CDate(Standard(7, 4 + KerZhiWu)) Then
                                Source(j, COAQianDaoSe) = 2
                            Else
                                Source(j, COAQianDaoSe) = 3
                                Source(j, COAXiaChi) = 1
                            End If
                        End If
                    End If
                End If
                If InStr(SelfStudyTable(i, Source(j, COAZhou) + 2), "C") > 0 Then '处理C项
                    If InStr(Source(j, COABanCi), 1) > 0 Then
                        Source(j, COAZiXi) = Source(j, COAZiXi) & "C"
                        If Len(Source(j, COAQianTuiSe)) = 0 And Len(Source(j, COAQianTui)) > 0 And IsDate(Source(j, COAQianTui)) Then
                            If CDate(Source(j, COAQianTui)) > CDate(Standard(6, 6 + KerZhiWu)) Then
                                Source(j, COAQianTuiSe) = 1
                            ElseIf CDate(Source(j, COAQianTui)) >= CDate(Standard(6, 5 + KerZhiWu)) Then
                                Source(j, COAQianTuiSe) = 2
                            Else
                                Source(j, COAQianTuiSe) = 3
                                Source(j, COAShangTui) = 1
                            End If
                        End If
                    ElseIf InStr(Source(j, COABanCi), 2) > 0 And InStr(SelfStudyTable(i, Source(j, COAZhou) + 2), "D") = 0 Then
                        Source(j, COAZiXi) = "★"
                        If Len(Source(j, COAQianDaoSe)) = 0 And Len(Source(j, COAQianDao)) > 0 And IsDate(Source(j, COAQianDao)) Then
                            If CDate(Source(j, COAQianDao)) < CDate(Standard(7, 3 + KerZhiWu)) Then
                                Source(j, COAQianDaoSe) = 1
                            ElseIf CDate(Source(j, COAQianDao)) <= CDate(Standard(7, 4 + KerZhiWu)) Then
                                Source(j, COAQianDaoSe) = 2
                            Else
                                Source(j, COAQianDaoSe) = 3
                                Source(j, COAXiaChi) = 1
                            End If
                        End If
                    End If
                End If
                If InStr(SelfStudyTable(i, Source(j, COAZhou) + 2), "D") > 0 Then '处理D项
                    If InStr(Source(j, COABanCi), 2) > 0 Then
                        Source(j, COAZiXi) = Source(j, COAZiXi) & "D"
                        If Len(Source(j, COAQianDaoSe)) = 0 And Len(Source(j, COAQianDao)) > 0 And IsDate(Source(j, COAQianDao)) Then
                            If CDate(Source(j, COAQianDao)) < CDate(Standard(8, 3 + KerZhiWu)) Then
                                Source(j, COAQianDaoSe) = 1
                            ElseIf CDate(Source(j, COAQianDao)) <= CDate(Standard(8, 4 + KerZhiWu)) Then
                                Source(j, COAQianDaoSe) = 2
                            Else
                                Source(j, COAQianDaoSe) = 3
                                Source(j, COAXiaChi) = 1
                            End If
                        End If
                        If InStr(SelfStudyTable(i, Source(j, COAZhou) + 2), "E") = 0 Then
                            If Len(Source(j, COAQianTuiSe)) = 0 And Len(Source(j, COAQianTui)) > 0 And IsDate(Source(j, COAQianTui)) Then
                                If CDate(Source(j, COAQianTui)) > CDate(Standard(9, 6 + KerZhiWu)) Then
                                    Source(j, COAQianTuiSe) = 1
                                ElseIf CDate(Source(j, COAQianTui)) >= CDate(Standard(9, 5 + KerZhiWu)) Then
                                    Source(j, COAQianTuiSe) = 2
                                Else
                                    Source(j, COAQianTuiSe) = 3
                                    Source(j, COAXiaTui) = 1
                                End If
                            End If
                        End If
                    End If
                End If
                If InStr(SelfStudyTable(i, Source(j, COAZhou) + 2), "E") > 0 Then '处理E项
                    If Source(j, COABanCi) = 2 Then
                        Source(j, COAZiXi) = Source(j, COAZiXi) & "E"
                        If Len(Source(j, COAQianTuiSe)) = 0 And Len(Source(j, COAQianTui)) > 0 And IsDate(Source(j, COAQianTui)) Then
                            If CDate(Source(j, COAQianTui)) > CDate(Standard(10, 6 + KerZhiWu)) Then
                                Source(j, COAQianTuiSe) = 1
                            ElseIf CDate(Source(j, COAQianTui)) >= CDate(Standard(10, 5 + KerZhiWu)) Then
                                Source(j, COAQianTuiSe) = 2
                            Else
                                Source(j, COAQianTuiSe) = 3
                                Source(j, COAXiaTui) = 1
                            End If
                        End If
                    End If
                End If
            End If
        Next
    Next
End Sub

Sub COANormalEXE()
    For j = 2 To SRowMax
        Select Case Source(j, COAZhiWu)
            Case Is = 1
                Call COAEXE1(j, 0)
            Case Is = 2
                Call COAEXE2(j, 5)
        End Select
    Next
End Sub
Sub COAGenerateEXE()
    Dim i, j, k As Integer
    S1RowMax = 1: S2RowMax = 1
    For j = 2 To SRowMax
        Select Case Source(j, COAZhiWu)
            Case Is = 1
                S1RowMax = S1RowMax + 1
                For k = 1 To SubColMax
                    Source1(S1RowMax, k) = Source(j, k)
                Next
            Case Is = 2
                S2RowMax = S2RowMax + 1                                                                             '输出职务2（班主任）
                For k = 1 To UBound(Source2, 2)
                    Source2(S2RowMax, k) = Source(j, k)
                Next
        End Select
    Next
' 只输出异常考勤
    Call GenerateBook(Source1, S1RowMax, 12, NameTeacherUN)
    Call GenerateBook(Source2, S2RowMax, 15, NameHeadMasterUN)
' 关闭及退出
    Application.DisplayAlerts = False
    Workbooks.Close                                                                                                 '关闭所有工作薄
    Application.DisplayAlerts = True
    Application.Quit                                                                                                '退出Excel
    Shell "explorer.exe " & OutFolder, vbNormalFocus
End Sub
Sub COAEXE1(KernelRow, KernelCol)
    Dim KernelBanCi, KernelResultCol As Integer
    KernelBanCi = Source(KernelRow, COABanCi) + 1
    KernelResultCol = 3 * (Source(KernelRow, COABanCi) - 1)                                                                                   '考核上午
    If Len(Source(KernelRow, COAQianDaoSe)) = 0 Then
       If Len(Source(KernelRow, COAQianDao)) = 0 Then
           Source(KernelRow, COAQianDaoSe) = 4
           Source(KernelRow, COAQianDao) = "漏签"
           Source(KernelRow, COAShangLou + KernelResultCol) = 1
       ElseIf IsDate(Source(KernelRow, COAQianDao)) Then
            If CDate(Source(KernelRow, COAQianDao)) < CDate(Standard(KernelBanCi, 3 + KernelCol)) Then
                Source(KernelRow, COAQianDaoSe) = 1
            ElseIf CDate(Source(KernelRow, COAQianDao)) <= CDate(Standard(KernelBanCi, 4 + KernelCol)) Then
                Source(KernelRow, COAQianDaoSe) = 2
            Else
                Source(KernelRow, COAQianDaoSe) = 3
                Source(KernelRow, COAShangChi + KernelResultCol) = 1
            End If
       End If
    End If
    If Len(Source(KernelRow, COAQianTuiSe)) = 0 Then
       If Len(Source(KernelRow, COAQianTui)) = 0 Then
           Source(KernelRow, COAQianTuiSe) = 4
           Source(KernelRow, COAQianTui) = "漏签"
           Source(KernelRow, COAShangLou + KernelResultCol) = Source(KernelRow, COAShangLou + KernelResultCol) + 1
       ElseIf IsDate(Source(KernelRow, COAQianTui)) Then
            If CDate(Standard(KernelBanCi, 6 + KernelCol)) < CDate(Source(KernelRow, COAQianTui)) Then
                Source(KernelRow, COAQianTuiSe) = 1
            ElseIf CDate(Standard(KernelBanCi, 5 + KernelCol)) < CDate(Source(KernelRow, COAQianTui)) Then
                Source(KernelRow, COAQianTuiSe) = 2
            Else
                Source(KernelRow, COAQianTuiSe) = 3
                Source(KernelRow, COAShangTui + KernelResultCol) = 1
            End If
       End If
    End If
End Sub
Sub COAEXE2(KernelRow, KernelCol)
    If Source(KernelRow, COAZhou) = BZBeginNum And Source(KernelRow, COABanCi) = 1 Then                                            '只在上午执行重置，且一周只重置一次
        BZSNum = 2: BZXNum = 2: BZWNum = 2
    ElseIf Source(KernelRow, COAZhou) = BZBeginNumX And Source(KernelRow, COABanCi) = 1 Then
        BZSNumX = 1: BZXNumX = 1: BZWNumX = 1
    End If
    Dim KernelBanCi, KernelResultCol As Integer
    KernelBanCi = Source(KernelRow, COABanCi) + 1
    KernelResultCol = 3 * (Source(KernelRow, COABanCi) - 1)                                                                         '考核上午
    If Len(Source(KernelRow, COAQianDaoSe)) = 0 Then
       If Len(Source(KernelRow, COAQianDao)) = 0 And Source(KernelRow, COABanCi) <> 3 Then                                          '统计上午和下午漏签
            Source(KernelRow, COAQianDaoSe) = 4
            Source(KernelRow, COAQianDao) = "漏签"
            If Source(KernelRow, COAZhou) = 6 Or Source(KernelRow, COAZhou) = 7 Then
                If Source(KernelRow, COABanCi) = 1 Then                                                                             '上午签到漏签
                    If BZSNumX > 0 Then
                        BZSNumX = BZSNumX - 1
                        Source(KernelRow, COAQianDaoSe) = 6
                    Else
                        Source(KernelRow, COAShangLou + KernelResultCol) = 1
                    End If
                ElseIf Source(KernelRow, COABanCi) = 2 Then                                                                          '下午签到漏签
                    If BZXNumX > 0 Then
                        BZXNumX = BZXNumX - 1
                        Source(KernelRow, COAQianDaoSe) = 6
                    Else
                        Source(KernelRow, COAShangLou + KernelResultCol) = 1
                    End If
                End If
            Else
                Source(KernelRow, COAShangLou + KernelResultCol) = 1
            End If
       ElseIf Source(KernelRow, COABanCi) <> 3 Then                                                                               '统计上午和下午迟到
           If CDate(Source(KernelRow, COAQianDao)) <= CDate(Standard(KernelBanCi, 3 + KernelCol)) Then
               Source(KernelRow, COAQianDaoSe) = 1
           ElseIf CDate(Source(KernelRow, COAQianDao)) <= CDate(Standard(KernelBanCi, 4 + KernelCol)) Then
               Source(KernelRow, COAQianDaoSe) = 2
           Else
               Source(KernelRow, COAQianDaoSe) = 3
               If Source(KernelRow, COABanCi) = 1 Then                                                                            '上午
                    If Source(KernelRow, COAZhou) = 6 Or Source(KernelRow, COAZhou) = 7 Then
                        If BZSNumX > 0 Then
                            If CDate(Source(KernelRow, COAQianDao)) > CDate(Standard(KernelBanCi, 4)) Then
                                Source(KernelRow, COAShangChi + KernelResultCol) = 1
                            Else
                                BZSNumX = BZSNumX - 1
                                Source(KernelRow, COAQianDaoSe) = 6
                            End If
                        Else
                            Source(KernelRow, COAShangChi + KernelResultCol) = 1
                        End If
                    Else
                        If BZSNum > 0 Then
                            If CDate(Source(KernelRow, COAQianDao)) > CDate(Standard(KernelBanCi, 4)) Then
                                Source(KernelRow, COAShangChi + KernelResultCol) = 1
                            Else
                                BZSNum = BZSNum - 1
                                Source(KernelRow, COAQianDaoSe) = 6
                            End If
                        Else
                            Source(KernelRow, COAShangChi + KernelResultCol) = 1
                        End If
                    End If
               ElseIf Source(KernelRow, COABanCi) = 2 Then                                                                      '下午
                    If Source(KernelRow, COAZhou) = 6 Or Source(KernelRow, COAZhou) = 7 Then
                        If BZXNumX > 0 Then
                            If CDate(Source(KernelRow, COAQianDao)) > CDate(Standard(KernelBanCi, 5)) Then
                                Source(KernelRow, COAShangChi + KernelResultCol) = 1
                            Else
                                BZXNumX = BZXNumX - 1
                                Source(KernelRow, COAQianDaoSe) = 6
                            End If
                        Else
                            Source(KernelRow, COAShangChi + KernelResultCol) = 1
                        End If
                    Else
                        If BZXNum > 0 Then
                            If CDate(Source(KernelRow, COAQianDao)) > CDate(Standard(KernelBanCi, 5)) Then
                                Source(KernelRow, COAShangChi + KernelResultCol) = 1
                            Else
                                BZXNum = BZXNum - 1
                                Source(KernelRow, COAQianDaoSe) = 6
                            End If
                        Else
                            Source(KernelRow, COAShangChi + KernelResultCol) = 1
                        End If
                    End If
               End If
           End If
       End If
    End If
    If Len(Source(KernelRow, COAQianTuiSe)) = 0 Then
'' debug
       If Len(Source(KernelRow, COAQianTui)) = 0 And Source(KernelRow, COABanCi) = 3 Then
          If Len(Source(KernelRow, COAQianDao)) > 0 And IsDate(Source(KernelRow, COAQianDao)) Then
            Source(KernelRow, COAQianTui) = Source(KernelRow, COAQianDao)
            Source(KernelRow, COAQianDao) = Empty
          ElseIf Len(Source(KernelRow - 1, COAQianTui)) > 0 And IsDate(Source(KernelRow - 1, COAQianTui)) Then
            Source(KernelRow, COAQianTui) = Source(KernelRow - 1, COAQianTui)
            Source(KernelRow - 1, COAQianTui) = Empty
          End If
       End If
''
       If Len(Source(KernelRow, COAQianTui)) = 0 And Source(KernelRow, COABanCi) <> 2 Then
           Source(KernelRow, COAQianTuiSe) = 4
                Source(KernelRow, COAQianTui) = "漏签"
           If Source(KernelRow, COAZhou) = 6 Or Source(KernelRow, COAZhou) = 7 Then
                If Source(KernelRow, COABanCi) = 1 Then
                    Source(KernelRow, COAQianTuiSe) = 6                                                                         '不统计周末的签退，以蓝色标出此项
                ElseIf Source(KernelRow, COABanCi) = 3 Then
                    If BZWNumX > 0 Then
                        BZWNumX = BZWNumX - 1
                        Source(KernelRow, COAQianTuiSe) = 6
                    Else
                        Source(KernelRow, COAShangLou + KernelResultCol) = Source(KernelRow, COAShangLou) + 1
                    End If
                End If
           Else
                Source(KernelRow, COAShangLou + KernelResultCol) = Source(KernelRow, COAShangLou) + 1
           End If
       ElseIf Source(KernelRow, COABanCi) <> 2 Then
           If CDate(Standard(KernelBanCi, 6 + KernelCol)) < CDate(Source(KernelRow, COAQianTui)) Then
               Source(KernelRow, COAQianTuiSe) = 1
           ElseIf CDate(Standard(KernelBanCi, 5 + KernelCol)) < CDate(Source(KernelRow, COAQianTui)) Then
               Source(KernelRow, COAQianTuiSe) = 2
           Else
               Source(KernelRow, COAQianTuiSe) = 3
               If Source(KernelRow, COAZhou) = 6 Or Source(KernelRow, COAZhou) = 7 Then
                    If BZWNumX > 0 And Source(KernelRow, COABanCi) = 3 Then
                     BZWNumX = BZWNumX - 1
                     Source(KernelRow, COAQianTuiSe) = 6
                    ElseIf Source(KernelRow, COABanCi) = 3 Then
                     Source(KernelRow, COAShangTui + KernelResultCol) = 1
                    End If
               Else
                    If BZWNum > 0 Then
                     BZWNum = BZWNum - 1
                     Source(KernelRow, COAQianTuiSe) = 6
                    Else
                     Source(KernelRow, COAShangTui + KernelResultCol) = 1
                    End If
               End If
           End If
       End If
    End If
End Sub
'
Sub GenerateBook(InSource, InRmax, InCmax, InName)
    Dim OutBook As Workbook
    Dim OutSource(1 To RowMax, 1 To SubColMax) As Variant
    Dim i, j, k, m, n, p, q, OutMax As Integer
    Dim OutZhou As String
    Application.SheetsInNewWorkbook = 1                                                                     '设置1个Sheet
    Set OutBook = Workbooks.Add
    Application.DisplayAlerts = False
    OutBook.SaveAs Filename:=OutFolder & "\" & InName & OutFileFix & ".xlsx"
    Sheets(1).Name = InName & OutFileFix
    Sheets(InName & OutFileFix).Columns("E:F").NumberFormatLocal = TimeFormat                               '设置时间格式
'汇总统计结果并过滤合格人员
    For i = 1 To InCmax
        OutSource(1, i) = Source(1, i)
    Next
    OutSource(1, InCmax + 1) = "总数"
    k = 0: p = 1
    For i = 2 To InRmax
        If InSource(i, COAXingMing) = InSource(i + 1, COAXingMing) Then
            k = k + 1
        Else
            For j = i - k To i - 1
                For m = COAShangChi To InCmax
                   InSource(i, m) = InSource(i, m) + InSource(j, m)
                   InSource(j, m) = Empty
                Next
            Next
            For m = COAShangChi To InCmax
                   InSource(i, InCmax + 1) = InSource(i, InCmax + 1) + InSource(i, m)
            Next
            If InSource(i, InCmax + 1) > 0 Then
                For j = i - k To i
                    p = p + 1
                    For q = 1 To SubColMax
                        OutSource(p, q) = InSource(j, q)
                    Next
                    If OutSource(p, COAQianDao) = 0 Then
                        OutSource(p, COAQianDao) = Empty
                    End If
                    If OutSource(p, COAQianTui) = 0 Then
                        OutSource(p, COAQianTui) = Empty
                    End If
                    Select Case OutSource(p, COAZhou)
                        Case Is = 1
                            OutZhou = "一"
                        Case Is = 2
                            OutZhou = "二"
                        Case Is = 3
                            OutZhou = "三"
                        Case Is = 4
                            OutZhou = "四"
                        Case Is = 5
                            OutZhou = "五"
                        Case Is = 6
                            OutZhou = "六"
                        Case Is = 7
                            OutZhou = "日"
                    End Select
                    OutSource(p, COARiQi) = Format(OutSource(p, COARiQi), DateFormat) & "(" & OutZhou & ")"
                    OutSource(p, COAZhou) = OutSource(p, COAZiXi)                                           '将周替换为自习信息
                    Select Case OutSource(p, COABanCi)                                                      '恢复排班号为上下晚
                        Case Is = 1
                            OutSource(p, COABanCi) = "上午"
                        Case Is = 2
                            OutSource(p, COABanCi) = "下午"
                        Case Is = 3
                            OutSource(p, COABanCi) = "晚上"
                    End Select
                Next
                For q = COAShangChi To InCmax
                    OutSource(p - 1, q) = Source(1, q)                                                      '统计结果上一行加入标题
                Next
                OutSource(p - 1, InCmax + 1) = "总数"
            End If
            k = 0
        End If
    Next
    OutMax = p
'输出统计结果非零人员
    Range(Cells(1, 1), Cells(OutMax, InCmax + 1)) = OutSource                                               '将数组写入新建表格中
    Call COAFormat(Range(Cells(1, 1), Cells(OutMax, InCmax + 1)))                                           '整体格式化
    Range(Cells(1, 1), Cells(1, InCmax + 1)).Font.Bold = True                                               '标题加黑
'据上色码上色
    For i = 2 To OutMax
        Select Case CInt(OutSource(i, COAQianDaoSe))
            Case Is = 1
                Call COAColor(Cells(i, COAQianDao), 4, 10)                                                           '草绿底+深绿字
            Case Is = 2
                Call COAColor(Cells(i, COAQianDao), 6, 10)                                                           '黄色底+深绿字
            Case Is = 3
                Call COAColor(Cells(i, COAQianDao), 3, 2)                                                            '深红底+白色字
                Call COAColor(Cells(i, COABanCi), 3, 2)
            Case Is = 4
                Call COAColor(Cells(i, COAQianDao), 3, 2)                                                            '大红底+白色字
                Call COAColor(Cells(i, COABanCi), 3, 2)
            Case Is = 5
                Call COAColor(Cells(i, COAQianDao), 10, 6)                                                           '深绿底+黄色字
            Case Is = 6
                Call COAColor(Cells(i, COAQianDao), 5, 2)                                                            '深紫色+白色字（弹性考勤去除项)
                Call COAColor(Cells(i, COABanCi), 5, 2)
        End Select
        Select Case CInt(OutSource(i, COAQianTuiSe))
            Case Is = 1
                Call COAColor(Cells(i, COAQianTui), 4, 10)                                                           '草绿底+深绿字
            Case Is = 2
                Call COAColor(Cells(i, COAQianTui), 6, 10)                                                           '黄色底+深绿字
            Case Is = 3
                Call COAColor(Cells(i, COAQianTui), 3, 2)                                                            '深红底+白色字
                Call COAColor(Cells(i, COABanCi), 3, 2)
            Case Is = 4
                Call COAColor(Cells(i, COAQianTui), 3, 2)                                                            '大红底+白色字
                Call COAColor(Cells(i, COABanCi), 3, 2)
            Case Is = 5
                Call COAColor(Cells(i, COAQianTui), 10, 6)                                                           '深绿底+黄色字
            Case Is = 6
                Call COAColor(Cells(i, COAQianTui), 5, 2)                                                            '深紫色+白色字（弹性考勤去除项)
                Call COAColor(Cells(i, COABanCi), 5, 2)
        End Select
        If Len(OutSource(i, COAZhou)) > 0 Then
            Call COAColor(Cells(i, COAZhou), 37, 51)
        End If
    Next
    k = 0: p = 0
    For i = 2 To OutMax
        If OutSource(i, COAXingMing) = OutSource(i + 1, COAXingMing) Then
            k = k + 1
        Else
            For m = COAShangChi To InCmax + 1
                If OutSource(i, m) > 0 Then                                                                 '统计区上色
                    Call COAColor(Cells(i, m), 3, 2)                                                        '深红底+白色字
                ElseIf OutSource(i, m) = 0 Then
                    Call COAColor(Cells(i, m), 4, 10)                                                       '草绿底+深绿字
                End If
            Next
            Range(Cells(i - k, 1), Cells(i, 1)).Merge                                                       '合并姓名
            For q = COAShangChi To InCmax Step 3
                Range(Cells(i - k, q), Cells(i - 2, q + 2)).Merge                                           '合并统计区
            Next
            Cells(i - k, COAShangChi) = OutSource(i, COAHuanKe)                                                         '加入调课信息
            k = 0
        End If
        If OutSource(i, COARiQi) = OutSource(i + 1, COARiQi) Then
            p = p + 1
        Else
            Range(Cells(i - p, COARiQi), Cells(i, COARiQi)).Merge                                                       '合并日期
            p = 0
        End If
    Next
'冻结首行，方便在电脑上对照查看
    Cells(1, 1).Select                                                                                      '唯一选定单元格的地方
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    ActiveWorkbook.Close savechanges:=True
    Application.DisplayAlerts = True
    Set OutBook = Nothing                                                                                   '取消OutBook
End Sub
'
Sub COAColor(COARange, InteriorColor, FontColor)
    With COARange
        .Interior.ColorIndex = InteriorColor
        .Font.ColorIndex = FontColor
    End With
End Sub
'
Sub COAFormat(COAFRange)
    COAFRange.Rows.AutoFit
    COAFRange.Columns.AutoFit
    With COAFRange
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With COAFRange.Font
        .Name = NameFont
        .Size = 11
    End With
    COAFRange.Borders(xlDiagonalDown).LineStyle = xlNone
    COAFRange.Borders(xlDiagonalUp).LineStyle = xlNone
    With COAFRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With COAFRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With COAFRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With COAFRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With COAFRange.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With COAFRange.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub
'
Sub GetLeaveBook()
    Dim GBeginDate As Date
    Dim GEndDate As Date
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 校准考勤时间为标准的周一到周日，因为年级需要周五提交上周五到本周四的考勤，而学校要求提交周一到周日的考勤报表              '
'                                                                                                                           '
' 注意：由于这个时间节点的修正，如果在周五生成请假表时有人还没有上交周五的假条，则需要将周五的假条手工加入到学校假条和总表  '
'                                                                                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    GEndDate = DateMax
    Do While Weekday(GEndDate, 2) < 7
        GEndDate = GEndDate + 1
    Loop
    GBeginDate = GEndDate - 6
''根据请假表生成上交的请假表leave
    ReDim PreLeave(1 To 6000, 1 To 7) As Variant
    PreLeave(1, 1) = Format(GBeginDate, "m" & "月" & "d" & "日") & "-" & Format(GEndDate, "m" & "月" & "d" & "日") & "考勤情况"
    PreLeave(2, 1) = "年级/科室"
    PreLeave(2, 2) = "时间"
    PreLeave(2, 3) = "姓名"
    PreLeave(2, 4) = "事由"
    k = 2
    For i = 2 To HRmax
        If IsDate(Holiday(i, 2)) And IsDate(Holiday(i, 3)) Then             '先判断为日期后再执行请假操作
            If CDate(GBeginDate) <= CDate(Holiday(i, 3)) And CDate(Holiday(i, 2)) <= CDate(GEndDate) And InStr(Holiday(i, 1), "*") = 0 Then
                k = k + 1
                PreLeave(k, 1) = Nianji
'' 取得请假的时间间隔
                If CDate(Holiday(i, 2)) <= CDate(GBeginDate) Then           '取得起始时间
                    PreLeave(k, 2) = CDate(GBeginDate)
                Else
                    PreLeave(k, 2) = CDate(Holiday(i, 2))
                End If
                If CDate(GEndDate) <= CDate(Holiday(i, 3)) Then             '取得终止时间
                    PreLeave(k, 3) = CDate(GEndDate)
                Else
                    PreLeave(k, 3) = CDate(Holiday(i, 3))
                End If
'' 生成预处理请假表（上交学校）
                PreLeave(k, 6) = Holiday(i, 1)
                If Holiday(i, 4) > 0 Then
                    PreLeave(k, 7) = Holiday(i, 4)
                    PreLeave(k, 4) = PreLeave(k, 3) - PreLeave(k, 2) + 1
                Else
                    PreLeave(k, 4) = 0.5 * (PreLeave(k, 3) - PreLeave(k, 2) + 1)
                    If Holiday(i, 5) > 0 Or Holiday(i, 6) > 0 Then
                        PreLeave(k, 5) = "上午"
                    End If
                    If Holiday(i, 7) > 0 Or Holiday(i, 8) > 0 Then
                        PreLeave(k, 5) = "下午"
                    End If
                    If Holiday(i, 9) > 0 Or Holiday(i, 10) > 0 Then
                        PreLeave(k, 5) = "晚上"
                    End If
                    If Holiday(i, 6) > 0 Then
                        PreLeave(k, 7) = Holiday(i, 5)
                    ElseIf Holiday(i, 7) > 0 Then
                        PreLeave(k, 7) = Holiday(i, 6)
                    ElseIf Holiday(i, 8) > 0 Then
                        PreLeave(k, 7) = Holiday(i, 7)
                    ElseIf Holiday(i, 9) > 0 Then
                        PreLeave(k, 7) = Holiday(i, 8)
                    ElseIf Holiday(i, 10) > 0 Then
                        PreLeave(k, 7) = Holiday(i, 9)
                    ElseIf Holiday(i, 11) > 0 Then
                        PreLeave(k, 7) = Holiday(i, 10)
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
  NameLeave = Nianji & Format(GBeginDate, "m" & "月" & "d" & "日") & "-" & Format(GEndDate, "m" & "月" & "d" & "日") & "考勤"
  Call OutToLeave(Leave, l, UBound(Leave, 2), OutFolder, NameLeave)
End Sub

Sub OutToLeave(LeaveSource, InRmax, InCmax, OutLeaveFolder, InName)
    Dim OutBook As Workbook
    Application.SheetsInNewWorkbook = 1                                                         '设置1个Sheet
    Set OutBook = Workbooks.Add
    ActiveWindow.FreezePanes = False                                                            '禁止冻结窗口
    Application.DisplayAlerts = False
        OutBook.SaveAs Filename:=OutLeaveFolder & "\" & InName & ".xlsx"
        Sheets(1).Name = InName
        Range(Cells(1, 1), Cells(InRmax, InCmax)) = LeaveSource                                 '调入请假信息
        Call COAFormat(Range(Cells(1, 1), Cells(InRmax, InCmax)))
        Range(Cells(1, 1), Cells(1, InCmax)).Merge
        Range(Cells(1, 1), Cells(1, InCmax)).Font.Size = 20
        Range(Cells(2, 1), Cells(2, InCmax)).Font.Size = 14
    ActiveWorkbook.Close savechanges:=True
    Application.DisplayAlerts = True
    Set OutBook = Nothing                                                                       '取消OutBook
End Sub




