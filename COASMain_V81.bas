Attribute VB_Name = "COASMain_V81"
' ��Ŀ��COASMain
' �汾��V81
' ���ߣ�����
' ��λ��ɽ��ʡƽԭ�ص�һ��ѧ
' ���䣺fengzhenhua@outlook.com
' ���ͣ�https://fengzhenhua-vip.github.io
' ��ҳ��https://github.com/fengzhenhua-vip
' ��Ȩ��2021��7��13��--2021��9��23��
' ��־������ϵͳ��7�棬�ǴӶ������ʵ�ֵģ��ۺ�V68֮ǰ�Ĺ�����������һЩ����������ʵ��ȫ�꼶����5������ɡ����ص������
'       1. ����ְ��ţ���ʦ1��������2
'       2. ������ƺţ�Source �����У���20��֮���ʾ���ƺţ�21ǩ��ɫ��22ǩ��ɫ��23ְ��ţ�24������Ϣ��25��ϰ���
'       3. �����κţ�1���磬2���磬3����
'       4. �����ʱ��DuiShiA(������),DuiShiB(����)
'       5. ��������ţ�����ְ��������Ӧ������Source1,Source2
'       6. ���뿼������ţ�1������2���棬3Υ�棬4©ǩ��5��Ϣ��ע����ٵȣ�
'       7. ����޶ȵļ���ƥ�����������޶ȼ��ٵ�Ԫ��ѡ����ֻ��1����
'       8. ��������ɫ�������͵�Ԫ��ϲ�����
'       9. ���ڴӶ�����ƣ����Խṹ������������չ��ͬʱ�ﵽȫ�꼶����5������ɣ���ʵ�򿪺͹ر��ļ�ռ���˴����ʱ��
'       10. �������¹��������ܵĴ��������������ֱ�������汾��V70
' ��־���Ż����������룬�����ȶ��Լ����ܣ������汾��V71
' ��־�������ļ�ʵ�ָ���ֹͣ��������ȥ����ؿ��ڼ�¼���ܣ���Ҫ�޸���GetSourceģ��
'       ��������ж����óɱ�������ʱ���䶯˳�򣬿��ǵ�������������Ƶ�Ҫ���������������������鷳�Ĳ�����ʹ����Щ������
'       ���������Source,Source1,Source2,DuiShiA,DuiShiB,InSource,OutSource��GenerateBook�е���ɫ���֣������汾��V72
' ��־��������GetHoliday,��Ϊ����Ҫְ����жϣ�����ȥ�����ⲿ���жϣ������汾��V73
' ��־���޸��������ļ����ڱ�׼ʱ��Sheet��ȥ����10CT��51CT�����ã���Ϊ��10ST��51STֱ�Ӽ��㣬���ø���������ͬʱ��V73�����˸���
'       �����ɶ�ʱ��ʱ�����������D�����µ�����������Ż�39���ӵ��жϡ������汾��V74
' ��־�����������ļ��汾ƥ���⣬�˰�֮�������ļ����޸�Ϊ�������ļ�V75.xlsm ����ʽ,�����COASMain ,�����汾V75
' ��־���Ż��˿��ں��Ĵ��룬���ӷ��㼯��ά�������Ӱ����ε��Կ��ڲ�����ɫ����������汾V76
' ��־���޸����жϽ�ʦ�Ͱ�������ְ��Ϊ׼���������Ű��ж�2021/9/3
' ��־���޸���ϰ�����ɶ�ʱ��B��bug,2021/9/9 �������汾��V77
' ��־���޸���ϰ������У׼�������ε��Կ���bug,���������ٽ�һ�����ƣ����ǿ��ǵ����п���ռ���Ҹ����ʱ�䣬���Գ���ʵ�ֺ��ټƻ��������졣�˴��޸��Ķ��϶࣬�����汾��V78
' ��־�������˿��ڷ�ʽ��ʹ֮�����������ã����˸���׼ȷ�������汾��ΪV80
' ��־���޸�����ϱ�����Զ�����2021/9/23
' ��־���޸�������һ��ʮһ���Զ��л�bug,�����α��޶���һ������ʱ���Բ�����չ���ܣ������汾��V81
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
'    Call COAChangeADD     ' ���ڲ��ƻ���չ�����򣬹ʽ������޶���һ���ڵĿ��ڣ����ǲ��ٴ����ӵ���������������Ļ���
    Call COASelfStudyMOD   ' �������޶���һ���ڣ����Խ����α�ֱ�ӵ�����ϰ���������������״�����ʱ����
    Call COASelfStudyADD
    Call COAChangeADD
    Call COANormalEXE
    Call COARecorrectBAC
    Call COAGenerateEXE
    Application.ScreenUpdating = True
End Sub

Sub COAConfigSet()
 '   ConfigPath = "D:\����ϵͳ"
    ConfigPath = CreateObject("WScript.Shell").specialfolders("Desktop") & "\����ϵͳ"                          'Ĭ������Ϊ����
    ConfigFolder = ConfigPath & "\" & "����ϵͳ����"
    ConfigFile = ConfigFolder & "\" & "��������V" & COAVersion & ".xlsm"                                        '������ǿ���Զ�У׼���ܵ������ļ�
    Set SFO = CreateObject("Scripting.FileSystemObject")                                                        '��SFOΪ�ļ��ж������
    If SFO.FileExists(ConfigFile) = False Then
        MsgBox "�����ļ��汾�뵱ǰϵͳ��������ʹ�ã�" & ConfigFolder & "\��������V" & COAVersion & ".xlsm"
        End
    Else
        OutPath = ConfigPath & "\" & Format(Now, "yyyy" & "��") & "����"
        DateFormat = "m""��""d""��"";@"
        TimeFormat = "h:mm;@"
        NameFont = "����"
        NameHeadMasterUN = "�쳣������"
        NameTeacherUN = "�쳣��ʦ"
        ConfigSheet1 = "ֹͣ����"
        ConfigSheet2 = "��ٱ�"
        OriginalSheet1 = "��ϰ����"                                                                                 '����4��רΪУ׼������
        OriginalSheet3 = "����У׼"
        OriginalSheet4 = "���α�"
        OriginalSheet5 = "��ʦ����"
        WuYi = Format(Now, "yyyy") & "/5/1"
        ShiYi = Format(Now, "yyyy") & "/10/1"
        COAXingMing = 1: COARiQi = 2: COAZhou = 4: COABanCi = 3: COAQianDao = 5: COAQianTui = 6
        COAShangChi = 7: COAShangTui = COAShangChi + 1: COAShangLou = COAShangChi + 2                               'ȷ��ͳ�������硢���硢��������
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
    ' ͳһ�������ñ�
        Set ConfigBook = GetObject(ConfigFile)
        ViRmax = ConfigBook.Sheets(ConfigSheet1).Cells(RowMax, 1).End(xlUp).Row
        ViCmax = ConfigBook.Sheets(ConfigSheet1).Cells(1, ColMax).End(xlToLeft).Column
        VipSource = ConfigBook.Sheets(ConfigSheet1).Range(ConfigBook.Sheets(ConfigSheet1).Cells(1, 1), ConfigBook.Sheets(ConfigSheet1).Cells(ViRmax, ViCmax))
    ' ��ȡ��ʱ��A
        CGRmax = ConfigBook.Sheets(OriginalSheet4).Cells(RowMax, 1).End(xlUp).Row                                                                    '���뻻�α�
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
    Nianji = "��������"
End Sub

Sub GetSource()
' ȡ��ԭʼ������������
    DeRmax = Sheets(1).Cells(RowMax, 1).End(xlUp).Row
    DeCmax = 13
' ���ɱ�׼ͨ�ÿ��������ʽ
    Dim DeLiSource As Variant
    Dim i, j, k, m As Integer
    DeLiSource = Sheets(1).Range(Cells(1, 3), Cells(DeRmax, DeCmax)).Value
    j = 1
    Source(j, COAXingMing) = "����": Source(j, COARiQi) = "����(��)": Source(j, COABanCi) = "���"
    Source(j, COAZhou) = "��ϰ": Source(j, COAQianDao) = "ǩ��": Source(j, COAQianTui) = "ǩ��"
    Source(j, COAShangChi) = "�ϳ�": Source(j, COAShangTui) = "����": Source(j, COAShangLou) = "��©"
    Source(j, COAXiaChi) = "�³�": Source(j, COAXiaTui) = "����": Source(j, COAXiaLou) = "��©"
    Source(j, COAWanChi) = "���": Source(j, COAWanTui) = "����": Source(j, COAWanLou) = "��©"
    DateMin = CDate(DeLiSource(3, 4)): DateMax = DateMin
    For i = 3 To DeRmax                                                                                      'DeLi������ǰ���кϲ��������Ե�2���ǿյĲ�������
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
        If m = 0 And (InStr(DeLiSource(i, 3), "������") > 0 Or (InStr(DeLiSource(i, 3), "������") = 0 And InStr(DeLiSource(i, 5), "��") = 0 And InStr(DeLiSource(i, 5), "��") = 0)) Then
            j = j + 1
            Source(j, COAXingMing) = VBA.Trim(Mid(DeLiSource(i, 1), 1, 1)) & VBA.Trim(Mid(DeLiSource(i, 1), 2))   '�򵥺�ȥ�����������м�Ŀո�,����ʹ��ͳһ�Ŀո�����
            Source(j, COARiQi) = CDate(DeLiSource(i, 4))
            If DateMin > Source(j, COARiQi) Then
                DateMin = Source(j, COARiQi)
            End If
            If DateMax < Source(j, COARiQi) Then
                DateMax = Source(j, COARiQi)
            End If
            If InStr(DeLiSource(i, 7), "����") > 0 Then
                Source(j, COABanCi) = 1
            ElseIf InStr(DeLiSource(i, 7), "����") > 0 Then
                Source(j, COABanCi) = 2
            ElseIf InStr(DeLiSource(i, 7), "����") > 0 Then
                Source(j, COABanCi) = 3
            End If
            If InStr(DeLiSource(i, 3), "������") > 0 Then
                Source(j, COAZhiWu) = 2
            ElseIf Len(DeLiSource(i, 7)) > 0 Then   ' �����м�������Ա������û�п����Ű࣬����Ϊ�գ��˲�Ӧ������
                Source(j, COAZhiWu) = 1
            End If
            Select Case DeLiSource(i, 5)
                Case Is = "һ"
                    Source(j, COAZhou) = 1
                Case Is = "��"
                    Source(j, COAZhou) = 2
                Case Is = "��"
                    Source(j, COAZhou) = 3
                Case Is = "��"
                    Source(j, COAZhou) = 4
                Case Is = "��"
                    Source(j, COAZhou) = 5
                Case Is = "��"
                    Source(j, COAZhou) = 6
                Case Is = "��"
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
' ���ð����ε��Կ�����ʼ��debug
    If Source(2, COAZhou) = 6 Or Source(2, COAZhou) = 6 Then
       BZBeginNum = 1: BZBeginNumX = Source(2, COAZhou)
    Else
       BZBeginNum = Source(2, COAZhou): BZBeginNumX = 6
    End If
' ����Ŀ��Excel�ļ��ĺ�׺������ļ���
    OutFileFix = "��" & Format(DateMin, "yyyy" & "��" & "m" & "��" & "d" & "��") & "-" & Format(DateMax, "yyyy" & "��" & "m" & "��" & "d" & "��") & "��"
    OutFolder = OutPath & "\" & Format(DateMax, "m" & "��" & "d" & "��") & "��ʽ�ϱ�"
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
Sub GetQingJia() '������������������б�
' ������ٵ�У׼��HolidayA,HolidayB�����߾��ޱ��⣬�ӵ�1���������Ч���ݣ��������ɿ���ͳ����Ϣ
    HARowMax = 0: HBRowMax = 0
    For i = 2 To HRmax
' ȡ�ý�ʦ����Ͷ�Ӧ����
       If InStr(Holiday(i, 1), "��") > 0 Or InStr(Holiday(i, 1), "������") > 0 Or InStr(Holiday(i, 1), "������") > 0 Then
           For j = 1 To TGCmax Step TeGrStep
               If TeacherGroup(1, j) = Holiday(i, 1) Then
                   GroupColum = j
               End If
           Next
           GroupRow = ConfigBook.Sheets(OriginalSheet5).Cells(RowMax, GroupColum).End(xlUp).Row
       End If
'��ȡ���ڸ�ʽ���HolidayA
       If IsDate(Holiday(i, 2)) Then
               If InStr(Holiday(i, 1), "��") > 0 Or InStr(Holiday(i, 1), "������") > 0 Or InStr(Holiday(i, 1), "������") > 0 Then
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
                                            If InStr(Holiday(1, k), "����") > 0 Then
                                                HolidayA(HARowMax, COABanCi) = 1
                                            ElseIf InStr(Holiday(1, k), "����") > 0 Then
                                                HolidayA(HARowMax, COABanCi) = 2
                                            ElseIf InStr(Holiday(1, k), "����") > 0 Then
                                                HolidayA(HARowMax, COABanCi) = 3
                                            Else
                                                MsgBox "��γ���"
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
                                    If InStr(Holiday(1, k), "����") > 0 Then
                                       HolidayA(HARowMax, COABanCi) = 1
                                    ElseIf InStr(Holiday(1, k), "����") > 0 Then
                                       HolidayA(HARowMax, COABanCi) = 2
                                    ElseIf InStr(Holiday(1, k), "����") > 0 Then
                                       HolidayA(HARowMax, COABanCi) = 3
                                    Else
                                       MsgBox "��γ���"
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
' ��ȡ�����ڸ�ʽ��ٱ�HolidayB
           For j = 1 To 7
            If Holiday(i, 2) = Choose(j, "һ", "��", "��", "��", "��", "��", "��") Then
               WeekX = j
            End If
           Next
           For j = 1 To 7
            If Holiday(i, 3) = Choose(j, "һ", "��", "��", "��", "��", "��", "��") Then
               WeekY = j
            End If
           Next
           If InStr(Holiday(i, 1), "��") > 0 Or InStr(Holiday(i, 1), "������") > 0 Or InStr(Holiday(i, 1), "������") > 0 Then
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
                                    If InStr(Holiday(1, k), "����") > 0 Then
                                        HolidayB(HBRowMax, COABanCi) = 1
                                     ElseIf InStr(Holiday(1, k), "����") > 0 Then
                                        HolidayB(HBRowMax, COABanCi) = 2
                                     ElseIf InStr(Holiday(1, k), "����") > 0 Then
                                        HolidayB(HBRowMax, COABanCi) = 3
                                     Else
                                        MsgBox "��γ���"
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
                               If InStr(Holiday(1, k), "����") > 0 Then
                                   HolidayB(HBRowMax, COABanCi) = 1
                               ElseIf InStr(Holiday(1, k), "����") > 0 Then
                                   HolidayB(HBRowMax, COABanCi) = 2
                               ElseIf InStr(Holiday(1, k), "����") > 0 Then
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
Sub COAHolidayADD() '��������Ϣ
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
Sub COARecorrectADD() '����У׼ʱ��
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
Sub COARecorrectBAC() '�ָ�����У׼ʱ��
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
                        Source(j, COAHuanKe) = "����" & Change(i, 1) & Change(i, 6) & "����" & Change(i, 7) & Change(i, 12)
                    End If
                End If
            End If
            If DateMin <= CDate(Change(i, 9)) And CDate(Change(i, 9)) <= DateMax Then
                If InStr(Source(j, COAXingMing), Change(i, 1)) > 0 And InStr(Source(j, COARiQi), Change(i, 9)) > 0 And InStr(Source(j, COAZhou), Change(i, 11)) > 0 Then
                    Call COAChangeKernel(i, 12, j)
                    If InStr(Source(j, COABanCi), Change(i, 10)) > 0 Then
                        Source(j, COAHuanKe) = "����" & Change(i, 7) & Change(i, 12) & "����" & Change(i, 1) & Change(i, 6)
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
            For i = 1 To SSRmax      '��ϰ���в���Dʱ�������Ż�
                If InStr(SelfStudyTable(i, 1), Source(KerRowJ, COAXingMing)) > 0 And InStr(SelfStudyTable(i, Source(KerRowJ, COAZhou) + 2), "D") = 0 Then
                    If Source(KerRowJ, COABanCi) = 2 Then
                        Source(KerRowJ, COAZiXi) = Source(KerRowJ, COAZiXi) & "��"
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
            For i = 1 To SSRmax      '��ϰ���в���Dʱ�������Ż�
                If InStr(SelfStudyTable(i, 1), Source(KerRowJ, COAXingMing)) > 0 And InStr(SelfStudyTable(i, Source(KerRowJ, COAZhou) + 2), "D") = 0 Then
                    If Source(KerRowJ, COABanCi) = 2 Then
                        Source(KerRowJ, COAZiXi) = Source(KerRowJ, COAZiXi) & "��"
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
            For i = 1 To SSRmax      '��ϰ���в���Dʱ�������Ż�
                If InStr(SelfStudyTable(i, 1), Source(KerRowJ, COAXingMing)) > 0 And InStr(SelfStudyTable(i, Source(KerRowJ, COAZhou) + 2), "E") = 0 Then
                    Source(KerRowJ, COAZiXi) = Source(KerRowJ, COAZiXi) & "��"
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
' ������ϰ����Ϊ������Ա�п��ܻ��Σ�����漰����ϰ�ĵĿ���
    Dim i, j As Integer
    Dim SSitem As String
' ����
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
' ����
    For i = 2 To CGRmax
        For j = 2 To SSRmax
'
            If DateMin <= CDate(Change(i, 3)) And CDate(Change(i, 3)) <= DateMax Then
                If InStr(SelfStudyTable(j, 1), Change(i, 7)) > 0 Then
                    For k = 3 To 7
                        If InStr(SelfStudyTable(1, k), Change(i, 5)) > 0 Then
                            If InStr(SelfStudyTable(j, k), Change(i, 6)) > 0 Then
                                MsgBox SelfStudyTable(j, 1) & "�Ѿ�����" & Change(i, 6) & "������Ч��"
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
                                MsgBox SelfStudyTable(j, 1) & "�Ѿ�����" & Change(i, 12) & "������Ч��"
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
                If InStr(SelfStudyTable(i, Source(j, COAZhou) + 2), "B") > 0 Then  '����B��
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
                        Source(j, COAZiXi) = "��"
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
                If InStr(SelfStudyTable(i, Source(j, COAZhou) + 2), "C") > 0 Then '����C��
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
                        Source(j, COAZiXi) = "��"
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
                If InStr(SelfStudyTable(i, Source(j, COAZhou) + 2), "D") > 0 Then '����D��
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
                If InStr(SelfStudyTable(i, Source(j, COAZhou) + 2), "E") > 0 Then '����E��
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
                S2RowMax = S2RowMax + 1                                                                             '���ְ��2�������Σ�
                For k = 1 To UBound(Source2, 2)
                    Source2(S2RowMax, k) = Source(j, k)
                Next
        End Select
    Next
' ֻ����쳣����
    Call GenerateBook(Source1, S1RowMax, 12, NameTeacherUN)
    Call GenerateBook(Source2, S2RowMax, 15, NameHeadMasterUN)
' �رռ��˳�
    Application.DisplayAlerts = False
    Workbooks.Close                                                                                                 '�ر����й�����
    Application.DisplayAlerts = True
    Application.Quit                                                                                                '�˳�Excel
    Shell "explorer.exe " & OutFolder, vbNormalFocus
End Sub
Sub COAEXE1(KernelRow, KernelCol)
    Dim KernelBanCi, KernelResultCol As Integer
    KernelBanCi = Source(KernelRow, COABanCi) + 1
    KernelResultCol = 3 * (Source(KernelRow, COABanCi) - 1)                                                                                   '��������
    If Len(Source(KernelRow, COAQianDaoSe)) = 0 Then
       If Len(Source(KernelRow, COAQianDao)) = 0 Then
           Source(KernelRow, COAQianDaoSe) = 4
           Source(KernelRow, COAQianDao) = "©ǩ"
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
           Source(KernelRow, COAQianTui) = "©ǩ"
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
    If Source(KernelRow, COAZhou) = BZBeginNum And Source(KernelRow, COABanCi) = 1 Then                                            'ֻ������ִ�����ã���һ��ֻ����һ��
        BZSNum = 2: BZXNum = 2: BZWNum = 2
    ElseIf Source(KernelRow, COAZhou) = BZBeginNumX And Source(KernelRow, COABanCi) = 1 Then
        BZSNumX = 1: BZXNumX = 1: BZWNumX = 1
    End If
    Dim KernelBanCi, KernelResultCol As Integer
    KernelBanCi = Source(KernelRow, COABanCi) + 1
    KernelResultCol = 3 * (Source(KernelRow, COABanCi) - 1)                                                                         '��������
    If Len(Source(KernelRow, COAQianDaoSe)) = 0 Then
       If Len(Source(KernelRow, COAQianDao)) = 0 And Source(KernelRow, COABanCi) <> 3 Then                                          'ͳ�����������©ǩ
            Source(KernelRow, COAQianDaoSe) = 4
            Source(KernelRow, COAQianDao) = "©ǩ"
            If Source(KernelRow, COAZhou) = 6 Or Source(KernelRow, COAZhou) = 7 Then
                If Source(KernelRow, COABanCi) = 1 Then                                                                             '����ǩ��©ǩ
                    If BZSNumX > 0 Then
                        BZSNumX = BZSNumX - 1
                        Source(KernelRow, COAQianDaoSe) = 6
                    Else
                        Source(KernelRow, COAShangLou + KernelResultCol) = 1
                    End If
                ElseIf Source(KernelRow, COABanCi) = 2 Then                                                                          '����ǩ��©ǩ
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
       ElseIf Source(KernelRow, COABanCi) <> 3 Then                                                                               'ͳ�����������ٵ�
           If CDate(Source(KernelRow, COAQianDao)) <= CDate(Standard(KernelBanCi, 3 + KernelCol)) Then
               Source(KernelRow, COAQianDaoSe) = 1
           ElseIf CDate(Source(KernelRow, COAQianDao)) <= CDate(Standard(KernelBanCi, 4 + KernelCol)) Then
               Source(KernelRow, COAQianDaoSe) = 2
           Else
               Source(KernelRow, COAQianDaoSe) = 3
               If Source(KernelRow, COABanCi) = 1 Then                                                                            '����
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
               ElseIf Source(KernelRow, COABanCi) = 2 Then                                                                      '����
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
                Source(KernelRow, COAQianTui) = "©ǩ"
           If Source(KernelRow, COAZhou) = 6 Or Source(KernelRow, COAZhou) = 7 Then
                If Source(KernelRow, COABanCi) = 1 Then
                    Source(KernelRow, COAQianTuiSe) = 6                                                                         '��ͳ����ĩ��ǩ�ˣ�����ɫ�������
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
    Application.SheetsInNewWorkbook = 1                                                                     '����1��Sheet
    Set OutBook = Workbooks.Add
    Application.DisplayAlerts = False
    OutBook.SaveAs Filename:=OutFolder & "\" & InName & OutFileFix & ".xlsx"
    Sheets(1).Name = InName & OutFileFix
    Sheets(InName & OutFileFix).Columns("E:F").NumberFormatLocal = TimeFormat                               '����ʱ���ʽ
'����ͳ�ƽ�������˺ϸ���Ա
    For i = 1 To InCmax
        OutSource(1, i) = Source(1, i)
    Next
    OutSource(1, InCmax + 1) = "����"
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
                            OutZhou = "һ"
                        Case Is = 2
                            OutZhou = "��"
                        Case Is = 3
                            OutZhou = "��"
                        Case Is = 4
                            OutZhou = "��"
                        Case Is = 5
                            OutZhou = "��"
                        Case Is = 6
                            OutZhou = "��"
                        Case Is = 7
                            OutZhou = "��"
                    End Select
                    OutSource(p, COARiQi) = Format(OutSource(p, COARiQi), DateFormat) & "(" & OutZhou & ")"
                    OutSource(p, COAZhou) = OutSource(p, COAZiXi)                                           '�����滻Ϊ��ϰ��Ϣ
                    Select Case OutSource(p, COABanCi)                                                      '�ָ��Ű��Ϊ������
                        Case Is = 1
                            OutSource(p, COABanCi) = "����"
                        Case Is = 2
                            OutSource(p, COABanCi) = "����"
                        Case Is = 3
                            OutSource(p, COABanCi) = "����"
                    End Select
                Next
                For q = COAShangChi To InCmax
                    OutSource(p - 1, q) = Source(1, q)                                                      'ͳ�ƽ����һ�м������
                Next
                OutSource(p - 1, InCmax + 1) = "����"
            End If
            k = 0
        End If
    Next
    OutMax = p
'���ͳ�ƽ��������Ա
    Range(Cells(1, 1), Cells(OutMax, InCmax + 1)) = OutSource                                               '������д���½������
    Call COAFormat(Range(Cells(1, 1), Cells(OutMax, InCmax + 1)))                                           '�����ʽ��
    Range(Cells(1, 1), Cells(1, InCmax + 1)).Font.Bold = True                                               '����Ӻ�
'����ɫ����ɫ
    For i = 2 To OutMax
        Select Case CInt(OutSource(i, COAQianDaoSe))
            Case Is = 1
                Call COAColor(Cells(i, COAQianDao), 4, 10)                                                           '���̵�+������
            Case Is = 2
                Call COAColor(Cells(i, COAQianDao), 6, 10)                                                           '��ɫ��+������
            Case Is = 3
                Call COAColor(Cells(i, COAQianDao), 3, 2)                                                            '����+��ɫ��
                Call COAColor(Cells(i, COABanCi), 3, 2)
            Case Is = 4
                Call COAColor(Cells(i, COAQianDao), 3, 2)                                                            '����+��ɫ��
                Call COAColor(Cells(i, COABanCi), 3, 2)
            Case Is = 5
                Call COAColor(Cells(i, COAQianDao), 10, 6)                                                           '���̵�+��ɫ��
            Case Is = 6
                Call COAColor(Cells(i, COAQianDao), 5, 2)                                                            '����ɫ+��ɫ�֣����Կ���ȥ����)
                Call COAColor(Cells(i, COABanCi), 5, 2)
        End Select
        Select Case CInt(OutSource(i, COAQianTuiSe))
            Case Is = 1
                Call COAColor(Cells(i, COAQianTui), 4, 10)                                                           '���̵�+������
            Case Is = 2
                Call COAColor(Cells(i, COAQianTui), 6, 10)                                                           '��ɫ��+������
            Case Is = 3
                Call COAColor(Cells(i, COAQianTui), 3, 2)                                                            '����+��ɫ��
                Call COAColor(Cells(i, COABanCi), 3, 2)
            Case Is = 4
                Call COAColor(Cells(i, COAQianTui), 3, 2)                                                            '����+��ɫ��
                Call COAColor(Cells(i, COABanCi), 3, 2)
            Case Is = 5
                Call COAColor(Cells(i, COAQianTui), 10, 6)                                                           '���̵�+��ɫ��
            Case Is = 6
                Call COAColor(Cells(i, COAQianTui), 5, 2)                                                            '����ɫ+��ɫ�֣����Կ���ȥ����)
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
                If OutSource(i, m) > 0 Then                                                                 'ͳ������ɫ
                    Call COAColor(Cells(i, m), 3, 2)                                                        '����+��ɫ��
                ElseIf OutSource(i, m) = 0 Then
                    Call COAColor(Cells(i, m), 4, 10)                                                       '���̵�+������
                End If
            Next
            Range(Cells(i - k, 1), Cells(i, 1)).Merge                                                       '�ϲ�����
            For q = COAShangChi To InCmax Step 3
                Range(Cells(i - k, q), Cells(i - 2, q + 2)).Merge                                           '�ϲ�ͳ����
            Next
            Cells(i - k, COAShangChi) = OutSource(i, COAHuanKe)                                                         '���������Ϣ
            k = 0
        End If
        If OutSource(i, COARiQi) = OutSource(i + 1, COARiQi) Then
            p = p + 1
        Else
            Range(Cells(i - p, COARiQi), Cells(i, COARiQi)).Merge                                                       '�ϲ�����
            p = 0
        End If
    Next
'�������У������ڵ����϶��ղ鿴
    Cells(1, 1).Select                                                                                      'Ψһѡ����Ԫ��ĵط�
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    ActiveWorkbook.Close savechanges:=True
    Application.DisplayAlerts = True
    Set OutBook = Nothing                                                                                   'ȡ��OutBook
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
' У׼����ʱ��Ϊ��׼����һ�����գ���Ϊ�꼶��Ҫ�����ύ�����嵽�����ĵĿ��ڣ���ѧУҪ���ύ��һ�����յĿ��ڱ���              '
'                                                                                                                           '
' ע�⣺�������ʱ��ڵ�����������������������ٱ�ʱ���˻�û���Ͻ�����ļ���������Ҫ������ļ����ֹ����뵽ѧУ�������ܱ�  '
'                                                                                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    GEndDate = DateMax
    Do While Weekday(GEndDate, 2) < 7
        GEndDate = GEndDate + 1
    Loop
    GBeginDate = GEndDate - 6
''������ٱ������Ͻ�����ٱ�leave
    ReDim PreLeave(1 To 6000, 1 To 7) As Variant
    PreLeave(1, 1) = Format(GBeginDate, "m" & "��" & "d" & "��") & "-" & Format(GEndDate, "m" & "��" & "d" & "��") & "�������"
    PreLeave(2, 1) = "�꼶/����"
    PreLeave(2, 2) = "ʱ��"
    PreLeave(2, 3) = "����"
    PreLeave(2, 4) = "����"
    k = 2
    For i = 2 To HRmax
        If IsDate(Holiday(i, 2)) And IsDate(Holiday(i, 3)) Then             '���ж�Ϊ���ں���ִ����ٲ���
            If CDate(GBeginDate) <= CDate(Holiday(i, 3)) And CDate(Holiday(i, 2)) <= CDate(GEndDate) And InStr(Holiday(i, 1), "*") = 0 Then
                k = k + 1
                PreLeave(k, 1) = Nianji
'' ȡ����ٵ�ʱ����
                If CDate(Holiday(i, 2)) <= CDate(GBeginDate) Then           'ȡ����ʼʱ��
                    PreLeave(k, 2) = CDate(GBeginDate)
                Else
                    PreLeave(k, 2) = CDate(Holiday(i, 2))
                End If
                If CDate(GEndDate) <= CDate(Holiday(i, 3)) Then             'ȡ����ֹʱ��
                    PreLeave(k, 3) = CDate(GEndDate)
                Else
                    PreLeave(k, 3) = CDate(Holiday(i, 3))
                End If
'' ����Ԥ������ٱ��Ͻ�ѧУ��
                PreLeave(k, 6) = Holiday(i, 1)
                If Holiday(i, 4) > 0 Then
                    PreLeave(k, 7) = Holiday(i, 4)
                    PreLeave(k, 4) = PreLeave(k, 3) - PreLeave(k, 2) + 1
                Else
                    PreLeave(k, 4) = 0.5 * (PreLeave(k, 3) - PreLeave(k, 2) + 1)
                    If Holiday(i, 5) > 0 Or Holiday(i, 6) > 0 Then
                        PreLeave(k, 5) = "����"
                    End If
                    If Holiday(i, 7) > 0 Or Holiday(i, 8) > 0 Then
                        PreLeave(k, 5) = "����"
                    End If
                    If Holiday(i, 9) > 0 Or Holiday(i, 10) > 0 Then
                        PreLeave(k, 5) = "����"
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
''''ͬһ���˵������Ϣ����������ļ����У�����ͬһ�������Ϣ�ֿ���д�ģ�����Ӧ���ϲ���һ��
    ReDim Leave(1 To UBound(PreLeave, 1), 1 To 4) As Variant
    For i = 1 To 2
        For j = 1 To 4
            Leave(i, j) = PreLeave(i, j)
        Next
    Next
    l = 2
    For i = 3 To UBound(PreLeave, 1) - 1
        If PreLeave(i, 6) > 0 And InStr(PreLeave(i, 6), "��") = 0 Then
            k = i + 1
            Do While PreLeave(i, 6) = PreLeave(k, 6)
                k = k + 1
            Loop
            If k > i + 1 Then
             l = l + 1
             Leave(l, 1) = PreLeave(i, 1)
             Leave(l, 2) = Format(PreLeave(i, 2), "m" & "��" & "d" & "��") & PreLeave(i, 5) & "-" & Format(PreLeave(k - 1, 3), "m" & "��" & "d" & "��") & PreLeave(k - 1, 5)
             Leave(l, 3) = PreLeave(l, 6)
             For n = i To k - 1
                Leave(l, 4) = Leave(l, 4) + PreLeave(n, 4)
             Next
             Leave(l, 4) = PreLeave(k - 1, 7) & "��" & Leave(l, 4) & "��"
             i = k - 1
            Else
             l = l + 1
             Leave(l, 1) = PreLeave(i, 1)
             If CDate(PreLeave(i, 2)) = CDate(PreLeave(i, 3)) Then
                Leave(l, 2) = Format(PreLeave(i, 2), "m" & "��" & "d" & "��") & PreLeave(i, 5)
             Else
                Leave(l, 2) = Format(PreLeave(i, 2), "m" & "��" & "d" & "��") & "-" & Format(PreLeave(i, 3), "m" & "��" & "d" & "��")
             End If
             Leave(l, 3) = PreLeave(i, 6)
             Leave(l, 4) = PreLeave(i, 7) & "��" & PreLeave(i, 4) & "��"
            End If
       End If
  Next
  NameLeave = Nianji & Format(GBeginDate, "m" & "��" & "d" & "��") & "-" & Format(GEndDate, "m" & "��" & "d" & "��") & "����"
  Call OutToLeave(Leave, l, UBound(Leave, 2), OutFolder, NameLeave)
End Sub

Sub OutToLeave(LeaveSource, InRmax, InCmax, OutLeaveFolder, InName)
    Dim OutBook As Workbook
    Application.SheetsInNewWorkbook = 1                                                         '����1��Sheet
    Set OutBook = Workbooks.Add
    ActiveWindow.FreezePanes = False                                                            '��ֹ���ᴰ��
    Application.DisplayAlerts = False
        OutBook.SaveAs Filename:=OutLeaveFolder & "\" & InName & ".xlsx"
        Sheets(1).Name = InName
        Range(Cells(1, 1), Cells(InRmax, InCmax)) = LeaveSource                                 '���������Ϣ
        Call COAFormat(Range(Cells(1, 1), Cells(InRmax, InCmax)))
        Range(Cells(1, 1), Cells(1, InCmax)).Merge
        Range(Cells(1, 1), Cells(1, InCmax)).Font.Size = 20
        Range(Cells(2, 1), Cells(2, InCmax)).Font.Size = 14
    ActiveWorkbook.Close savechanges:=True
    Application.DisplayAlerts = True
    Set OutBook = Nothing                                                                       'ȡ��OutBook
End Sub




