Attribute VB_Name = "V68���ڴ���ģ��"
' ͨ�ÿ���ϵͳV68
' ���ߣ�����
' ���ڣ�2021��4��25��--2021��7��2��
' ���ã������׼��������Դ����Original
' ˵������У׼���ɵ������������ļ�ֱ������ΪCorrect ,�������ĺô����ڿ��Ը�����ϰ�仯��ʱ����У׼����������ɺ���õ�ʱ�伸�����Ժ��Բ��ƣ��ͷ�������������������������
' ע�⣺�������ʱ��ڵ�����������������������ٱ�ʱ���˻�û���Ͻ�����ļ���������Ҫ������ļ����ֹ����뵽ѧУ�������ܱ�
' ��־���淶����ȫ�ֱ�����ֵ���������׿��Ƹ�������2021/4/27
' ��־�������˹��ڼ��ں���һʱ��ֹͣ���ڵ����ã����������ļ��ϲ���һ���ļ��У����������ɱ�����
' ��־�����ٱ���ԭʼ�ļ���ֱ�ӽ�Ŀ���ļ����ɵ�Ŀ���ļ��У�Ȼ��ر�Դ�ļ���������Ŀ���ļ��� 2021/4/28
' ��־���Ż��˻����������ã�ʹ������Զ�����ʱ�����ɶ�Ӧ���ļ���
' ��־�����ļ������ɲ���ת��ͨ�ô���ģ�飬�淶���� 2021/4/29
' ��־��У׼��Ͷ���У׼��������ڣ������ӿ���Ψһ��λ��ĳһ�У���һ�ܵ�������Բ�����һ����Ӱ��2021/4/30
' ��־��ͳ��Υ��ʱ�����뺯��CDate()ǿ��ת����ʱ���ʽ
' ��־��������α�͵��μ�¼2021/5/1
' ��־���Ż�����ģ��2021/5/5
' ��־�����������Զ�ѡ��ʮ��һ������һ�յ�У׼ʱ�䡢��׼ʱ�䡣ϵͳ�����ļ���Ŀ¼����׺�����ļ��е����ʱ�����Сʱ��ȷ���������Ե�ǰ����ʱ��䵱��׺
' ��־��ͳ��ģ����GoTo�Ż��˴���2021/5/13
' ��־�����Ӱ����εĿ��ڵ��Կ��ˣ���һ����������1�κ�������ʦһ������ǩ��������9��50ǰ��������1�Σ���������Ҳ����1�Σ�������һ�������������3�κ�������ʦһ��ǩ����������������1��
' ��־���޸İ����ε��Կ��˹���ȥ��������Ϣ���bug,ȥ��׷�ٵ�6��ʱ��bug2021/5/17
' ��־���޸����б���Ϊ���б����������ڲ�ͬģ����Ϲ���ʱ�ǳ�����  2021/5/23
' ��־����ǿ��ٱ�ı�д��ʽ����ʱ�Զ����v59��ʽ                 2021/5/23
' ��־���޸����α�bug 2021/5/24
' ��־������ѧУ������ٱ���Զ����룬�������ٵ��ܱ��Զ����빦��
' ��־���޸�ѧУ������ٱ�bug 2021/6/2 �����汾��ΪV62
' ��־���޸������ε��Կ���bug 2021/6/3 �����汾��ΪV63
' ��־���޸���������У׼��һ�ܣ�Ŀǰ����Ҫ�������Ͻ����ڣ���ѧУҪ����һ�Ͻ���һ�ܵĽ��������һ���������������������Ѿ����������ⲻ�� 2021/6/3 �����汾��V64
' ��־���޸���ٻ����ܱ�bug��ͬʱ�����ʦ�����α�������ж��ᣬ�ڵ����ϲ鿴�������ֱ��     2021/6/4
' ��־����д�˻��γ��򣬼��˵��α�ǣ����ڱ�ע���ע���ɡ�ͬʱ���θ��Ӻ�������ɶ�����ǿ 2021/6/4 �����汾��V65
' ��־�������ж�ʱ��ʱIsNumberic ΪIsDate ,�������Աܿ�Դ������ʱ��ʱ��ת�������Ӻ�������V65�Ѿ����ã����Ǳ�����������˼����Լ�Ч�ʣ������汾��ΪV66
' ��־���޸������ε��Կ��ˣ����ڲ�ͳ������ǩ��������ǩ�������µ���ɫ����V66
' ��־������ʹ��CDateת����ʱ�䣬��������10��-17�����������޸������ļ���׼�Ա���Ϣ��59��ֽ磬����ϵͳ��õ�00��ʱ��߱��ϸ��������V66
' ��־�����ӳ��������RowMax=10000�������ColMax=200,ɾ��GroupName,�Ż��������ļ�ȡ�÷�ʽ��GetHolidayģ��2021/6/13 �������汾��V67
' ��־���������ļ�����Ϊ����ȫУ��ʦ��ϵ��ʽ�������ļ����������Ϲ�ģ�����ϴ����ǽ�Col����Ϊ1000����Ӧ����������������㹻�ռ䡣ͬʱ�����������ļ�����������˶԰������
'       �����Ϣʱ�жϵ�bug,�˰��޸��˴�bug,׷���˶�TeGrStep �� TeGrZhiWu �����кŵ��Զ��жϡ������汾��ΪV68 2021/7/2
'
'
'' ����ȫ�ֱ���
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
    Public OriginalSheet1 As String                                                                            '����4��רΪУ׼������
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
    Public BeginDate As Date                                                                                    '���ڿ�ʼʱ��
    Public EndDate As Date                                                                                      '���ڽ���ʱ��
    Public WuYi As Date
    Public ShiYi As Date
    Public Morning As Integer                                                                                   '�����ϰ�ʱ��Υ�����
    Public Afternoon As Integer
    Public Evening As Integer
    Public MorningX As Integer                                                                                  '��ĩʱ��Υ�����
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
    Public SRmax, SCmax, HRmax, HCmax, TGRmax, TGCmax, CRmax, CCmax, STRmax, STCmax, ViRmax, ViCmax, ORmax, OCmax As Integer      'ͨ����ֵ
    Public SSRmax, SSCmax, RCTRmax, RCTCmax, CTRmax, CTCmax, CTERmax, CTECmax As Integer                                          'רΪУ׼������
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
    Public PreLeave As Variant                     ' ׼��������������ѧУ�Ͻ�����ٿ��ڱ���δ�ϲ�
    Public Leave As Variant                        ' ��Leaveͬһ���˵���Ϣ����Ϊһ�飬Ȼ�����ɸ�ʽ�������ٿ��ڱ�
    Public LeaveName As Variant
    Public TTsize As Variant
    Public OutLeavePath As String
    Public LeaveBook As Workbook
    Public ChangedRow As Integer
    Public Const RowMax As Integer = 10000
    Public Const ColMax As Integer = 1000
    Public TeGrStep, TeGrZhiWu As Integer
'ȫ�ֱ�����ֵ
Sub ������������()
    ConfigPath = "D:\����ϵͳ"
    ConfigFolder = ConfigPath & "\" & "����ϵͳ����"
    OutPath = ConfigPath & "\" & Format(Now, "yyyy" & "��") & "����"
    DateFormat = "m""��""d""��"";@"
    TimeFormat = "h:mm;@"
    NameFont = "����"
    NameOriginal = "Դ��"
    NameHeadMasterUN = "�쳣������"
    NameTeacherUN = "�쳣��ʦ"
    NameHeadMaster = "������"
    NameTeacher = "��ʦ"
'    ConfigFile = ConfigFolder & "\" & "��������.xlsx"                                                          '������ͨ�������Զ�У׼�������ļ�
    ConfigFile = ConfigFolder & "\" & "��������.xlsm"                                                           '������ǿ���Զ�У׼���ܵ������ļ�
    ConfigSheet1 = "ֹͣ����"
    ConfigSheet2 = "��ٱ�"
    OriginalSheet1 = "��ϰ����"                                                                                 '����4��רΪУ׼������
    OriginalSheet3 = "����У׼"
    OriginalSheet4 = "���α�"
    OriginalSheet5 = "��ʦ����"
    WuYi = Format(Now, "yyyy") & "/5/1"
    ShiYi = Format(Now, "yyyy") & "/10/1"
    If CDate(WuYi) < CDate(Now) < CDate(ShiYi) Then
        ConfigSheet3 = "51ST"
        OriginalSheet2 = "51CT"
    Else
        ConfigSheet3 = "10ST"
        OriginalSheet2 = "10CT"
    End If
    VipSwitch = 1                                                                                               '����vip
    NormalSwitch = 1                                                                                            '0�������������1���
    StopSymbol = "*"
End Sub
'' �����ٱ�
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
    Do Until TeacherGroup(2, TeGrZhiWu + 1) = "ְ��"
        TeGrZhiWu = TeGrZhiWu + 1
    Loop
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' У׼����ʱ��Ϊ��׼����һ�����գ���Ϊ�꼶��Ҫ�����ύ�����嵽�����ĵĿ��ڣ���ѧУҪ���ύ��һ�����յĿ��ڱ���              '
'                                                                                                                           '
' ע�⣺�������ʱ��ڵ�����������������������ٱ�ʱ���˻�û���Ͻ�����ļ���������Ҫ������ļ����ֹ����뵽ѧУ�������ܱ�  '
'                                                                                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    GEndDate = EndDate
    Do While Weekday(GEndDate, 2) < 7
        GEndDate = GEndDate + 1
    Loop
    GBeginDate = GEndDate - 6
   
''������ٱ������Ͻ�����ٱ�leave
    LeaveName = "�߶�����" & Format(GBeginDate, "m" & "��" & "d" & "��") & "-" & Format(GEndDate, "m" & "��" & "d" & "��") & "����"
    ReDim PreLeave(1 To 6000, 1 To 7) As Variant
    PreLeave(1, 1) = Format(GBeginDate, "m" & "��" & "d" & "��") & "-" & Format(GEndDate, "m" & "��" & "d" & "��") & "�������"
    PreLeave(2, 1) = "�꼶/����"
    PreLeave(2, 2) = "ʱ��"
    PreLeave(2, 3) = "����"
    PreLeave(2, 4) = "����"
    k = 2
    For i = 2 To HRmax
        If IsDate(PreHoliday(i, 3)) And IsDate(PreHoliday(i, 4)) Then           '���ж�Ϊ���ں���ִ����ٲ���
            If CDate(GBeginDate) <= CDate(PreHoliday(i, 4)) And CDate(PreHoliday(i, 3)) <= CDate(GEndDate) And InStr(PreHoliday(i, 1), "*") = 0 Then
                k = k + 1
                PreLeave(k, 1) = "�߶�"
'' ȡ����ٵ�ʱ����
                If CDate(PreHoliday(i, 3)) <= CDate(GBeginDate) Then           'ȡ����ʼʱ��
                    PreLeave(k, 2) = CDate(GBeginDate)
                Else
                    PreLeave(k, 2) = CDate(PreHoliday(i, 3))
                End If
                If CDate(GEndDate) <= CDate(PreHoliday(i, 4)) Then             'ȡ����ֹʱ��
                    PreLeave(k, 3) = CDate(GEndDate)
                Else
                    PreLeave(k, 3) = CDate(PreHoliday(i, 4))
                End If
'' ����Ԥ������ٱ��Ͻ�ѧУ��
                PreLeave(k, 6) = PreHoliday(i, 1)
                If PreHoliday(i, 5) > 0 Then
                    PreLeave(k, 7) = PreHoliday(i, 5)
                    PreLeave(k, 4) = PreLeave(k, 3) - PreLeave(k, 2) + 1
                Else
                    PreLeave(k, 4) = 0.5 * (PreLeave(k, 3) - PreLeave(k, 2) + 1)
                    If PreHoliday(i, 6) > 0 Or PreHoliday(i, 7) > 0 Then
                        PreLeave(k, 5) = "����"
                    End If
                    If PreHoliday(i, 8) > 0 Or PreHoliday(i, 9) > 0 Then
                        PreLeave(k, 5) = "����"
                    End If
                    If PreHoliday(i, 10) > 0 Or PreHoliday(i, 11) > 0 Then
                        PreLeave(k, 5) = "����"
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
'' ������ٵ�У׼��Holiday
    ReDim Holiday(1 To 3000, 1 To 6)
    Holiday(1, 1) = "����"
    Holiday(1, 2) = "ʱ��"
    Holiday(1, 3) = "���"
    Holiday(1, 4) = "����"
    Holiday(1, 5) = "ǩ��"
    Holiday(1, 6) = "ǩ��"
    p = 1
    For i = 2 To HRmax
        If PreHoliday(i, 1) > 0 Then
    ''''' ȡ�ý�ʦ����Ͷ�Ӧ���� debug
            If InStr(PreHoliday(i, 1), "��") > 0 Or InStr(PreHoliday(i, 1), "������") > 0 Or InStr(PreHoliday(i, 1), "������") > 0 Then
                For j = 1 To TGCmax Step TeGrStep
                    If TeacherGroup(1, j) = PreHoliday(i, 1) Then
                        GroupColum = j
                    End If
                Next
                GroupRow = ConfigBook.Sheets(OriginalSheet5).Cells(RowMax, GroupColum).End(xlUp).Row
            End If
    '''''
            If IsDate(PreHoliday(i, 3)) Then
            ' ȡ����Чʱ��
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
            ' ת����ͬģʽ
    '''''''
            If InStr(PreHoliday(i, 1), "��") > 0 Or InStr(PreHoliday(i, 1), "������") > 0 Or InStr(PreHoliday(i, 1), "������") > 0 Then
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
' ȡ�����ڵ���ʼ
                For j = 1 To 7
                 If PreHoliday(i, 3) = Choose(j, "һ", "��", "��", "��", "��", "��", "��") Then
                    WeekX = j
                 End If
                Next
                For j = 1 To 7
                 If PreHoliday(i, 4) = Choose(j, "һ", "��", "��", "��", "��", "��", "��") Then
                    WeekY = j
                 End If
                Next
    ' �������뵽�����
''''
                If InStr(PreHoliday(i, 1), "��") > 0 Or InStr(PreHoliday(i, 1), "������") > 0 Or InStr(PreHoliday(i, 1), "������") > 0 Then
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
                                            Holiday(p, 4) = "һ"
                                        Case Is = 2
                                            Holiday(p, 4) = "��"
                                        Case Is = 3
                                            Holiday(p, 4) = "��"
                                        Case Is = 4
                                            Holiday(p, 4) = "��"
                                        Case Is = 5
                                            Holiday(p, 4) = "��"
                                        Case Is = 6
                                            Holiday(p, 4) = "��"
                                        Case Is = 7
                                            Holiday(p, 4) = "��"
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
                                       Holiday(p, 4) = "һ"
                                   Case Is = 2
                                       Holiday(p, 4) = "��"
                                   Case Is = 3
                                       Holiday(p, 4) = "��"
                                   Case Is = 4
                                       Holiday(p, 4) = "��"
                                   Case Is = 5
                                       Holiday(p, 4) = "��"
                                   Case Is = 6
                                       Holiday(p, 4) = "��"
                                   Case Is = 7
                                       Holiday(p, 4) = "��"
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
    HRmax = p                                                                                                   '���Holiday�ķǿ�����������
    HCmax = UBound(Holiday, 2)
End Sub
Sub ��׼ͨ�ÿ���(Original)
    Dim i, j, k, l, m, n, o, p As Integer
    Dim SelfStudyTemp As Variant
    Dim SFO As Object
'' ������ü�����ļ����Ƿ���ڣ��������ڣ����½�
    Set SFO = CreateObject("Scripting.FileSystemObject")                                                        '��SFOΪ�ļ��ж������
    If SFO.FolderExists(ConfigPath) = False Then
       MkDir ConfigPath
    End If
    If SFO.FolderExists(ConfigFolder) = False Then
       MkDir ConfigFolder
    End If
    If SFO.FolderExists(OutPath) = False Then
       MkDir OutPath
    End If
'' ��ȡ�����ļ�
    Set ConfigBook = GetObject(ConfigFile)
''
    ViRmax = ConfigBook.Sheets(ConfigSheet1).Cells(RowMax, 1).End(xlUp).Row
    ViCmax = ConfigBook.Sheets(ConfigSheet1).Cells(1, ColMax).End(xlToLeft).Column
    VipSource = ConfigBook.Sheets(ConfigSheet1).Range(ConfigBook.Sheets(ConfigSheet1).Cells(1, 1), ConfigBook.Sheets(ConfigSheet1).Cells(ViRmax, ViCmax))

'' ���ͳ��ʱ�䷶Χ
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
'' �����ٱ�
    Call GetHoliday
''
    STRmax = ConfigBook.Sheets(ConfigSheet3).Cells(RowMax, 1).End(xlUp).Row
    STCmax = ConfigBook.Sheets(ConfigSheet3).Cells(1, ColMax).End(xlToLeft).Column
    Standard = ConfigBook.Sheets(ConfigSheet3).Range(ConfigBook.Sheets(ConfigSheet3).Cells(1, 1), ConfigBook.Sheets(ConfigSheet3).Cells(STRmax, STCmax))
''                                                                                                          '����4��רΪУ׼������
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
    CGRmax = ConfigBook.Sheets(OriginalSheet4).Cells(RowMax, 1).End(xlUp).Row                               '���뻻�α�
    CGCmax = ConfigBook.Sheets(OriginalSheet4).Cells(1, ColMax).End(xlToLeft).Column
    Change = ConfigBook.Sheets(OriginalSheet4).Range(ConfigBook.Sheets(OriginalSheet4).Cells(1, 1), ConfigBook.Sheets(OriginalSheet4).Cells(CGRmax, CGCmax))
''
    ReDim CorrectTable(1 To 6000, 1 To 9) As Variant
    CTRmax = UBound(CorrectTable, 1)
    CTCmax = UBound(CorrectTable, 2)
''''''''''''''''''
' У׼�������ϰ�����ɣ����Ե��λ�Ӱ�쵽��ϰ�����������У׼��ϰ��
    Dim SSTB As Variant
    Dim SSTC As Variant
    Dim SSTD As Variant
    Dim SSTE As Variant
    SSTB = "B"
    SSTC = "C"
    SSTD = "D"
    SSTE = "E"
    For i = 2 To CGRmax                                                                                 '���εĶ��˶�����ԭ������ϰ�����û����ϰ����ִ�н����Ч
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
                              Select Case Change(i, l + 4)                                              'A��������ϰ����֮�У�����ʹ��A�ų�Ҫȥ������ϰ
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
'���η��������ͬһ�����Լ���һ����ϰ����һ����ϰ�Ի�����㲻�ý��е��Σ����Բ���׷�ӵ��μ�¼ '
'          ���ͬһ�����Լ���һ����ϰ����һ�ڹ�����ϰ�Ի����򲻱ص�������ֻ�������Ӧ��ϰ�� '
'          �ɡ����ԣ�������ĵ���Է���ϰ�У�������ͬһ���˵Ļ��������                       '
'                                                                                             '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    For i = 2 To CGRmax                                                                                         ' ����Է��Ŀ�,���������ϰ��������ϰ�ظ�������ʾ��ͻ��������ʧ��
        If InStr(Change(i, 1), Change(i, 6)) = 0 Then                                                           ' ���ε�˫����������ͬһ����
            If InStr(Change(i, 10), "B") > 0 Or InStr(Change(i, 10), "C") > 0 Or InStr(Change(i, 10), "D") > 0 Or InStr(Change(i, 10), "E") > 0 Then
                If CDate(BeginDate) <= CDate(Change(i, 7)) And CDate(Change(i, 7)) <= CDate(EndDate) Then
                    For j = 2 To SSRmax
                        If Change(i, 1) = SelfStudyTable(j, 1) Then
                            k = Weekday(Change(i, 7), 2) + 2
                            If InStr(SelfStudyTable(j, k), Change(i, 10)) > 0 Then
                                MsgBox Change(i, 1) & "��" & Change(i, 6) & "�������ͻʧ��"
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
                                MsgBox Change(i, 6) & "��" & Change(i, 1) & "�������ͻʧ��"
                            Else
                             SelfStudyTable(j, k) = SelfStudyTable(j, k) & Change(i, 5)
                            End If
                        End If
                    Next
                End If
            End If
        End If
    Next
' ���������Ҫ�����Ľ���м�¼��������Changed ����¼����ǰ���ɣ���������д�������
o = 2 * CGRmax
ReDim Changed(1 To o, 1 To 10) As Variant
k = 0
For i = 2 To CGRmax
    If CDate(BeginDate) <= CDate(Change(i, 2)) Or CDate(BeginDate) <= CDate(Change(i, 7)) Then                  ' ֻ�л��α���ʱ�䳬��ͳ��ʱ��ʱ��д��Changed
        k = k + 1
        For j = 1 To 4
          Changed(k, j) = Change(i, j)
        Next
        Changed(k, 5) = Format(Change(i, 2), DateFormat)
        If InStr(Change(i, 5), "B") + InStr(Change(i, 5), "C") + InStr(Change(i, 5), "D") + InStr(Change(i, 5), "E") > 0 Then
              If InStr(Change(i, 5), "B") > 0 Then
                Changed(k, 5) = Changed(k, 5) & " ��1��"
              End If
              If InStr(Change(i, 5), "C") > 0 Then
                 Changed(k, 5) = Changed(k, 5) & " ��5��"
              End If
              If InStr(Change(i, 5), "E") > 0 Then
                 Changed(k, 5) = Changed(k, 5) & " ��6��"
              End If
              If InStr(Change(i, 5), "D") > 0 Then
                Changed(k, 5) = Changed(k, 5) & " ��9��"
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
                Changed(k, 10) = Changed(k, 10) & " ��1��"
              End If
              If InStr(Change(i, 10), "C") > 0 Then
                 Changed(k, 10) = Changed(k, 10) & " ��5��"
              End If
              If InStr(Change(i, 10), "E") > 0 Then
                 Changed(k, 10) = Changed(k, 10) & " ��6��"
              End If
              If InStr(Change(i, 10), "D") > 0 Then
                Changed(k, 10) = Changed(k, 10) & " ��9��"
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
'���μ�¼����������Changed��Ϊ���ע��������׼��  '
'                                                   '
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'' ����ϰ��ת��ΪУ׼������ǰ���Ѿ�������ϰУ׼������֮�󲻱���׷��У׼������¼���������Ч��
    k = 0
    For i = 2 To SSRmax
        For j = 3 To 7
            If InStr(SelfStudyTable(i, 2), NameTeacher) > 0 Then
''��Խ�ʦ������,����1��B�͵�5��C��Ӱ������
              If InStr(SelfStudyTable(i, j), "B") + InStr(SelfStudyTable(i, j), "C") > 0 Then
                  k = k + 1
                  CorrectTable(k, 1) = SelfStudyTable(i, 1)
                  CorrectTable(k, 3) = SelfStudyTable(i, 2) & "����"
                  CorrectTable(k, 4) = SelfStudyTable(1, j)
                  If InStr(SelfStudyTable(i, j), "B") > 0 Then
                        CorrectTable(k, 5) = CorrectTime(2, 2)
                        CorrectTable(k, 6) = CorrectTime(2, 3)
                        CorrectTable(k, 9) = "��1��"
                  End If
                  If InStr(SelfStudyTable(i, j), "C") > 0 Then
                    CorrectTable(k, 7) = CorrectTime(3, 4)
                    CorrectTable(k, 8) = CorrectTime(3, 5)
                    If CorrectTable(k, 9) = 0 Then
                        CorrectTable(k, 9) = "��5��"
                    Else
                        CorrectTable(k, 9) = "��1,5��"
                    End If
                  End If
''�����е�1�ڻ��5�ڵ�ʱ������������������չ�
                If InStr(SelfStudyTable(i, j), "E") > 0 Then
                     k = k + 1
                     CorrectTable(k, 1) = SelfStudyTable(i, 1)
                     CorrectTable(k, 3) = SelfStudyTable(i, 2) & "����"
                     CorrectTable(k, 4) = SelfStudyTable(1, j)
                     CorrectTable(k, 5) = CorrectTime(6, 2)
                     If InStr(SelfStudyTable(i, j), "D") > 0 Then
                       CorrectTable(k, 8) = CorrectTime(5, 5)
                       CorrectTable(k, 9) = "��6,9��"
                     Else
                       CorrectTable(k, 7) = CorrectTime(4, 4)
                       CorrectTable(k, 9) = "��6��"
                     End If
                Else
                     k = k + 1
                     CorrectTable(k, 1) = SelfStudyTable(i, 1)
                     CorrectTable(k, 3) = SelfStudyTable(i, 2) & "����"
                     CorrectTable(k, 4) = SelfStudyTable(1, j)
                     CorrectTable(k, 6) = CorrectTime(4, 3)
                     CorrectTable(k, 9) = "��"
                     If InStr(SelfStudyTable(i, j), "D") > 0 Then
                       CorrectTable(k, 8) = CorrectTime(5, 5)
                       CorrectTable(k, 9) = CorrectTable(k, 9) & "��9��"
                     End If
                 End If
               ElseIf InStr(SelfStudyTable(i, j), "D") > 0 Then
                  k = k + 1
                  CorrectTable(k, 1) = SelfStudyTable(i, 1)
                  CorrectTable(k, 3) = SelfStudyTable(i, 2) & "����"
                  CorrectTable(k, 4) = SelfStudyTable(1, j)
                  CorrectTable(k, 8) = CorrectTime(5, 5)
                  If InStr(SelfStudyTable(i, j), "E") > 0 Then
                    CorrectTable(k, 9) = "��6,9��"
                  Else
                    CorrectTable(k, 9) = "��9��"
                  End If
               ElseIf InStr(SelfStudyTable(i, j), "E") > 0 Then
                  k = k + 1
                  CorrectTable(k, 1) = SelfStudyTable(i, 1)
                  CorrectTable(k, 3) = SelfStudyTable(i, 2) & "����"
                  CorrectTable(k, 4) = SelfStudyTable(1, j)
                  CorrectTable(k, 5) = CorrectTime(6, 2)
                  CorrectTable(k, 7) = CorrectTime(4, 4)
                  CorrectTable(k, 9) = "��6��"
               End If
 ''��԰����ε�����
            Else
                If InStr(SelfStudyTable(i, j), "B") + InStr(SelfStudyTable(i, j), "C") > 0 Then
                  k = k + 1
                  CorrectTable(k, 1) = SelfStudyTable(i, 1)
                  CorrectTable(k, 3) = SelfStudyTable(i, 2) & "����"
                  CorrectTable(k, 4) = SelfStudyTable(1, j)
                  If InStr(SelfStudyTable(i, j), "B") > 0 Then
                        CorrectTable(k, 9) = "��1��"
                  End If
                  If InStr(SelfStudyTable(i, j), "C") > 0 Then
                    CorrectTable(k, 7) = CorrectTime(3, 4)
                    CorrectTable(k, 8) = CorrectTime(3, 5)
                    If CorrectTable(k, 9) = 0 Then
                        CorrectTable(k, 9) = "��5��"
                    Else
                        CorrectTable(k, 9) = "��1,5��"
                    End If
                  End If
                ElseIf InStr(SelfStudyTable(i, j), "D") > 0 Then
                  k = k + 1
                  CorrectTable(k, 1) = SelfStudyTable(i, 1)
                  CorrectTable(k, 3) = SelfStudyTable(i, 2) & "����"
                  CorrectTable(k, 4) = SelfStudyTable(1, j)
                  CorrectTable(k, 9) = "��9��"
                End If
            End If
        Next
    Next
'����У׼,n��¼ƥ����
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
                            If InStr(CorrectTable(i, 9), "��") > 0 Then
                                  If InStr(CorrectTable(i, 9), "��9��") > 0 Then
                                      CorrectTable(i, 9) = ReCorrectTable(j, 9) & "��9��"
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
''ȡ��CorrectTable �ķǿյ�Ԫ����
    p = 0
    For i = 1 To CTRmax
       If CorrectTable(i, 1) > 0 Then
        p = p + 1
       End If
    Next
    l = p
'' ���Correct(У׼��)����һ����ֻȡ�ǿռ�¼�����Լ�Сͳ��ʱ�ĶԱ�����
    ReDim Correct(1 To p, 1 To 9) As Variant
    For i = 1 To p
        For j = 1 To 9
            Correct(i, j) = CorrectTable(i, j)
        Next
    Next
    CRmax = UBound(Correct, 1)
    CCmax = UBound(Correct, 2)
''����vip����ģ��
If VipSwitch = 0 Then
    Source = Original
Else
    ORmax = UBound(Original, 1)
    OCmax = UBound(Original, 2)
    ReDim Source(1 To ORmax, 1 To OCmax) As Variant
'  ����VIP��Ա
    k = 0
    For i = 2 To ORmax
        l = 0
        If Original(i, 3) <> 0 Then
' ����i�к�VipSource �б�ȶԣ���l��¼�ȶ���������ȶ���Ϊ1���������ƥ�������
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
'��l=0����˵�����в���VipSource �У����뵽Source
            If l = 0 Then
                k = k + 1
                For m = 1 To 6
                    Source(k, m) = Original(i, m)
                Next
            End If
        End If
     Next
End If
' Source ��ͷ����
    Source(1, 1) = "����"
    Source(1, 2) = "����"
    Source(1, 3) = "���"
    Source(1, 4) = "��ϰ"
    Source(1, 5) = "ǩ��"
    Source(1, 6) = "ǩ��"
    Source(1, 7) = "�ϳ�"
    Source(1, 8) = "����"
    Source(1, 9) = "��©"
    Source(1, 10) = "�³�"
    Source(1, 11) = "����"
    Source(1, 12) = "��©"
    Source(1, 13) = "���"
    Source(1, 14) = "����"
    Source(1, 15) = "��©"
' Source ��������������
    SRmax = UBound(Original, 1)
    SCmax = UBound(Original, 2)
' �������Ϣ���뵽Source
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
'У׼Source�е�ǩ��ǩ������
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
'��У׼���Source����ͳ������
    For i = 2 To SRmax
     If InStr(Source(i, 3), Standard(2, 1)) > 0 Then                        '��ʦ����
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
     ElseIf InStr(Source(i, 3), Standard(3, 1)) > 0 Then                    '��ʦ����
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
     ElseIf InStr(Source(i, 3), Standard(4, 1)) > 0 Then                    '����������
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
     ElseIf InStr(Source(i, 3), Standard(5, 1)) > 0 Then                   '����������
        If Source(i, 5) = 0 Then
           Source(i, 12) = Source(i, 12) + 1
        ElseIf IsDate(Source(i, 5)) Then
            If CDate(Source(i, 5)) >= CDate(Standard(5, 3)) Then
               Source(i, 10) = 1
            End If
        End If
' �����˰���������ǩ�ˣ�����ɾ��
'        If Source(i, 6) = 0 Then
'              Source(i, 12) = Source(i, 12) + 1
'        ElseIf Isdate(Source(i, 6)) Then
'           If CDate(Source(i, 6)) < CDate(Standard(5, 4)) Then
'              Source(i, 11) = 1
'           End If
'        End If
     ElseIf InStr(Source(i, 3), Standard(6, 1)) > 0 Then                   '����������
'�����˰���������ǩ��������ɾ��
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
' ���ɽ�ʦ����
    ReDim Teacher(1 To SRmax, 1 To 13)
    k = 1
    For j = 1 To 12
        Teacher(1, j) = Source(1, j)
    Next
        Teacher(1, 13) = "����"
    For i = 1 To SRmax
        If InStr(Source(i, 3), NameTeacher) > 0 Then
            k = k + 1
            For j = 1 To 12
                Teacher(k, j) = Source(i, j)
            Next
        End If
    Next
' ���ɰ���������
    ReDim HeadMaster(1 To SRmax, 1 To 16)
    k = 1
    For j = 1 To 15
        HeadMaster(1, j) = Source(1, j)
    Next
        HeadMaster(1, 16) = "����"
    For i = 1 To SRmax
        If InStr(Source(i, 3), NameHeadMaster) > 0 Then
            k = k + 1
            For j = 1 To 15
                HeadMaster(k, j) = Source(i, j)
            Next
        End If
    Next
'  ��������ļ���
    OutFileFix = "��" & Format(BeginDate, "yyyy" & "��" & "m" & "��" & "d" & "��") & "-" & Format(EndDate, "yyyy" & "��" & "m" & "��" & "d" & "��") & "��"
    OutFolder = OutPath & "\" & Format(EndDate, "m" & "��" & "d" & "��") & "��ʽ�ϱ�"
    If SFO.FolderExists(OutFolder) = False Then
       MkDir OutFolder
    Else
        i = 1                                                         ' �����ʽ�ϱ����ɣ����Ա�ŵķ�ʽ�����ļ���
        Do While SFO.FolderExists(OutFolder & i) = True
            i = i + 1
        Loop
        OutFolder = OutFolder & i
        MkDir OutFolder
    End If
' ����Teacher
    Call ����ͳ�ƴ���(Teacher)
' ����HeadMaster
    Call ����ͳ�ƴ���(HeadMaster)
' ������ٿ����ܱ�
    Call OutToLeave(Leave, UBound(Leave, 1), UBound(Leave, 2), OutFolder, LeaveName)
' ��ٿ��������������ٱ�
    Call AddToTotalLeave
    Application.DisplayAlerts = False
    Workbooks.Close                                                 '�ر����й�����
    Application.DisplayAlerts = True
    Application.Quit                                                '�˳�Excel
    Shell "explorer.exe " & OutFolder, vbNormalFocus                '��Ŀ���ļ���
End Sub
Sub ����ͳ�ƴ���(THDATA)
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
            If SubSource(l, o) > 0 Then                                             'д���쳣�ļ�¼
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
    
' �����쳣�����ο��ڵ��Թ��� ,Morning �ȼ�¼�����ε���ʦ�����ϰ�Ĵ������ǿ���ѡ����ȥ���Ĵ���2021/5/13

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'��ʼ���˵��Թ����п������꼶�ı䣬������˫�߱���˲��֣���Ҫʱ�޸Ĵ˲��ִ��� '
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
                If InStr(Abnormal(l, 4), "��") + InStr(Abnormal(l, 4), "��") > 0 Then                   '�������������յ����
                    If InStr(Abnormal(l, 3), "����") > 0 Then
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
                    If InStr(Abnormal(l, 3), "����") > 0 Then
                        If CDate(Abnormal(l, 5)) = 0 Then
                            AfternoonX = AfternoonX + 1
                        ElseIf IsDate(Abnormal(l, 5)) Then
                            If CDate(Standard(5, 3)) <= CDate(Abnormal(l, 5)) Then
                                AfternoonX = AfternoonX + 1
                            End If
                        End If
                    End If
                    If InStr(Abnormal(l, 3), "����") > 0 Then                                           '������ǩ��Ϊ��ʱ��ȥ�������ǩ�˺�����ǩ��
                        If Abnormal(l, 6) = 0 Then
 '                           If IsDate(Abnormal(l - 1, 6))  Then                                        '������δ���Ҫ�����ϱ���һ��ǩ������ȡ��ע��
 '                            EveningX = EveningX + 1
 '                           ElseIf IsDate(Abnormal(l, 5))  Then
 '                            EveningX = EveningX + 1
 '                           End If
                            EveningX = EveningX + 1                                                     '������δ���Ҫ�����ϱ���һ��ǩ������ע�͵�����
                        ElseIf IsDate(Abnormal(l, 6)) Then
                            If CDate(Abnormal(l, 6)) < CDate(Standard(6, 4)) Then
                                EveningX = EveningX + 1
                            End If
                        End If
                     End If
                Else                                                                                   '������һ����������
                    If InStr(Abnormal(l, 3), "����") > 0 Then                                          '��������ǩ��
                       If IsDate(Abnormal(l, 5)) Then
                          If CDate(Standard(4, 3)) <= CDate(Abnormal(l, 5)) Then                       '�����е�1�ڵ����Ҫ����������
                                r = 0
                                For s = 1 To UBound(CorrectTable, 1)
                                    If InStr(Abnormal(l, 1), CorrectTable(s, 1)) > 0 And CorrectTable(s, 1) <> 0 Then
                                        If CDate(Abnormal(l, 2)) = CDate(CorrectTable(s, 2)) Or CorrectTable(s, 2) = 0 Then
                                            If InStr(Abnormal(l, 3), CorrectTable(s, 3)) > 0 Then
                                                If InStr(Abnormal(l, 4), CorrectTable(s, 4)) > 0 Then
                                                    If InStr(CorrectTable(s, 9), "��1") > 0 Then
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
                    If InStr(Abnormal(l, 3), "����") > 0 Then                                           '��������ǩ��,ƥ��������ϰ����ȷ���Ƿ������Ż�
                        If IsDate(Abnormal(l, 5)) And Abnormal(l, 5) > 0 Then
                            If CDate(Standard(5, 3)) <= CDate(Abnormal(l, 5)) Then
                                r = 0
                                For s = 1 To UBound(CorrectTable, 1)
                                    If InStr(Abnormal(l, 1), CorrectTable(s, 1)) > 0 And CorrectTable(s, 1) <> 0 Then
                                        If CDate(Abnormal(l, 2)) = CDate(CorrectTable(s, 2)) Or CorrectTable(s, 2) = 0 Then
                                            If InStr("����������", CorrectTable(s, 3)) > 0 Then
                                                If InStr(Abnormal(l, 4), CorrectTable(s, 4)) > 0 Then
                                                    If InStr(CorrectTable(s, 9), "��1") > 0 Or InStr(CorrectTable(s, 9), "��5") > 0 Then
                                                        r = s
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                Next
                                If r > 0 Then   '�������1�ڻ��5�ڵģ������������չ�,ǰ��������Դ����밴��������
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
                    If InStr(Abnormal(l, 3), "����") > 0 Then                                   '��������ǩ��
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
' �����ϰ�ʱ��, ������¼����ȥ������
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
' �������գ�������¼��Υ����
            Abnormal(m, o) = Abnormal(m, o) - MorningXX                     'ȥ��������©ǩ��ͳ����
            If MorningX <> 0 Then                                           '��ĩ����ǩ�����2��Υ��
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
'�������˵��Թ����п������꼶�ı䣬������˫�߱���˲��֣���Ҫʱ�޸Ĵ˲��ִ��� '
'                                                                             '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Next
End If
' ���ݻ��ܽ�����ɱ���,Normalswitch=1 �����쳣�����κ��쳣��ʦ��,Normalswitch=1 ����ȫ������
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

'''' V50����ϵͳԤ����ɫ
''
Sub ����ɫ()

'����ɫΪǳ��ɫ��+����ɫ����
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
Sub Ԥ��ɫ()
'
' ��ɫ��+��ɫ����
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

Sub Υ��ɫ()
'
' ��ɫ��+���ɫ����
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

Sub ©ǩɫ()
'
' ���ɫ��+��ɫ��
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
Sub ��עɫ()
'
' ����ɫ��+��ɫ��
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
Sub ��������ɫ()
'
' ����ɫ,������һ�����
'
   With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
End Sub
Sub ������ɫ()
'
' ��ɫ��������һ�����
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
'''' ��ʦ����
            If InStr(THColor(i, 3), "��ʦ����") > 0 Then
'''''ǩ��
                Cells(i, 5).Select
                If THColor(i, 5) = 0 Then                                                          ' ��ɫ�޸�
                   Call ©ǩɫ
                ElseIf IsDate(THColor(i, 5)) Then
                    If CDate(THColor(i, 5)) < CDate(Standard(2, 2)) Then
                        Call ����ɫ
                    ElseIf CDate(THColor(i, 5)) < CDate(Standard(2, 3)) Then
                     Call Ԥ��ɫ
                    Else
                     Call Υ��ɫ
                    End If
                Else
                  Call ��עɫ
                End If
'''''ǩ��
              Cells(i, 6).Select
              If THColor(i, 6) = 0 Then
                  Call ©ǩɫ
              ElseIf IsDate(THColor(i, 6)) Then
                   If CDate(THColor(i, 6)) < CDate(Standard(2, 4)) Then
                      Call Υ��ɫ
                   ElseIf CDate(THColor(i, 6)) < CDate(Standard(2, 5)) Then
                     Call Ԥ��ɫ
                   Else
                     Call ����ɫ
                   End If
              Else
               Call ��עɫ
              End If
''''��ʦ����
    ElseIf InStr(THColor(i, 3), "��ʦ����") > 0 Then
'''''ǩ��
              Cells(i, 5).Select
              If THColor(i, 5) = 0 Then
                 Call ©ǩɫ
              ElseIf IsDate(THColor(i, 5)) Then
                    If CDate(THColor(i, 5)) < CDate(Standard(3, 2)) Then
                      Call ����ɫ
                    ElseIf CDate(THColor(i, 5)) < CDate(Standard(3, 3)) Then
                     Call Ԥ��ɫ
                    Else
                     Call Υ��ɫ
                    End If
             Else
              Call ��עɫ
             End If
'''''ǩ��
              Cells(i, 6).Select
              If THColor(i, 6) = 0 Then
                    Call ©ǩɫ
              ElseIf IsDate(THColor(i, 6)) Then
                    If CDate(THColor(i, 6)) < CDate(Standard(3, 4)) Then
                         Call Υ��ɫ
                    Else
                     Call ����ɫ
                    End If
              Else
                Call ��עɫ
              End If
''''����������
    ElseIf InStr(THColor(i, 3), "����������") > 0 Then
'''''ǩ��
              Cells(i, 5).Select
              If THColor(i, 5) = 0 Then
                 Call ©ǩɫ
              ElseIf IsDate(THColor(i, 5)) Then
                 If CDate(THColor(i, 5)) < CDate(Standard(4, 2)) Then
                   Call ����ɫ
                 ElseIf CDate(THColor(i, 5)) < CDate(Standard(4, 3)) Then
                  Call Ԥ��ɫ
                 Else
                  Call Υ��ɫ
                 End If
              Else
               Call ��עɫ
              End If
'''''ǩ��
                Cells(i, 6).Select
                If THColor(i, 6) = 0 Then
                 Call ©ǩɫ
                ElseIf IsDate(THColor(i, 6)) Then
                    If CDate(THColor(i, 6)) < CDate(Standard(4, 4)) Then
                      Call Υ��ɫ
                    Else
                      Call ����ɫ
                    End If
                Else
                 Call ��עɫ
                End If
''''����������ֻͳ��ǩ��
    ElseIf InStr(THColor(i, 3), "����������") > 0 Then
              Cells(i, 5).Select
              If THColor(i, 5) = 0 Then
                    Call ©ǩɫ
              ElseIf IsDate(THColor(i, 5)) Then
                    If CDate(THColor(i, 5)) < CDate(Standard(5, 2)) Then
                      Call ����ɫ
                    ElseIf CDate(THColor(i, 5)) < CDate(Standard(5, 3)) Then
                     Call Ԥ��ɫ
                    Else
                     Call Υ��ɫ
                    End If
              Else
               Call ��עɫ
              End If
              Cells(i, 6).Select                                                '���ڰ����β�ͳ������ǩ�ˣ������е��Կ��ˣ�����ֻ��ע��ɫ
              If THColor(i, 6) = 0 Then
              ElseIf IsDate(THColor(i, 6)) Then
                Call ����ɫ                                                     'ֻҪǩ����Ϊ����
              Else
                Call ��עɫ
              End If
''''����������ֻͳ��ǩ��
    ElseIf InStr(THColor(i, 3), "����������") > 0 Then
            Cells(i, 5).Select                                                '���ڰ����β�ͳ������ǩ���������е��Կ��ˣ�����ֻ��ע��ɫ
            If THColor(i, 5) = 0 Then
            ElseIf IsDate(THColor(i, 5)) Then
              Call ����ɫ                                                     'ֻҪǩ����Ϊ����
            Else
              Call ��עɫ
            End If
            Cells(i, 6).Select
            If THColor(i, 6) = 0 Then
                   Call ©ǩɫ
             ElseIf IsDate(THColor(i, 6)) Then
                If CDate(THColor(i, 6)) < CDate(Standard(6, 4)) Then
                     Call Υ��ɫ
                Else
                     Call ����ɫ
                End If
            Else
              Call ��עɫ
            End If
    End If
   Next
' ��ͳ������ɫ
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
           Call Υ��ɫ
          Else
           Call ����ɫ
          End If
       Next
     End If
   Next
End Sub
Sub RecoverSource(RECS, RECR, RECC)
'�ĸ����������飬д��ʱ���У�д��ʱ����
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
' ����У׼��ʱ��ָ�
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
                        k = j                                   '��ϰ����ʱ����Correct�л��ж�����¼���������һ��Ϊ׼�������һ��������ƥ�䣬����ǰ�������ڵĹ̶�ֵΪ׼
                    End If
                End If
              End If
            End If
        Next
        If k = 0 Then
            RecSource(i, 4) = ""
        Else
            RecSource(i, 4) = Correct(k, 9)                     ' ����8����ϰ���д�뵽ԭ�ܴε�4�б����
        End If
        If RecSource(i, 5) = 0 Then
            If InStr(RecSource(i, 3), "����������") > 0 Then
            Else
                RecSource(i, 5) = "©ǩ"
            End If
        Else
            If InStr(RecSource(i, 3), "����������") > 0 Then
'                RecSource(i, 5) = ""                           '���ǰ���������ǩ���ģ����ٸ���ԭʼǩ�����,���ǲ����뿼��
            End If
        End If
        If RecSource(i, 6) = 0 Then
            If InStr(RecSource(i, 3), "����������") > 0 Then
            Else
                RecSource(i, 6) = "©ǩ"
            End If
        Else
            If InStr(RecSource(i, 3), "����������") > 0 Then
'               RecSource(i, 6) = ""                            '���ǰ���������ǩ���ģ����ٸ���ԭʼǩ����ǣ�����Ҳ�����뿼��
            End If
        End If
    Next
' ��������д��Sheet
    Range(Cells(1, 1), Cells(RECR, RECC)) = RecSource
' �ϲ�ͳ������������
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
            Call �ϲ�ѡ�е�Ԫ��
            If m > 0 Then
                Cells(j, 7) = "����:" & Changed(m, 1) & Chr(10) & Changed(m, 5) & Chr(10) & "����:" & Changed(m, 6) & Chr(10) & Changed(m, 10)
                m = 0
            End If
            Range(Cells(j, 7), Cells(k, 9)).Select
            Call �ϲ�ѡ�е�Ԫ��
            Range(Cells(j, 10), Cells(k, 12)).Select
            Call �ϲ�ѡ�е�Ԫ��
            If InStr(Cells(2, 3), NameHeadMaster) > 0 Then
                Range(Cells(j, 13), Cells(k, 15)).Select
                Call �ϲ�ѡ�е�Ԫ��
            End If
            j = i
        End If
    Next
' �ϲ�ʱ����
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
            Call �ϲ�ѡ�е�Ԫ��
            l = i
        End If
    Next
End Sub
'
Sub �ϲ�ѡ�е�Ԫ��()
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
Sub ��ʽ��()
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
''''����Ŀ���ļ�
Sub OutToBook(InSource, InRmax, InCmax, InName)
    Dim OutBook As Workbook
    Application.SheetsInNewWorkbook = 1                                                      ' ����1��Sheet
    Set OutBook = Workbooks.Add
    Application.DisplayAlerts = False
        OutBook.SaveAs Filename:=OutFolder & "\" & InName & OutFileFix
        Sheets(1).name = InName & OutFileFix
        Sheets(InName & OutFileFix).Columns("E:F").NumberFormatLocal = TimeFormat            '����ʱ���ʽ
        Sheets(InName & OutFileFix).Columns("B:B").NumberFormatLocal = DateFormat            '�������ڸ�ʽ
        Range(Cells(1, 1), Cells(InRmax, InCmax)).Select
        Call FontSet(NameFont)                                                               '���������ʽ
        Call WriteColorTo(InSource)
        Call RecoverSource(InSource, InRmax, InCmax)
        Call ��ʽ��
'' �������У������ڵ����϶��ղ鿴
        With ActiveWindow
            .SplitColumn = 0
            .SplitRow = 1
        End With
        ActiveWindow.FreezePanes = True
    ActiveWorkbook.Close savechanges:=True
    Application.DisplayAlerts = True
    Set OutBook = Nothing                                                                    'ȡ��OutBook
End Sub
''' ����������ٱ�
Sub OutToLeave(InSource, InRmax, InCmax, OutLeavePath, InName)
    Dim OutBook As Workbook
    Application.SheetsInNewWorkbook = 1                                                      ' ����1��Sheet
    Set OutBook = Workbooks.Add
    ActiveWindow.FreezePanes = False                                                         ' ��ֹ���ᴰ��
    Application.DisplayAlerts = False
        OutBook.SaveAs Filename:=OutLeavePath & "\" & InName & ".xlsx"
        Sheets(1).name = InName
        Range(Cells(1, 1), Cells(InRmax, InCmax)) = InSource                                 '���������Ϣ
        Range(Cells(1, 1), Cells(1, InCmax)).Select
        Call �ϲ�ѡ�е�Ԫ��
        Range(Cells(1, 1), Cells(InRmax, InCmax)).Select
        Call FontSet(NameFont)                                                               '���������ʽ
        Call ��ʽ��
        Range(Cells(1, 1), Cells(1, InCmax)).Select
        TitleSize (20)
        Range(Cells(2, 1), Cells(2, InCmax)).Select
        TitleSize (14)
''''
        Cells(1, 1).Select                                                                   ' �Զ������иߺ��п�
        Selection.CurrentRegion.Select
        Selection.Rows.AutoFit
        Selection.Columns.AutoFit
        Cells(1, 1).Select
    ActiveWorkbook.Close savechanges:=True
    Application.DisplayAlerts = True
    Set OutBook = Nothing                                                                    'ȡ��OutBook
End Sub
Sub TitleSize(TTsize)
'
    With Selection.Font
        .name = "����"
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
    ActiveWindow.FreezePanes = False                                                                            ' ��ֹ���ᴰ��
    LeaveFolder = OutPath & "\" & Format(Now, "yyyy" & "��") & "ͳ���ܱ�"
    ToLeaveName = Format(Now, "yyyy" & "��") & "����ܱ�"
    LeaveFile = LeaveFolder & "\" & ToLeaveName & ".xlsx"
    Dim SFO As Object
    Set SFO = CreateObject("Scripting.FileSystemObject")                                                        '��SFOΪ�ļ��ж������
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
        For j = 3 To LORow                                                                                      ' ǰ�����Ǳ��⣬���Ͽյģ���˲��ܴ�1��ʼ
            If IsDate(LeaveOld(j, 2)) Then                                                                      ' ��Ϊ���ڸ�ʽ������Ҫ��ʽ�����ٶԱ�
               LeaveOld(j, 2) = Format(LeaveOld(j, 2), "m" & "��" & "d" & "��")
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
    Call FontSet(NameFont)                                                                                      ' ���������ʽ
    Call ��ʽ��
    Sheets(ToLeaveName).Range(Cells(1, 1), Cells(1, 4)).Select
    TitleSize (20)
    Sheets(ToLeaveName).Range(Cells(2, 1), Cells(2, 4)).Select
    TitleSize (14)
''''
    Cells(1, 1).Select                                                                                         ' �Զ������иߺ��п�
    Selection.CurrentRegion.Select
    Selection.Rows.AutoFit
    Selection.Columns.AutoFit
    Cells(1, 1).Select
    Workbooks(ToLeaveName).Close savechanges:=True
End Sub
