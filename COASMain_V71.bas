Attribute VB_Name = "COASMain_V71"
' ��Ŀ��COASMain
' �汾��V71
' ���ߣ�����
' ��λ��ɽ��ʡƽԭ�ص�һ��ѧ
' ���䣺fengzhenhua@outlook.com
' ���ͣ�https://fengzhenhua-vip.github.io
' ��ҳ��https://github.com/fengzhenhua-vip
' ��Ȩ��2021��7��13��--2021��7��16��
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
'
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
Public CorrectTime As Variant
Public CTERmax, CTECmax As Integer
Public SelfStudyTable As Variant
Public SSRmax, SSCmax As Integer
Public DuiShiA(1 To RowMax, 1 To SubColMax) As Variant
Public DuiShiB(1 To RowMax, 1 To SubColMax) As Variant
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
'
Sub COASMain()
    Application.ScreenUpdating = False
    Call COAConfigSet
    Call GetSource
    Call GetDuiShiBiao
    Call GetHoliday
    Call GetLeaveBook
    Call COAExecute
    Application.ScreenUpdating = True
End Sub

Sub COAConfigSet()
 '   ConfigPath = "D:\����ϵͳ"
    ConfigPath = CreateObject("WScript.Shell").specialfolders("Desktop") & "\����ϵͳ"
    ConfigFolder = ConfigPath & "\" & "����ϵͳ����"
    OutPath = ConfigPath & "\" & Format(Now, "yyyy" & "��") & "����"
    DateFormat = "m""��""d""��"";@"
    TimeFormat = "h:mm;@"
    NameFont = "����"
    NameHeadMasterUN = "�쳣������"
    NameTeacherUN = "�쳣��ʦ"
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
    StopSymbol = "*"
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
End Sub
Sub GetSource()
    Set ConfigBook = GetObject(ConfigFile)
    ViRmax = ConfigBook.Sheets(ConfigSheet1).Cells(RowMax, 1).End(xlUp).Row
    ViCmax = ConfigBook.Sheets(ConfigSheet1).Cells(1, ColMax).End(xlToLeft).Column
    VipSource = ConfigBook.Sheets(ConfigSheet1).Range(ConfigBook.Sheets(ConfigSheet1).Cells(1, 1), ConfigBook.Sheets(ConfigSheet1).Cells(ViRmax, ViCmax))
' ȡ��ԭʼ������������
    DeRmax = Sheets(1).Cells(RowMax, 1).End(xlUp).Row
    DeCmax = 13
' ���ɱ�׼ͨ�ÿ��������ʽ
    Dim DeLiSource As Variant
    Dim i, j, k, m As Integer
    DeLiSource = Sheets(1).Range(Cells(1, 3), Cells(DeRmax, DeCmax)).Value
    j = 1
    Source(j, 1) = "����": Source(j, 2) = "����(��)": Source(j, 3) = "���": Source(j, 4) = "��ϰ": Source(j, 5) = "ǩ��": Source(j, 6) = "ǩ��"
    Source(j, 7) = "�ϳ�": Source(j, 8) = "����": Source(j, 9) = "��©": Source(j, 10) = "�³�": Source(j, 11) = "����": Source(j, 12) = "��©"
    Source(j, 13) = "���": Source(j, 14) = "����": Source(j, 15) = "��©"
    DateMin = CDate(DeLiSource(3, 4)): DateMax = DateMin
    For i = 3 To DeRmax                                                                                      'DeLi������ǰ���кϲ��������Ե�2���ǿյĲ�������
            m = 0
            For k = 1 To ViRmax
                If InStr(VBA.Trim(Mid(DeLiSource(i, 1), 1, 1)) & VBA.Trim(Mid(DeLiSource(i, 1), 2)), VipSource(k, 1)) > 0 Then
                    m = 1
                End If
            Next
            If m = 0 Then
                j = j + 1
                Source(j, 1) = VBA.Trim(Mid(DeLiSource(i, 1), 1, 1)) & VBA.Trim(Mid(DeLiSource(i, 1), 2))   '�򵥺�ȥ�����������м�Ŀո�,����ʹ��ͳһ�Ŀո�����
                Source(j, 2) = CDate(DeLiSource(i, 4))
                If DateMin > Source(j, 2) Then
                    DateMin = Source(j, 2)
                End If
                If DateMax < Source(j, 2) Then
                    DateMax = Source(j, 2)
                End If
                If InStr(DeLiSource(i, 7), "����") Then
                    Source(j, 3) = 1
                ElseIf InStr(DeLiSource(i, 7), "����") Then
                    Source(j, 3) = 2
                ElseIf InStr(DeLiSource(i, 7), "����") Then
                    Source(j, 3) = 3
                End If
                If InStr(DeLiSource(i, 7), "��ʦ") Then
                    Source(j, 23) = 1
                ElseIf InStr(DeLiSource(i, 7), "������") Then
                    Source(j, 23) = 2
                End If
                Source(j, 4) = DeLiSource(i, 5)
                If DeLiSource(i, 10) > 0 Then
                    Source(j, 5) = CDate(DeLiSource(i, 10))
                End If
                If DeLiSource(i, 11) > 0 Then
                    Source(j, 6) = CDate(DeLiSource(i, 11))
                End If
            End If
    Next
    SRowMax = j
' ����Ŀ��Excel�ļ��ĺ�׺������ļ���
    OutFileFix = "��" & Format(DateMin, "yyyy" & "��" & "m" & "��" & "d" & "��") & "-" & Format(DateMax, "yyyy" & "��" & "m" & "��" & "d" & "��") & "��"
    OutFolder = OutPath & "\" & Format(DateMax, "m" & "��" & "d" & "��") & "��ʽ�ϱ�"
    If SFO.FolderExists(OutFolder) = False Then
       MkDir OutFolder
    End If
End Sub
Sub GetDuiShiBiao()
    Dim i, j, k, m As Integer
    CTERmax = ConfigBook.Sheets(OriginalSheet2).Cells(RowMax, 1).End(xlUp).Row
    CTECmax = ConfigBook.Sheets(OriginalSheet2).Cells(1, ColMax).End(xlToLeft).Column
    CorrectTime = ConfigBook.Sheets(OriginalSheet2).Range(ConfigBook.Sheets(OriginalSheet2).Cells(1, 1), ConfigBook.Sheets(OriginalSheet2).Cells(CTERmax, CTECmax))
' ��ȡ��ʱ��A
    CGRmax = ConfigBook.Sheets(OriginalSheet4).Cells(RowMax, 1).End(xlUp).Row                                                                    '���뻻�α�
    CGCmax = ConfigBook.Sheets(OriginalSheet4).Cells(1, ColMax).End(xlToLeft).Column
    Change = ConfigBook.Sheets(OriginalSheet4).Range(ConfigBook.Sheets(OriginalSheet4).Cells(1, 1), ConfigBook.Sheets(OriginalSheet4).Cells(CGRmax, CGCmax))
    k = 0
    For i = 2 To CGRmax
        If (DateMin < CDate(Change(i, 2)) And CDate(Change(i, 2)) < DateMax) Or (DateMin < CDate(Change(i, 7)) And CDate(Change(i, 7)) < DateMax) Then
            k = k + 1
            DuiShiA(k, 1) = Change(i, 6)
            DuiShiA(k, 2) = CDate(Change(i, 2))
            If InStr(Change(i, 3), "����") Then
                DuiShiA(k, 3) = 1
            ElseIf InStr(Change(i, 3), "����") Then
                DuiShiA(k, 3) = 2
            ElseIf InStr(Change(i, 3), "����") Then
                DuiShiA(k, 3) = 3
            End If
            DuiShiA(k, 4) = Change(i, 4)
            If InStr(Change(i, 5), "B") > 0 Then
                DuiShiA(k, 5) = CDate(CorrectTime(2, 2)): DuiShiA(k, 25) = "��1��"
            ElseIf InStr(Change(i, 5), "C") > 0 Then
                DuiShiA(k, 8) = CDate(CorrectTime(3, 5)): DuiShiA(k, 25) = "��5��"
            ElseIf InStr(Change(i, 5), "D") > 0 Then
                 DuiShiA(k, 5) = CDate(CorrectTime(6, 2)): DuiShiA(k, 25) = "��6��"
            ElseIf InStr(Change(i, 5), "E") > 0 Then
                DuiShiA(k, 8) = CDate(CorrectTime(5, 5)): DuiShiA(k, 25) = "��9��"
            Else
                DuiShiA(k, 25) = Change(i, 5)
            End If
            k = k + 1
            DuiShiA(k, 1) = Change(i, 1)
            DuiShiA(k, 2) = CDate(Change(i, 7))
            If InStr(Change(i, 8), "����") Then
                DuiShiA(k, 3) = 1
            ElseIf InStr(Change(i, 8), "����") Then
                DuiShiA(k, 3) = 2
            ElseIf InStr(Change(i, 8), "����") Then
                DuiShiA(k, 3) = 3
            End If
            DuiShiA(k, 4) = Change(i, 9)
            If InStr(Change(i, 10), "B") > 0 Then
                DuiShiA(k, 5) = CDate(CorrectTime(2, 2)): DuiShiA(k, 25) = "��1��"
            ElseIf InStr(Change(i, 10), "C") > 0 Then
                DuiShiA(k, 8) = CDate(CorrectTime(3, 5)): DuiShiA(k, 25) = "��5��"
            ElseIf InStr(Change(i, 10), "D") > 0 Then
                DuiShiA(k, 5) = CDate(CorrectTime(6, 2)): DuiShiA(k, 25) = "��6��"
            ElseIf InStr(Change(i, 10), "E") > 0 Then
                DuiShiA(k, 8) = CDate(CorrectTime(5, 5)): DuiShiA(k, 25) = "��9��"
            Else
                DuiShiA(k, 25) = Change(i, 10)
            End If
' ���������Ϣ
            DuiShiA(k - 1, 24) = "����:" & DuiShiA(k, 1) & "Chr(13)" & DuiShiA(k - 1, 25) & _
                                 "����:" & DuiShiA(k - 1, 1) & "Chr(13)" & DuiShiA(k, 25)
            DuiShiA(k, 24) = "����:" & DuiShiA(k - 1, 1) & "Chr(13)" & DuiShiA(k, 25) & _
                             "����:" & DuiShiA(k, 2) & "Chr(13)" & DuiShiA(k - 1, 25)
'����ζ�����������ǩ�������Ż�
            If (InStr(Change(i, 5), "B") > 0 Or InStr(Change(i, 5), "C") > 0) Then
                k = k + 1
                DuiShiA(k, 1) = Change(i, 6)
                DuiShiA(k, 2) = CDate(Change(i, 2))
                DuiShiA(k, 3) = 2 ' "����"
                DuiShiA(k, 4) = Change(i, 4)
                DuiShiA(k, 6) = CDate(CorrectTime(4, 3))
                DuiShiA(k, 25) = "��"
            End If
            If InStr(Change(i, 10), "B") > 0 Or InStr(Change(i, 10), "C") > 0 Then
                k = k + 1
                DuiShiA(k, 1) = Change(i, 1)
                DuiShiA(k, 2) = CDate(Change(i, 7))
                DuiShiA(k, 3) = 2 '"����"
                DuiShiA(k, 4) = Change(i, 9)
                DuiShiA(k, 6) = CDate(CorrectTime(4, 3))
                DuiShiA(k, 25) = "��"
            End If
        End If
    Next
    DSRAmax = k
' ��ȡ��ʱ��B
    SSRmax = ConfigBook.Sheets(OriginalSheet1).Cells(RowMax, 1).End(xlUp).Row
    SSCmax = ConfigBook.Sheets(OriginalSheet1).Cells(1, ColMax).End(xlToLeft).Column
    SelfStudyTable = ConfigBook.Sheets(OriginalSheet1).Range(ConfigBook.Sheets(OriginalSheet1).Cells(1, 1), ConfigBook.Sheets(OriginalSheet1).Cells(SSRmax, SSCmax))
    k = 0
    For i = 2 To SSRmax
        For j = 2 To 6
            If InStr(SelfStudyTable(i, j), "B") > 0 Or InStr(SelfStudyTable(i, j), "C") > 0 Or InStr(SelfStudyTable(i, j), "D") > 0 Or InStr(SelfStudyTable(i, j), "E") > 0 Then
                If InStr(SelfStudyTable(i, j), "B") > 0 Then
                    k = k + 1
                    DuiShiB(k, 1) = SelfStudyTable(i, 1)
                    DuiShiB(k, 3) = 1 '"����"
                    DuiShiB(k, 4) = SelfStudyTable(1, j)
                    DuiShiB(k, 5) = CDate(CorrectTime(2, 2))
                    DuiShiB(k, 25) = "��1��"
                End If
                If InStr(SelfStudyTable(i, j), "C") > 0 Then
                    k = k + 1
                    DuiShiB(k, 1) = SelfStudyTable(i, 1)
                    DuiShiB(k, 3) = 1 '"����"
                    DuiShiB(k, 4) = SelfStudyTable(1, j)
                    DuiShiB(k, 8) = CDate(CorrectTime(3, 5))
                    If InStr(SelfStudyTable(i, j), "B") > 0 Then
                        DuiShiB(k, 25) = "��1,5��"
                    Else
                        DuiShiB(k, 25) = "��5��"
                    End If
                End If
                If InStr(SelfStudyTable(i, j), "D") > 0 Then
                    k = k + 1
                    DuiShiB(k, 1) = SelfStudyTable(i, 1)
                    DuiShiB(k, 3) = 2 '"����"
                    DuiShiB(k, 4) = SelfStudyTable(1, j)
                    DuiShiB(k, 5) = CDate(CorrectTime(6, 2))
                    DuiShiB(k, 25) = "��6��"
                ElseIf InStr(SelfStudyTable(i, j), "B") > 0 Or InStr(SelfStudyTable(i, j), "C") > 0 Then
                    k = k + 1
                    DuiShiB(k, 1) = SelfStudyTable(i, 1)
                    DuiShiB(k, 3) = 2 '"����"
                    DuiShiB(k, 4) = SelfStudyTable(1, j)
                    DuiShiB(k, 6) = CDate(CorrectTime(4, 3))
                    DuiShiB(k, 25) = "��"                                           '����ϰ������������ǩ�������Ż�
                End If
                If InStr(SelfStudyTable(i, j), "E") > 0 Then
                    k = k + 1
                    DuiShiB(k, 1) = SelfStudyTable(i, 1)
                    DuiShiB(k, 3) = 2 '"����"
                    DuiShiB(k, 4) = SelfStudyTable(1, j)
                    DuiShiB(k, 8) = CDate(CorrectTime(5, 5))
                    If InStr(SelfStudyTable(i, j), "D") > 0 Then
                        DuiShiB(k, 25) = "��6,9��"
                    Else
                        DuiShiB(k, 25) = "��9��"
                    End If
                End If
            End If
        Next
    Next
    DSRBmax = k
' ʹ�ö���У׼���DuiShiA��DuiShiBУ׼
    RCTRmax = ConfigBook.Sheets(OriginalSheet3).Cells(RowMax, 1).End(xlUp).Row
    RCTCmax = ConfigBook.Sheets(OriginalSheet3).Cells(1, ColMax).End(xlToLeft).Column
    ReCorrectTable = ConfigBook.Sheets(OriginalSheet3).Range(ConfigBook.Sheets(OriginalSheet3).Cells(1, 1), ConfigBook.Sheets(OriginalSheet3).Cells(RCTRmax, RCTCmax))
    For i = 2 To RCTRmax
        If InStr(ReCorrectTable(i, 3), "����") Then
            ReCorrectTable(i, 3) = 1
        ElseIf InStr(ReCorrectTable(i, 3), "����") Then
            ReCorrectTable(i, 3) = 2
        ElseIf InStr(ReCorrectTable(i, 3), "����") Then
            ReCorrectTable(i, 3) = 3
        End If
        If Len(ReCorrectTable(i, 2)) = 0 Then
            For j = 1 To DSRBmax
                If DuiShiB(j, 1) = ReCorrectTable(i, 1) And DuiShiB(j, 3) = ReCorrectTable(i, 3) And DuiShiB(j, 4) = ReCorrectTable(i, 4) Then
                    For m = 5 To 8
                        DuiShiB(j, m) = CDate(ReCorrectTable(i, m))
                    Next
                    DuiShiB(j, 25) = DuiShiB(j, 25) & "��"                           '����У׼����
                End If
            Next
        Else
            For j = 1 To DSRAmax
                If DuiShiB(j, 1) = ReCorrectTable(i, 1) And DuiShiB(j, 2) = CDate(ReCorrectTable(i, 2)) And DuiShiB(j, 3) = ReCorrectTable(i, 3) Then
                    For m = 5 To 8
                        DuiShiA(j, m) = CDate(ReCorrectTable(i, m))
                    Next
                    DuiShiA(j, 25) = DuiShiA(j, 25) & "��"
                End If
            Next
        End If
    Next
End Sub
'
Sub GetHoliday()
    HRmax = ConfigBook.Sheets(ConfigSheet2).Cells(RowMax, 1).End(xlUp).Row
    HCmax = ConfigBook.Sheets(ConfigSheet2).Cells(1, ColMax).End(xlToLeft).Column
    Holiday = ConfigBook.Sheets(ConfigSheet2).Range(ConfigBook.Sheets(ConfigSheet2).Cells(1, 1), ConfigBook.Sheets(ConfigSheet2).Cells(HRmax, HCmax))
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
       If IsDate(Holiday(i, 3)) Then
           If DateMin <= CDate(Holiday(i, 4)) Then
               If DateMin <= CDate(Holiday(i, 3)) Then
                DateX = CDate(Holiday(i, 3))
               Else
                DateX = DateMin
               End If
               If DateMax <= CDate(Holiday(i, 4)) Then
                DateY = DateMax
               Else
                DateY = CDate(Holiday(i, 4))
               End If
               If InStr(Holiday(i, 1), "��") > 0 Or InStr(Holiday(i, 1), "������") > 0 Or InStr(Holiday(i, 1), "������") > 0 Then
                   For a = 3 To GroupRow
                         If Len(Holiday(i, 5)) > 0 Then
                             DateZ = DateX
                             Do While DateZ <= DateY
                                  HARowMax = HARowMax + 1
                                  HolidayA(HARowMax, 1) = TeacherGroup(a, GroupColum)
                                  HolidayA(HARowMax, 2) = DateZ
                                  If InStr(TeacherGroup(a, GroupColum + TeGrZhiWu), "��ʦ") > 0 Then
                                   HolidayA(HARowMax, 23) = 1
                                  ElseIf InStr(TeacherGroup(a, GroupColum + TeGrZhiWu), "������") > 0 Then
                                   HolidayA(HARowMax, 23) = 2
                                  End If
                                  If Len(Holiday(i, 5)) > 0 Then
                                    HolidayA(HARowMax, 5) = Holiday(i, 5)
                                    HolidayA(HARowMax, 6) = Holiday(i, 5)
                                  End If
                                  DateZ = DateZ + 1
                             Loop
                         Else
                             For k = 6 To 10 Step 2
                                 DateZ = DateX
                                 If Len(Holiday(i, k)) > 0 Or Len(Holiday(i, k + 1)) > 0 Then
                                  Do While DateZ <= DateY
                                       HARowMax = HARowMax + 1
                                       HolidayA(HARowMax, 1) = TeacherGroup(a, GroupColum)
                                       HolidayA(HARowMax, 2) = DateZ
                                       If InStr(Holiday(1, k), "����") > 0 Then
                                           HolidayA(HARowMax, 3) = 1
                                       ElseIf InStr(Holiday(1, k), "����") > 0 Then
                                           HolidayA(HARowMax, 3) = 2
                                       ElseIf InStr(Holiday(1, k), "����") > 0 Then
                                           HolidayA(HARowMax, 3) = 3
                                       End If
                                       If InStr(TeacherGroup(a, GroupColum + TeGrZhiWu), "��ʦ") > 0 Then
                                        HolidayA(HARowMax, 23) = 1
                                       ElseIf InStr(TeacherGroup(a, GroupColum + TeGrZhiWu), "������") > 0 Then
                                        HolidayA(HARowMax, 23) = 2
                                       End If
                                       If Len(Holiday(i, k)) > 0 Then HolidayA(HARowMax, 5) = Holiday(i, k)
                                       If Len(Holiday(i, k + 1)) > 0 Then HolidayA(HARowMax, 6) = Holiday(i, k + 1)
                                       DateZ = DateZ + 1
                                  Loop
                                 End If
                              Next
                         End If
                   Next
                 Else
                   If Len(Holiday(i, 5)) > 0 Then
                      Do While DateX <= DateY
                           HARowMax = HARowMax + 1
                           HolidayA(HARowMax, 1) = Holiday(i, 1)
                           HolidayA(HARowMax, 2) = DateX
                           If InStr(Holiday(i, 2), "����") > 0 Then
                               HolidayA(HARowMax, 3) = 1
                           ElseIf InStr(Holiday(i, 2), "����") > 0 Then
                               HolidayA(HARowMax, 3) = 2
                           ElseIf InStr(Holiday(i, 2), "����") > 0 Then
                               HolidayA(HARowMax, 3) = 3
                           End If
                           If InStr(Holiday(i, 2), "��ʦ") > 0 Then
                               HolidayA(HARowMax, 23) = 1
                           ElseIf InStr(Holiday(i, 2), "������") > 0 Then
                               HolidayA(HARowMax, 23) = 2
                           End If
                           If Len(Holiday(i, 5)) > 0 Then
                            HolidayA(HARowMax, 5) = Holiday(i, 5)
                            HolidayA(HARowMax, 6) = Holiday(i, 5)
                           End If
                           DateX = DateX + 1
                      Loop
                   Else
                      For k = 6 To 10 Step 2
                          DateZ = DateX
                          If Len(Holiday(i, k)) > 0 Or Len(Holiday(i, k + 1)) > 0 Then
                           Do While CDate(DateZ) <= CDate(DateY)
                                HARowMax = HARowMax + 1
                                HolidayA(HARowMax, 1) = Holiday(i, 1)
                                HolidayA(HARowMax, 2) = DateZ
                                If InStr(Holiday(1, k), "����") > 0 Then
                                   HolidayA(HARowMax, 3) = 1
                                ElseIf InStr(Holiday(1, k), "����") > 0 Then
                                   HolidayA(HARowMax, 3) = 2
                                ElseIf InStr(Holiday(1, k), "����") > 0 Then
                                   HolidayA(HARowMax, 3) = 3
                                End If
                                If InStr(Holiday(i, 2), "��ʦ") > 0 Then
                                   HolidayA(HARowMax, 23) = 1
                                ElseIf InStr(Holiday(i, 2), "������") > 0 Then
                                   HolidayA(HARowMax, 23) = 2
                                End If
                                If Len(Holiday(i, k)) > 0 Then HolidayA(HARowMax, 5) = Holiday(i, k)
                                If Len(Holiday(i, k + 1)) > 0 Then HolidayA(HARowMax, 6) = Holiday(i, k + 1)
                                DateZ = DateZ + 1
                           Loop
                          End If
                       Next
                    End If
                 End If
           End If
    Else
' ��ȡ�����ڸ�ʽ��ٱ�HolidayB
           For j = 1 To 7
            If Holiday(i, 3) = Choose(j, "һ", "��", "��", "��", "��", "��", "��") Then
               WeekX = j
            End If
           Next
           For j = 1 To 7
            If Holiday(i, 4) = Choose(j, "һ", "��", "��", "��", "��", "��", "��") Then
               WeekY = j
            End If
           Next
           If InStr(Holiday(i, 1), "��") > 0 Or InStr(Holiday(i, 1), "������") > 0 Or InStr(Holiday(i, 1), "������") > 0 Then
               For a = 3 To GroupRow
                   For k = 6 To 10 Step 2
                       WeekZ = WeekX
                       If Len(Holiday(i, k)) > 0 Or Len(Holiday(i, k + 1)) > 0 Then
                           Do While WeekZ <= WeekY
                               HBRowMax = HBRowMax + 1
                               HolidayB(HBRowMax, 1) = TeacherGroup(a, GroupColum)
                               If InStr(Holiday(1, k), "����") > 0 Then
                                   HolidayB(HBRowMax, 3) = 1
                                ElseIf InStr(Holiday(1, k), "����") > 0 Then
                                   HolidayB(HBRowMax, 3) = 2
                                ElseIf InStr(Holiday(1, k), "����") > 0 Then
                                   HolidayB(HBRowMax, 3) = 3
                                End If
                               If InStr(TeacherGroup(a, GroupColum + TeGrZhiWu), "��ʦ") Then
                                   HolidayB(HBRowMax, 23) = 1
                               ElseIf InStr(TeacherGroup(a, GroupColum + TeGrZhiWu), "������") Then
                                   HolidayB(HBRowMax, 23) = 2
                               End If
                               Select Case WeekZ
                                   Case Is = 1
                                       HolidayB(HBRowMax, 4) = "һ"
                                   Case Is = 2
                                       HolidayB(HBRowMax, 4) = "��"
                                   Case Is = 3
                                       HolidayB(HBRowMax, 4) = "��"
                                   Case Is = 4
                                       HolidayB(HBRowMax, 4) = "��"
                                   Case Is = 5
                                       HolidayB(HBRowMax, 4) = "��"
                                   Case Is = 6
                                       HolidayB(HBRowMax, 4) = "��"
                                   Case Is = 7
                                       HolidayB(HBRowMax, 4) = "��"
                               End Select
                               If Len(Holiday(i, k)) > 0 Then HolidayB(HBRowMax, 5) = Holiday(i, k)
                               If Len(Holiday(i, k + 1)) > 0 Then HolidayB(HBRowMax, 6) = Holiday(i, k + 1)
                               WeekZ = WeekZ + 1
                           Loop
                       End If
                    Next
                Next
           Else
                For k = 6 To 10 Step 2
                  WeekZ = WeekX
                  If Len(Holiday(i, k)) > 0 Or Len(Holiday(i, k + 1)) > 0 Then
                      Do While WeekZ <= WeekY
                          HBRowMax = HBRowMax + 1
                          HolidayB(HBRowMax, 1) = Holiday(i, 1)
                          If InStr(Holiday(1, k), "����") > 0 Then
                              HolidayB(HBRowMax, 3) = 1
                           ElseIf InStr(Holiday(1, k), "����") > 0 Then
                              HolidayB(HBRowMax, 3) = 2
                           ElseIf InStr(Holiday(1, k), "����") > 0 Then
                              HolidayB(HBRowMax, 3) = 3
                           End If
                           If InStr(Holiday(i, 2), "��ʦ") Then
                               HolidayB(HBRowMax, 23) = 1
                           ElseIf InStr(Holiday(i, 2), "������") Then
                               HolidayB(HBRowMax, 23) = 2
                           End If
                          Select Case WeekZ
                              Case Is = 1
                                  HolidayB(HBRowMax, 4) = "һ"
                              Case Is = 2
                                  HolidayB(HBRowMax, 4) = "��"
                              Case Is = 3
                                  HolidayB(HBRowMax, 4) = "��"
                              Case Is = 4
                                  HolidayB(HBRowMax, 4) = "��"
                              Case Is = 5
                                  HolidayB(HBRowMax, 4) = "��"
                              Case Is = 6
                                  HolidayB(HBRowMax, 4) = "��"
                              Case Is = 7
                                  HolidayB(HBRowMax, 4) = "��"
                          End Select
                          If Len(Holiday(i, k)) > 0 Then HolidayB(HBRowMax, 5) = Holiday(i, k)
                          If Len(Holiday(i, k + 1)) > 0 Then HolidayB(HBRowMax, 6) = Holiday(i, k + 1)
                          WeekZ = WeekZ + 1
                      Loop
                  End If
               Next
           End If
       End If
    Next
End Sub
'
Sub COAExecute()
    Dim i, j, k, m, n As Integer
    STRmax = ConfigBook.Sheets(ConfigSheet3).Cells(RowMax, 1).End(xlUp).Row
    STCmax = ConfigBook.Sheets(ConfigSheet3).Cells(1, ColMax).End(xlToLeft).Column
    Standard = ConfigBook.Sheets(ConfigSheet3).Range(ConfigBook.Sheets(ConfigSheet3).Cells(1, 1), ConfigBook.Sheets(ConfigSheet3).Cells(STRmax, STCmax))
    Dim DuiShiTemp(1 To 2, 5 To 8) As Variant
    S1RowMax = 1: S2RowMax = 1
    For i = 2 To SRowMax
' �˶�������
        m = 0
        For j = 1 To HARowMax
            If Len(HolidayA(j, 1)) > 0 And (InStr(Source(i, 1), HolidayA(j, 1)) > 0 Or InStr(HolidayA(j, 1), "*") > 0) And InStr(Source(i, 2), HolidayA(j, 2)) > 0 And InStr(Source(i, 3), HolidayA(j, 3)) > 0 Then
                 If Len(HolidayA(j, 5)) > 0 Then Source(i, 5) = HolidayA(j, 5): m = 1: Source(i, 21) = 5
                 If Len(HolidayA(j, 6)) > 0 Then Source(i, 6) = HolidayA(j, 6): m = 1: Source(i, 22) = 5
            End If
        Next
        If m = 0 Then
            For j = 1 To HBRowMax
                If Len(HolidayB(j, 1)) > 0 And (InStr(Source(i, 1), HolidayB(j, 1)) > 0 Or InStr(HolidayB(j, 1), "*") > 0) And InStr(Source(i, 3), HolidayB(j, 3)) > 0 And InStr(Source(i, 4), HolidayB(j, 4)) > 0 Then
                    If Len(HolidayB(j, 5)) > 0 Then Source(i, 5) = HolidayB(j, 5): m = 1: Source(i, 21) = 5
                    If Len(HolidayB(j, 6)) > 0 Then Source(i, 6) = HolidayB(j, 6): m = 1: Source(i, 22) = 5
                End If
            Next
        End If
' �ݶ�ʱ��DuiShiA,DuiShiBУ׼Source
        m = 0: Erase DuiShiTemp
        For j = 1 To DSRAmax
            If InStr(Source(i, 1), DuiShiA(j, 1)) > 0 And InStr(Source(i, 2), DuiShiA(j, 2)) > 0 And InStr(Source(i, 3), DuiShiA(j, 3)) > 0 Then
                If Len(Source(i, 21)) = 0 And Len(Source(i, 5)) > 0 Then
                    Source(i, 5) = CDate(Source(i, 5)) + CDate(DuiShiA(j, 5)) - CDate(DuiShiA(j, 6))
                End If
                If Len(Source(i, 22)) = 0 And Len(Source(i, 6)) > 0 Then
                    Source(i, 6) = CDate(Source(i, 6)) + CDate(DuiShiA(j, 7)) - CDate(DuiShiA(j, 8))
                End If
                Source(i, 24) = DuiShiA(j, 24)
                Source(i, 25) = DuiShiA(j, 25)
                For k = 5 To 8
                    DuiShiTemp(1, k) = DuiShiA(j, k)
                Next
                m = 1
            End If
        Next
        If m = 0 Then
            For j = 1 To DSRBmax
                If InStr(Source(i, 1), DuiShiB(j, 1)) > 0 And InStr(Source(i, 3), DuiShiB(j, 3)) > 0 And InStr(Source(i, 4), DuiShiB(j, 4)) > 0 Then
                    If Len(Source(i, 21)) = 0 And Len(Source(i, 5)) > 0 Then
                        Source(i, 5) = CDate(Source(i, 5)) + CDate(DuiShiB(j, 5)) - CDate(DuiShiB(j, 6))
                    End If
                    If Len(Source(i, 22)) = 0 And Len(Source(i, 6)) > 0 Then
                        Source(i, 6) = CDate(Source(i, 6)) + CDate(DuiShiB(j, 7)) - CDate(DuiShiB(j, 8))
                    End If
                    Source(i, 25) = DuiShiB(j, 25)
                    For k = 5 To 8
                        DuiShiTemp(1, k) = DuiShiB(j, k)
                    Next
                    m = 1
                End If
            Next
        End If
' ���ɿ������
        Select Case Source(i, 23)
            Case Is = 1                                                                         '���˽�ʦ,����������
                Select Case Source(i, 3)
                    Case Is = 1                                                                 '��������
                         If Len(Source(i, 21)) = 0 Then
                            If Len(Source(i, 5)) = 0 Then
                                Source(i, 9) = 1: Source(i, 5) = "©ǩ": Source(i, 21) = 4
                            Else
                                If Source(i, 5) <= CDate(Standard(2, 2)) Then Source(i, 21) = 1
                                If CDate(Standard(2, 2)) < CDate(Source(i, 5)) And CDate(Source(i, 5)) <= CDate(Standard(2, 3)) Then Source(i, 21) = 2
                                If CDate(Standard(2, 3)) < CDate(Source(i, 5)) Then Source(i, 7) = 1: Source(i, 21) = 3
                            End If
                         End If
                         If Len(Source(i, 22)) = 0 Then
                            If Len(Source(i, 6)) = 0 Then
                                Source(i, 9) = Source(i, 9) + 1: Source(i, 6) = "©ǩ": Source(i, 22) = 4
                            Else
                                If CDate(Standard(2, 5)) < CDate(Source(i, 6)) Then Source(i, 22) = 1
                                If CDate(Standard(2, 4)) < CDate(Source(i, 6)) And CDate(Source(i, 6)) <= CDate(Standard(2, 5)) Then Source(i, 22) = 2
                                If CDate(Source(i, 6)) <= CDate(Standard(2, 4)) Then Source(i, 8) = 1: Source(i, 22) = 3
                            End If
                         End If
                    Case Is = 2                                                                                 '��������
                         If Len(Source(i, 21)) = 0 Then
                            If Len(Source(i, 5)) = 0 Then
                                Source(i, 12) = 1: Source(i, 5) = "©ǩ": Source(i, 21) = 4
                            Else
                                If Source(i, 5) <= CDate(Standard(3, 2)) Then Source(i, 21) = 1
                                If CDate(Standard(3, 2)) < CDate(Source(i, 5)) And CDate(Source(i, 5)) <= CDate(Standard(3, 3)) Then Source(i, 21) = 2
                                If CDate(Standard(3, 3)) < CDate(Source(i, 5)) Then Source(i, 10) = 1: Source(i, 21) = 3
                            End If
                         End If
                         If Len(Source(i, 22)) = 0 Then
                            If Len(Source(i, 6)) = 0 Then
                                Source(i, 12) = Source(i, 12) + 1: Source(i, 6) = "©ǩ": Source(i, 22) = 4
                            Else
                                If CDate(Standard(3, 5)) < CDate(Source(i, 6)) Then Source(i, 22) = 1
                                If CDate(Standard(3, 4)) < CDate(Source(i, 6)) And CDate(Source(i, 6)) <= CDate(Standard(3, 5)) Then Source(i, 22) = 2
                                If CDate(Source(i, 6)) <= CDate(Standard(3, 4)) Then Source(i, 11) = 1: Source(i, 22) = 3
                            End If
                         End If
                End Select
                If m = 1 Then                                                                                      '�ָ�ǩ��ǩ��ʱ��
                    If CInt(Source(i, 21)) < 4 Then Source(i, 5) = CDate(Source(i, 5)) - CDate(DuiShiTemp(1, 5)) + CDate(DuiShiTemp(1, 6))
                    If CInt(Source(i, 22)) < 4 Then Source(i, 6) = CDate(Source(i, 6)) - CDate(DuiShiTemp(1, 7)) + CDate(DuiShiTemp(1, 8))
                End If
                S1RowMax = S1RowMax + 1                                                                            '���ְ��1����ʦ��
                For j = 1 To UBound(Source1, 2)
                    Source1(S1RowMax, j) = Source(i, j)
                Next
            Case Is = 2                                                                                             '���˰�����
                Select Case Source(i, 3)
                    Case Is = 1                                                                                     '��������
                         If Len(Source(i, 21)) = 0 Then
                            If Len(Source(i, 5)) = 0 Then
                                Source(i, 9) = 1: Source(i, 5) = "©ǩ": Source(i, 21) = 4
                            Else
                                If CDate(Source(i, 5)) <= CDate(Standard(4, 2)) Then Source(i, 21) = 1
                                If CDate(Standard(4, 2)) < CDate(Source(i, 5)) And CDate(Source(i, 5)) <= CDate(Standard(4, 3)) Then Source(i, 21) = 2
                                If CDate(Standard(4, 3)) < CDate(Source(i, 5)) Then Source(i, 7) = 1: Source(i, 21) = 3
                            End If
                         End If
                         If Len(Source(i, 22)) = 0 Then
                            If Len(Source(i, 6)) = 0 Then
                                Source(i, 9) = Source(i, 9) + 1: Source(i, 6) = "©ǩ": Source(i, 22) = 4
                            Else
                                If CDate(Standard(4, 5)) < CDate(Source(i, 6)) Then Source(i, 22) = 1
                                If CDate(Standard(4, 4)) < CDate(Source(i, 6)) And CDate(Source(i, 6)) <= CDate(Standard(4, 5)) Then Source(i, 22) = 2
                                If CDate(Source(i, 6)) <= CDate(Standard(4, 4)) Then Source(i, 8) = 1: Source(i, 22) = 3
                            End If
                         End If
                    Case Is = 2                                                                                     '��������
                         If Len(Source(i, 21)) = 0 Then
                            If Len(Source(i, 5)) = 0 Then
                                Source(i, 12) = 1: Source(i, 5) = "©ǩ": Source(i, 21) = 4
                            Else
                                If Source(i, 5) <= CDate(Standard(5, 2)) Then Source(i, 21) = 1
                                If CDate(Standard(5, 2)) < CDate(Source(i, 5)) And CDate(Source(i, 5)) <= CDate(Standard(5, 3)) Then Source(i, 21) = 2
                                If CDate(Standard(5, 3)) < CDate(Source(i, 5)) Then Source(i, 10) = 1: Source(i, 21) = 3
                            End If
                         End If
                         If Len(Source(i, 22)) = 0 Then
                             If Len(Source(i, 6)) = 0 Then
'                                Source(i, 12) = Source(i, 12) + 1: Source(i, 6) = "©ǩ": Source(i, 22) = 4
                             Else
                                 If CDate(Standard(5, 5)) < CDate(Source(i, 6)) Then Source(i, 22) = 1
                                 If CDate(Standard(5, 4)) < CDate(Source(i, 6)) And CDate(Source(i, 6)) <= CDate(Standard(5, 5)) Then Source(i, 22) = 2
'                                If  Cdate(Source(i, 6)) <= CDate(Standard(5, 4)) Then Source(i, 11) = 1: Source(i, 22) = 3
                             End If
                         End If
                    Case Is = 3                                                                                     '��������
                         If Len(Source(i, 21)) = 0 Then
                            If Len(Source(i, 5)) = 0 Then
'                                Source(i, 15) = 1: Source(i, 5) = "©ǩ": Source(i, 21) = 4
                            Else
                                If CDate(Source(i, 5)) <= CDate(Standard(6, 2)) Then Source(i, 21) = 1
                                If CDate(Standard(6, 2)) < CDate(Source(i, 5)) And CDate(Source(i, 5)) <= CDate(Standard(6, 3)) Then Source(i, 21) = 2
'                                If CDate(Standard(6, 3)) < CDate(Source(i, 5)) Then Source(i, 13) = 1: Source(i, 21) = 3
                            End If
                         End If
                         If Len(Source(i, 22)) = 0 Then
                            If Len(Source(i, 6)) = 0 Then
                                Source(i, 15) = Source(i, 15) + 1: Source(i, 6) = "©ǩ": Source(i, 22) = 4
                            Else
                                If CDate(Standard(6, 5)) < CDate(Source(i, 6)) Then Source(i, 22) = 1
                                If CDate(Standard(6, 4)) < CDate(Source(i, 6)) And CDate(Source(i, 6)) <= CDate(Standard(6, 5)) Then Source(i, 22) = 2
                                If CDate(Source(i, 6)) <= CDate(Standard(6, 4)) Then Source(i, 14) = 1: Source(i, 22) = 3
                            End If
                         End If
                End Select
                If m = 1 Then                                                                                         '�ָ�ǩ��ǩ��ʱ��
                    If 0 < Source(i, 21) And Source(i, 21) < 4 Then Source(i, 5) = CDate(Source(i, 5)) - CDate(DuiShiTemp(1, 5)) + CDate(DuiShiTemp(1, 6))
                    If 0 < Source(i, 22) And Source(i, 22) < 4 Then Source(i, 6) = CDate(Source(i, 6)) - CDate(DuiShiTemp(1, 7)) + CDate(DuiShiTemp(1, 8))
                End If
                S2RowMax = S2RowMax + 1                                                                               '���ְ��2�������Σ�
                For j = 1 To UBound(Source2, 2)
                    Source2(S2RowMax, j) = Source(i, j)
                Next
        End Select
    Next
' ����Ŀ��Excel�ļ�
'    OutFileFix = "��" & Format(DateMin, "yyyy" & "��" & "m" & "��" & "d" & "��") & "-" & Format(DateMax, "yyyy" & "��" & "m" & "��" & "d" & "��") & "��"
'    OutFolder = OutPath & "\" & Format(DateMax, "m" & "��" & "d" & "��") & "��ʽ�ϱ�"
'    If SFO.FolderExists(OutFolder) = False Then
'       MkDir OutFolder
'    End If
' ֻ����쳣����
    Call GenerateBook(Source1, S1RowMax, 12, NameTeacherUN)
    Call GenerateBook(Source2, S2RowMax, 15, NameHeadMasterUN)
' �رռ��˳�
    Application.DisplayAlerts = False
    Workbooks.Close                                                                                                     '�ر����й�����
    Application.DisplayAlerts = True
    Application.Quit                                                                                                    '�˳�Excel
    Shell "explorer.exe " & OutFolder, vbNormalFocus                                                                    '��Ŀ���ļ���
End Sub
'
Sub GenerateBook(InSource, InRmax, InCmax, InName)
    Dim OutBook As Workbook
    Dim OutSource(1 To RowMax, 1 To SubColMax) As Variant
    Dim i, j, k, m, n, p, q, OutMax As Integer
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
        If InSource(i, 1) = InSource(i + 1, 1) Then
            k = k + 1
        Else
            For j = i - k To i - 1
                For m = 7 To InCmax
                   InSource(i, m) = InSource(i, m) + InSource(j, m)
                   InSource(j, m) = Empty
                Next
            Next
            For m = 7 To InCmax
                   InSource(i, InCmax + 1) = InSource(i, InCmax + 1) + InSource(i, m)
            Next
            If InSource(i, InCmax + 1) > 0 Then
                For j = i - k To i
                    p = p + 1
                    For q = 1 To SubColMax
                        OutSource(p, q) = InSource(j, q)
                    Next
                    OutSource(p, 2) = Format(OutSource(p, 2), DateFormat) & "(" & OutSource(p, 4) & ")"
                    OutSource(p, 4) = OutSource(p, 25)                                                      '�����滻Ϊ��ϰ��Ϣ
                    Select Case OutSource(p, 3)                                                             '�ָ��Ű��Ϊ������
                        Case Is = 1
                            OutSource(p, 3) = "����"
                        Case Is = 2
                            OutSource(p, 3) = "����"
                        Case Is = 3
                            OutSource(p, 3) = "����"
                    End Select
                Next
                For q = 7 To InCmax
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
        Select Case CInt(OutSource(i, 21))
            Case Is = 1
                Call COAColor(Cells(i, 5), 4, 10)                                                           '���̵�+������
            Case Is = 2
                Call COAColor(Cells(i, 5), 6, 10)                                                           '��ɫ��+������
            Case Is = 3
                Call COAColor(Cells(i, 5), 3, 2)                                                            '����+��ɫ��
                Call COAColor(Cells(i, 3), 3, 2)
            Case Is = 4
                Call COAColor(Cells(i, 5), 3, 6)                                                            '����+��ɫ��
                Call COAColor(Cells(i, 3), 3, 6)
            Case Is = 5
                Call COAColor(Cells(i, 5), 10, 6)                                                           '���̵�+��ɫ��
        End Select
        Select Case CInt(OutSource(i, 22))
            Case Is = 1
                Call COAColor(Cells(i, 6), 4, 10)                                                           '���̵�+������
            Case Is = 2
                Call COAColor(Cells(i, 6), 6, 10)                                                           '��ɫ��+������
            Case Is = 3
                Call COAColor(Cells(i, 6), 3, 2)                                                            '����+��ɫ��
                Call COAColor(Cells(i, 3), 3, 2)
            Case Is = 4
                Call COAColor(Cells(i, 6), 3, 6)                                                            '����+��ɫ��
                Call COAColor(Cells(i, 3), 3, 6)
            Case Is = 5
                Call COAColor(Cells(i, 6), 10, 6)                                                           '���̵�+��ɫ��
        End Select
        If Len(OutSource(i, 4)) > 0 Then
            Call COAColor(Cells(i, 4), 37, 51)
        End If
    Next
    k = 0: p = 0
    For i = 2 To OutMax
        If OutSource(i, 1) = OutSource(i + 1, 1) Then
            k = k + 1
        Else
            For m = 7 To InCmax + 1
                If OutSource(i, m) > 0 Then                                                                 'ͳ������ɫ
                    Call COAColor(Cells(i, m), 3, 2)                                                        '����+��ɫ��
                ElseIf OutSource(i, m) = 0 Then
                    Call COAColor(Cells(i, m), 4, 10)                                                       '���̵�+������
                End If
            Next
            Range(Cells(i - k, 1), Cells(i, 1)).Merge                                                       '�ϲ�����
            For q = 7 To InCmax Step 3
                Range(Cells(i - k, q), Cells(i - 2, q + 2)).Merge                                           '�ϲ�ͳ����
            Next
            Cells(i - k, 7) = OutSource(i, 24)                                                              '���������Ϣ
            k = 0
        End If
        If OutSource(i, 2) = OutSource(i + 1, 2) Then
            p = p + 1
        Else
            Range(Cells(i - p, 2), Cells(i, 2)).Merge                                                       '�ϲ�����
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
    GEndDate = DateMin
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
        If IsDate(Holiday(i, 3)) And IsDate(Holiday(i, 4)) Then           '���ж�Ϊ���ں���ִ����ٲ���
            If CDate(GBeginDate) <= CDate(Holiday(i, 4)) And CDate(Holiday(i, 3)) <= CDate(GEndDate) And InStr(Holiday(i, 1), "*") = 0 Then
                k = k + 1
                PreLeave(k, 1) = "�߶�"
'' ȡ����ٵ�ʱ����
                If CDate(Holiday(i, 3)) <= CDate(GBeginDate) Then           'ȡ����ʼʱ��
                    PreLeave(k, 2) = CDate(GBeginDate)
                Else
                    PreLeave(k, 2) = CDate(Holiday(i, 3))
                End If
                If CDate(GEndDate) <= CDate(Holiday(i, 4)) Then             'ȡ����ֹʱ��
                    PreLeave(k, 3) = CDate(GEndDate)
                Else
                    PreLeave(k, 3) = CDate(Holiday(i, 4))
                End If
'' ����Ԥ������ٱ��Ͻ�ѧУ��
                PreLeave(k, 6) = Holiday(i, 1)
                If Holiday(i, 5) > 0 Then
                    PreLeave(k, 7) = Holiday(i, 5)
                    PreLeave(k, 4) = PreLeave(k, 3) - PreLeave(k, 2) + 1
                Else
                    PreLeave(k, 4) = 0.5 * (PreLeave(k, 3) - PreLeave(k, 2) + 1)
                    If Holiday(i, 6) > 0 Or Holiday(i, 7) > 0 Then
                        PreLeave(k, 5) = "����"
                    End If
                    If Holiday(i, 8) > 0 Or Holiday(i, 9) > 0 Then
                        PreLeave(k, 5) = "����"
                    End If
                    If Holiday(i, 10) > 0 Or Holiday(i, 11) > 0 Then
                        PreLeave(k, 5) = "����"
                    End If
                    If Holiday(i, 6) > 0 Then
                        PreLeave(k, 7) = Holiday(i, 6)
                    ElseIf Holiday(i, 7) > 0 Then
                        PreLeave(k, 7) = Holiday(i, 7)
                    ElseIf Holiday(i, 8) > 0 Then
                        PreLeave(k, 7) = Holiday(i, 8)
                    ElseIf Holiday(i, 9) > 0 Then
                        PreLeave(k, 7) = Holiday(i, 9)
                    ElseIf Holiday(i, 10) > 0 Then
                        PreLeave(k, 7) = Holiday(i, 10)
                    ElseIf Holiday(i, 11) > 0 Then
                        PreLeave(k, 7) = Holiday(i, 11)
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
  NameLeave = "�߶�����" & Format(GBeginDate, "m" & "��" & "d" & "��") & "-" & Format(GEndDate, "m" & "��" & "d" & "��") & "����"
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




