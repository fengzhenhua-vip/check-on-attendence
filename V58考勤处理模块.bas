Attribute VB_Name = "V58���ڴ���ģ��"
' ͨ�ÿ���ϵͳV58
' ���ߣ�����
' ���ڣ�2021��4��25��--2021��5��13��
' ���ã������׼��������Դ����Original
' ˵������У׼���ɵ������������ļ�ֱ������ΪCorrect ,�������ĺô����ڿ��Ը�����ϰ�仯��ʱ����У׼����������ɺ���õ�ʱ�伸�����Ժ��Բ��ƣ��ͷ�������������������������
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
'
'����ȫ�ֱ���
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
    Public Morning As Integer                                              '�����ϰ�ʱ��Υ�����
    Public Afternoon As Integer
    Public Evening As Integer
    Public MorningX As Integer                                             '��ĩʱ��Υ�����
    Public AfternoonX As Integer
    Public EveningX As Integer
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
    ConfigFile = ConfigFolder & "\" & "��������.xlsx"
    ConfigSheet1 = "ֹͣ����"
    ConfigSheet2 = "��ٱ�"
    OriginalSheet1 = "��ϰ����"                                                                                 '����4��רΪУ׼������
    OriginalSheet3 = "����У׼"
    OriginalSheet4 = "���α�"
    WuYi = Format(Now, "yyyy") & "/5/1"
    ShiYi = Format(Now, "yyyy") & "/10/1"
    If CDate(WuYi) < CDate(Now) < CDate(ShiYi) Then
        ConfigSheet3 = "��һ���׼ʱ��"
        OriginalSheet2 = "��һ��У׼ʱ��"
    Else
        ConfigSheet3 = "ʮһ���׼ʱ��"
        OriginalSheet2 = "ʮһ��У׼ʱ��"
    End If
    VipSwitch = 1                                                                                               '����vip
    NormalSwitch = 1                                                                                            '0�������������1���
    StopSymbol = "*"
End Sub
Sub ��׼ͨ�ÿ���(Original)
    Dim i, j, k, l, m, n, o, p As Integer
    Dim Source As Variant
    Dim Holiday As Variant
    Dim VipSource As Variant
    Dim Teacher() As Variant
    Dim HeadMaster() As Variant
' ����У׼��ʱ����������
    Dim SelfStudyTable As Variant
    Dim CorrectTable As Variant
    Dim ReCorrectTable As Variant
    Dim CorrectTime As Variant
    Dim SRmax, SCmax, HRmax, HCmax, CRmax, CCmax, STRmax, STCmax, ViRmax, ViCmax, ORmax, OCmax As Integer       'ͨ����ֵ
    Dim SSRmax, SSCmax, RCTRmax, RCTCmax, CTRmax, CTCmax, CTERmax, CTECmax As Integer                           'רΪУ׼������
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
''                                                                                                          '����4��רΪУ׼������
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
    CGRmax = ConfigBook.Sheets(OriginalSheet4).Range("a65536").End(xlUp).Row                                '���뻻�α�
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
                If InStr(CorrectTable(i, 2), ReCorrectTable(j, 2)) > 0 Then
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
''CorrectTable ������α�Change
l = p                                                                '��õ�ǰCorrectTable�е�������l
For i = 2 To CGRmax
' ����У׼
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
    CorrectTable(p, 9) = "��"
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
' �������BC�����粻����������20���Ӵ���
    If InStr(Change(i, 5), "B") + InStr(Change(i, 5), "C") > 0 Then
        p = p + 1
        For j = 1 To 4
            CorrectTable(p, j) = CorrectTable(p - 1, j)
        Next
        CorrectTable(p, 3) = "��ʦ����"
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
            If InStr(CorrectTable(m, 9), "��9��") > 0 Then
                CorrectTable(m, 9) = "��9��"
            End If
        End If
    End If
' ����У׼
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
    CorrectTable(p, 9) = "��"
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
' �������BC����������20����������
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
            If InStr(CorrectTable(m, 9), "��") > 0 Then
            Else
                p = p + 1
                For j = 1 To 4
                    CorrectTable(p, j) = CorrectTable(p - 1, j)
                Next
                CorrectTable(p, 3) = "��ʦ����"
                CorrectTable(p, 6) = CorrectTime(2, 2)
                CorrectTable(p, 9) = "��" & CorrectTable(m, 9)
            End If
        Else
            p = p + 1
            For j = 1 To 4
                CorrectTable(p, j) = CorrectTable(p - 1, j)
            Next
            CorrectTable(p, 3) = "��ʦ����"
            CorrectTable(p, 6) = CorrectTime(2, 2)
            CorrectTable(p, 9) = "��"
        End If
    End If
Next
'����Changed �����ںϲ���Ԫ���м�¼�������
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
          Changed(k, 5) = Changed(k, 5) & " ��1��"
        End If
        If InStr(Change(i, 5), "C") > 0 Then
           Changed(k, 5) = Changed(k, 5) & " ��5��"
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
  If InStr(Change(i, 10), "B") + InStr(Change(i, 10), "C") + InStr(Change(i, 10), "D") > 0 Then
        If InStr(Change(i, 10), "B") > 0 Then
          Changed(k, 10) = Changed(k, 10) & " ��1��"
        End If
        If InStr(Change(i, 10), "C") > 0 Then
           Changed(k, 10) = Changed(k, 10) & " ��5��"
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
Next
' ͬһ���˵ĵ����������һ�飬ֻ����һ������
For i = 2 To 2 * CGRmax
    For j = 2 To 2 * CGRmax
        If i > j Then
            If Changed(i, 1) = Changed(j, 1) Then
                Changed(j, 1) = 0
            End If
        End If
    Next
Next
'' ���Correct(У׼��)
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
    For i = 1 To ORmax
        l = 0
        If Original(i, 3) <> 0 Then
' ����i�к�VipSource �б�ȶԣ���l��¼�ȶ���������ȶ���Ϊ1���������ƥ�������
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
'У׼Source�е�ǩ��ǩ������
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
'��У׼���Source����ͳ������
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
' �����˰���������ǩ�ˣ�����ɾ��
'        If IsNumeric(Source(i, 6)) Then
'           If Source(i, 6) = 0 Then
'              Source(i, 12) = Source(i, 12) + 1
'           ElseIf CDate(Source(i, 6)) < CDate(Standard(5, 4)) Then
'              Source(i, 11) = 1
'           End If
'        End If
     ElseIf InStr(Source(i, 3), Standard(6, 1)) > 0 Then
'�����˰���������ǩ��������ɾ��
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
' ��ȡͳ��ʱ�䷶Χ
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
'  ��������ļ���
    OutFileFix = "��" & Format(BeginDate, "yyyy" & "��" & "m" & "��" & "d" & "��") & "-" & Format(EndDate, "yyyy" & "��" & "m" & "��" & "d" & "��") & "��"
    OutFolder = OutPath & "\" & Format(EndDate, "m" & "��" & "d" & "��") & "��ʽ�ϱ�"
    If SFO.FolderExists(OutFolder) = False Then
       MkDir OutFolder
    End If
' ����Teacher
    Call ����ͳ�ƴ���(Teacher)
' ����HeadMaster
    Call ����ͳ�ƴ���(HeadMaster)
    Application.DisplayAlerts = False
    Workbooks.Close                                                 '�ر����й�����
    Application.DisplayAlerts = True
    Application.Quit                                                '�˳�Excel
    Shell "explorer.exe " & OutFolder, vbNormalFocus                '��Ŀ���ļ���
End Sub
Sub ����ͳ�ƴ���(THDATA)
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
' �����쳣�����ο��ڵ��Թ��� adding,Morning �ȼ�¼�����ε���ʦ�����ϰ�Ĵ������ǿ���ѡ����ȥ���Ĵ���2021/5/13
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
                If InStr(Abnormal(l, 4), "��") + InStr(Abnormal(l, 4), "��") > 0 Then
                    If InStr(Abnormal(l, 3), "����") > 0 Then
                       If IsNumeric(Abnormal(l, 5)) Then
                         If CDate(Standard(4, 3)) <= CDate(Abnormal(l, 5)) Then
                            If CDate(Abnormal(l, 5)) < CDate(Standard(2, 3)) Then
                                     MorningX = MorningX + 1
                            End If
                         End If
                       End If
                    End If
                    If InStr(Abnormal(l, 3), "����") > 0 Then
                        If IsNumeric(Abnormal(l, 5)) Then
                            If CDate(Standard(5, 3)) < CDate(Abnormal(l, 5)) Then
                                If CDate(Abnormal(l, 5)) < CDate(Standard(3, 3)) Then
                                    AfternoonX = AfternoonX + 1
                                End If
                            End If
                        End If
                    End If
                    If InStr(Abnormal(l, 3), "����") > 0 Then
                        If IsNumeric(Abnormal(l, 6)) Then
                            If CDate(Abnormal(l, 6)) < CDate(Standard(6, 4)) Then
                                EveningX = EveningX + 1
                            End If
                        End If
                    End If
                Else
                    If InStr(Abnormal(l, 3), "����") > 0 Then
                       If IsNumeric(Abnormal(l, 5)) Then
                         If CDate(Standard(4, 3)) <= CDate(Abnormal(l, 5)) Then
                            If CDate(Abnormal(l, 5)) < CDate(Standard(2, 3)) Then
                                     Morning = Morning + 1
                            End If
                         End If
                       End If
                    End If
                    If InStr(Abnormal(l, 3), "����") > 0 Then
                        If IsNumeric(Abnormal(l, 5)) Then
                            If CDate(Standard(5, 3)) < CDate(Abnormal(l, 5)) Then
                                If CDate(Abnormal(l, 5)) < CDate(Standard(3, 3)) Then
                                    Afternoon = Afternoon + 1
                                End If
                            End If
                        End If
                    End If
                    If InStr(Abnormal(l, 3), "����") > 0 Then
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
'''' ��ʦ����
    If InStr(THColor(i, 3), "��ʦ����") > 0 Then
'''''ǩ��
      Cells(i, 5).Select
      If IsNumeric(THColor(i, 5)) Then
        If THColor(i, 5) < Standard(2, 2) Then
''''''''''''''
         If THColor(i, 5) = 0 Then
            Call ©ǩɫ
         Else
            Call ����ɫ
         End If
''''''''''''''
        ElseIf THColor(i, 5) < Standard(2, 3) Then
         Call Ԥ��ɫ
        Else
         Call Υ��ɫ
        End If
      Else
        Call ��עɫ
      End If
'''''ǩ��
      Cells(i, 6).Select
      If IsNumeric(THColor(i, 6)) Then
       If THColor(i, 6) < Standard(2, 4) Then
         If THColor(i, 6) = 0 Then
          Call ©ǩɫ
         Else
          Call Υ��ɫ
         End If
       ElseIf THColor(i, 6) < Standard(2, 5) Then
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
      If IsNumeric(THColor(i, 5)) Then
       If THColor(i, 5) < Standard(3, 2) Then
''''''''''''''
        If THColor(i, 5) = 0 Then
         Call ©ǩɫ
        Else
         Call ����ɫ
        End If
''''''''''''''
       ElseIf THColor(i, 5) < Standard(3, 3) Then
        Call Ԥ��ɫ
       Else
        Call Υ��ɫ
       End If
     Else
       Call ��עɫ
     End If
'''''ǩ��
      Cells(i, 6).Select
      If IsNumeric(THColor(i, 6)) Then
       If THColor(i, 6) < Standard(3, 4) Then
        If THColor(i, 6) = 0 Then
            Call ©ǩɫ
        Else
            Call Υ��ɫ
        End If
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
      If IsNumeric(THColor(i, 5)) Then
       If THColor(i, 5) < Standard(4, 2) Then
''''''''''''''
        If THColor(i, 5) = 0 Then
         Call ©ǩɫ
        Else
         Call ����ɫ
        End If
''''''''''''''
       ElseIf THColor(i, 5) < Standard(4, 3) Then
        Call Ԥ��ɫ
       Else
        Call Υ��ɫ
       End If
      Else
       Call ��עɫ
      End If
'''''ǩ��
      Cells(i, 6).Select
      If IsNumeric(THColor(i, 6)) Then
       If THColor(i, 6) < Standard(4, 4) Then
        If THColor(i, 6) = 0 Then
         Call ©ǩɫ
        Else
         Call Υ��ɫ
        End If
       Else
         Call ����ɫ
       End If
      Else
       Call ��עɫ
      End If
''''����������ֻͳ��ǩ��
    ElseIf InStr(THColor(i, 3), "����������") > 0 Then
      Cells(i, 5).Select
      If IsNumeric(THColor(i, 5)) Then
       If THColor(i, 5) < Standard(5, 2) Then
''''''''''''''
        If THColor(i, 5) = 0 Then
         Call ©ǩɫ
        Else
         Call ����ɫ
        End If
''''''''''''''
       ElseIf THColor(i, 5) < Standard(5, 3) Then
        Call Ԥ��ɫ
       Else
        Call Υ��ɫ
       End If
      Else
       Call ��עɫ
      End If
''''����������ֻͳ��ǩ��
    ElseIf InStr(THColor(i, 3), "����������") > 0 Then
      Cells(i, 6).Select
      If IsNumeric(THColor(i, 6)) Then
       If THColor(i, 6) < Standard(6, 4) Then
        If THColor(i, 6) = 0 Then
            Call ©ǩɫ
        Else
            Call Υ��ɫ
        End If
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
' ����У׼��ʱ��ָ�
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
                RecSource(i, 5) = ""
            End If
        End If
        If RecSource(i, 6) = 0 Then
            If InStr(RecSource(i, 3), "����������") > 0 Then
            Else
                RecSource(i, 6) = "©ǩ"
            End If
        Else
            If InStr(RecSource(i, 3), "����������") > 0 Then
                RecSource(i, 6) = ""
            End If
        End If
    Next
' ��������д��Sheet
    Range(Cells(1, 1), Cells(RECR, RECC)) = RecSource
' �ϲ�ͳ������������
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
            Call �ϲ�ѡ�е�Ԫ��
            If m > 0 Then
                Cells(j, 7) = "����:" & Changed(m, 1) & Chr(10) & Changed(m, 5) & Chr(10) & "����:" & Changed(m, 6) & Chr(10) & Changed(m, 10)
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
    ActiveWorkbook.Close savechanges:=True
    Application.DisplayAlerts = True
    Set OutBook = Nothing                                                                    'ȡ��OutBook
End Sub
