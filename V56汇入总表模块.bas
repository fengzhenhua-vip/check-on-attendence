Attribute VB_Name = "V56�����ܱ�ģ��"
' �����ܱ�ģ��V55
' ���ߣ�����
' ���ڣ�2021��4��29��
' ���ã���ÿ��ͳ�ƵĽ�������ܱ�����ʱ���Υ��������������ÿ�ܻ��ܽ����Ҫ�˹�У׼�����ϵ���������һ����

Sub �����ܱ�()
    Call ������������
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
    TotalFolder = OutPath & "\" & Format(Now, "yyyy" & "��") & "ͳ���쳣�ܱ�"
    ToTeName = NameTeacherUN & Format(Now, "yyyy" & "��") & "�ܱ�"
    ToHeName = NameHeadMasterUN & Format(Now, "yyyy" & "��") & "�ܱ�"
    ToTeFile = TotalFolder & "\" & ToTeName & ".xlsx"
    ToHeFile = TotalFolder & "\" & ToHeName & ".xlsx"
    Dim SFO As Object
    Set SFO = CreateObject("Scripting.FileSystemObject")                                                        '��SFOΪ�ļ��ж������

    If SFO.FolderExists(TotalFolder) = False Then
       MkDir TotalFolder
    End If
    If SFO.FileExists(ToTeFile) = False Then
       Call �����쳣�ܱ�(ToTeFile, ToTeName)
    End If
    If SFO.FileExists(ToHeFile) = False Then
       Call �����쳣�ܱ�(ToHeFile, ToHeName)
    End If
' �쳣�����������ToTalSource
    Imax = Range("c65536").End(xlUp).Row
    Jmax = Cells(1, 200).End(xlToLeft).Column
    ToTalSource = Range(Cells(1, 1), Cells(Imax, Jmax))
' �����ܱ��������
    If InStr(ToTalSource(2, 3), NameTeacher) > 0 Then
        Set TBook = GetObject(ToTeFile)
        TeacherSource = TBook.Sheets(ToTeName).Range("a1:bb" & 600)
      ElseIf InStr(ToTalSource(2, 3), NameHeadMaster) > 0 Then
        Set TBook = GetObject(ToHeFile)
        TeacherSource = TBook.Sheets(ToHeName).Range("a1:bb" & 600)
    End If
' ȡ��TeacherSource �����ǿ��к�����
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
 ' ȡ��TotalSource �������
    ToTalDate = ToTalSource(2, 2)
    For i = 3 To UBound(ToTalSource, 1)
       If CDate(ToTalDate) < CDate(ToTalSource(i, 2)) Then
        ToTalDate = ToTalSource(i, 2)
       End If
    Next
' ��ȡTeacherSource ����д���к�WriteColumn
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
'  ��ȡToTalSource ����Ч���ݵ�SubToTalSource
    l = 0
    k = 0                                                               '��¼��SubToTalSource �е���Ч��������
    Dim SubToTalSource(1 To 600, 1 To 2) As Variant                     '���н�ְ�������ᳬ��600�������ݶ�Ϊ600
    If InStr(ToTalSource(2, 3), NameTeacher) > 0 Then
        For j = 2 To UBound(ToTalSource, 1)
           If ToTalSource(j, 13) > 0 Then                               '��ʦ13��
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
        For j = 2 To UBound(ToTalSource, 1)                             '������16��
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
' ��Ŀ�� TeacherSource ������Ч����
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
''����m��¼��TeacherSource �е���Ч������,��һ������������Sort ��������������
    Dim SortC(1 To 600, 1 To 1) As Variant
    Dim SortCMin As Variant
    k = 4                                                                   '��С�����к�
'����Ѿ����������������к�n
    If WriteColumn < TBCmax Then
        n = TBCmax
    Else
        n = WriteColumn
    End If
' ������,���Ƚ�����ʱӦ����ǿ������ת��Ϊ���������Ƚ�
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
' ���Ѿ���������������
    For i = 2 To m                                                           'm��¼��TeacherSource�е�����������
       TeacherSource(i, 3) = 0
       For j = 4 To n
            TeacherSource(i, 3) = TeacherSource(i, 3) + TeacherSource(i, j)
       Next
    Next
' ���ݵ�3��������������
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
' ��1��׷������
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
' ����д�뵽Ŀ���ļ�
     If InStr(ToTalSource(2, 3), NameTeacher) > 0 Then
        Workbooks.Open Filename:=ToTeFile
     Else
        Workbooks.Open Filename:=ToHeFile
     End If
     Range("a1:bb" & 600) = TeacherSource
     Call ��ʽ��
     Range(Cells(1, 1), Cells(m, n)).Select
     Call FontSet(NameFont)
     Cells(1, 1).Select
     If InStr(ToTalSource(2, 3), NameTeacher) > 0 Then
        Workbooks(ToTeName).Close savechanges:=True
    Else
        Workbooks(ToHeName).Close savechanges:=True
    End If
End Sub
Sub �����쳣�ܱ�(InFile, InName)
    Dim ToTalBook As Workbook
    Application.SheetsInNewWorkbook = 1                                             ' ����1��Sheet
    Set ToTalBook = Workbooks.Add
    Application.DisplayAlerts = False
        ToTalBook.SaveAs Filename:=InFile
        Sheets(1).name = InName
        Sheets(InName).Range("C2:BB600").NumberFormatLocal = "0;[��ɫ]0"            '����ʱ���ʽ
        Sheets(InName).Rows("1:1").NumberFormatLocal = DateFormat                   '�������ڸ�ʽ
        Cells(1, 1) = "����"
        Cells(1, 2) = "����"
        Cells(1, 3) = "����"
        ActiveWorkbook.Close savechanges:=True
    Application.DisplayAlerts = True
    Set ToTalBook = Nothing                                                         'ȡ��ToTalBook
End Sub

