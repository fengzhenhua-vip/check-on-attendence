Attribute VB_Name = "AddToTotalLeave_V71"
' ��Ŀ��AddToTotalLeave
' �汾��V71
' ���ߣ�����
' ���ڣ�2021��7��16��
' ���ã���ÿ��ͳ�ƵĽ�������ܱ�����ʱ���Υ��������������ÿ�ܻ��ܽ����Ҫ�˹�У׼�����ϵ���������һ����
' ��־��2021��7��13��-15��ʵ���˵�7������ϵͳ����ȫ�Ӷ���������¹��������Ż�����ʾ��ʽ�������������£�
'       1.����COAV70�ĸ�ʽҪ��
'       2.�Ż�AddToToalLeaveģ�飬�淶���룬����Ч��
'       3.�����汾��ΪV70����COASMain��ͬ���������ƥ���ʶ��
' ��־���޸�����bug,ֻ�����쳣��ʦ���쳣�����α����ʱ��ִ�л����ܱ�ͬʱ�Ż��˲��ִ��룬�����Եõ�������������ɫ��ʾ����������汾��V71
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
        TotalFolder = OutPath & "\" & Format(Now, "yyyy" & "��") & "ͳ���ܱ�"
        ToTeName = NameTeacherUN & Format(Now, "yyyy" & "��") & "�ܱ�"
        ToHeName = NameHeadMasterUN & Format(Now, "yyyy" & "��") & "�ܱ�"
        ToTeFile = TotalFolder & "\" & ToTeName & ".xlsx"
        ToHeFile = TotalFolder & "\" & ToHeName & ".xlsx"
        Set SFO = CreateObject("Scripting.FileSystemObject")                                                                             '��SFOΪ�ļ��ж������
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
' �쳣�����������ToTalSource
        Imax = Cells(RowMax, 3).End(xlUp).Row
        Jmax = Cells(1, ColMax).End(xlToLeft).Column
        ToTalSource = Range(Cells(1, 1), Cells(Imax, Jmax))
' �����ܱ��������
        Set TBook = GetObject(TotalFile)
        TeacherSource = TBook.Sheets(TotalName).Range(TBook.Sheets(TotalName).Cells(1, 1), TBook.Sheets(TotalName).Cells(RowMax, ColMax))
' ȡ��TeacherSource �����ǿ��к�����
        TBRmax = TBook.Sheets(TotalName).Cells(RowMax, 3).End(xlUp).Row
        TBCmax = TBook.Sheets(TotalName).Cells(1, ColMax).End(xlToLeft).Column
' ȡ��TotalSource �������
        ToTalDate = Mid(ToTalSource(2, 2), 1, Len(ToTalSource(2, 2)) - 3)
        For i = 4 To Imax
           If Len(ToTalSource(i, 2)) > 3 Then
            If CDate(ToTalDate) < CDate(Mid(ToTalSource(i, 2), 1, Len(ToTalSource(i, 2)) - 3)) Then
             ToTalDate = Mid(ToTalSource(i, 2), 1, Len(ToTalSource(i, 2)) - 3)
            End If
           End If
        Next
' ��ȡTeacherSource ����д���к�WriteColumn
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
'  ��ȡToTalSource ����Ч���ݵ�SubToTalSource
        k = 0                                                                                               '��¼��SubToTalSource �е���Ч��������
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
        Dim SortC(1 To RowMax, 1 To 1) As Variant
        Dim SortCMin As Variant
        k = 4                                                                                                    '��С�����к�
'����Ѿ����������������к�n
        If WriteColumn < TBCmax Then
            n = TBCmax
        Else
            n = WriteColumn
        End If
'������,���Ƚ�����ʱӦ����ǿ������ת��Ϊ���������Ƚ�
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
'���Ѿ���������������
        For i = 2 To m                                                                                           'm=TeacherSource����������
           TeacherSource(i, 3) = 0
           For j = 4 To n
                TeacherSource(i, 3) = TeacherSource(i, 3) + TeacherSource(i, j)
           Next
        Next
'���ݵ�3��������������
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
    Application.SheetsInNewWorkbook = 1                                                                 ' ����1��Sheet
    Set ToTalBook = Workbooks.Add
    Application.DisplayAlerts = False
        ToTalBook.SaveAs Filename:=InFile
        Cells(1, 1) = "����"
        Cells(1, 2) = "����"
        Cells(1, 3) = "����"
        Sheets(1).Name = InName
        ActiveWorkbook.Close savechanges:=True
    Application.DisplayAlerts = True
    Set ToTalBook = Nothing                                                                              'ȡ��ToTalBook
End Sub
