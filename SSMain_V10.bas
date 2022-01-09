Attribute VB_Name = "SSMain_V10"
' ��Ŀ�� SSMain
' �汾�� V10
' ���ߣ�����
' ��λ��ɽ��ʡƽԭ�ص�һ��ѧ
' ���䣺fengzhenhua@outlook.com
' ���ͣ�https://fengzhenhua-vip.github.io
' ��ҳ��https://github.com/fengzhenhua-vip
' ��Ȩ��2021��12��30��--2022��1��1��
' ��־�� ��ɵ�һ�棬ʵ��ÿ�θ����������Զ�ͳ�ƹ���2022/01/01��ͳ��ʱ�򿪱��ο��Եĳɼ��������ҳ�棬�����ǲ�������
'
Public SSLimOne, SSLimTwo As Variant
Sub SSMain_V10()
    Dim SSCfgPath, SSCfgFile As String
    Dim SSBook As Workbook
    Dim SSImax, SSJmax As Integer
    Dim SSLimName1, SSLimName2 As String
    Dim OutFolder As String
    Dim SSlimTemp As Variant
    OutFolder = ActiveWorkbook.Path & "\ͳ�ƽ��"
    SSLimName1 = "һ��": SSLimName2 = "����"
    SSCfgPath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\����ϵͳ" & "\" & "����ϵͳ����"
    SSCfgFile = SSCfgPath & "\�Աȱ�ģ��.xlsx"
    Set SFO = CreateObject("Scripting.FileSystemObject")
' ��ȡ�����ļ��е�һ�ߺͶ��ߵ���Ӧ���飬ע��һ������ߵ�ѧУ��Ŀ����ȫһ���ģ����Բ��طֿ���д
    Set SSBook = GetObject(SSCfgFile)
    SSImax = SSBook.Sheets(SSLimName1).Cells(10000, 1).End(xlUp).Row
    SSJmax = SSBook.Sheets(SSLimName1).Cells(2, 1000).End(xlToLeft).Column
    SSLimOne = SSBook.Sheets(SSLimName1).Range(SSBook.Sheets(SSLimName1).Cells(1, 1), SSBook.Sheets(SSLimName1).Cells(SSImax, SSJmax))
    SSLimTwo = SSBook.Sheets(SSLimName2).Range(SSBook.Sheets(SSLimName2).Cells(1, 1), SSBook.Sheets(SSLimName2).Cells(SSImax, SSJmax))
' �ڵ�ǰĿ¼�´����ļ��к�Ŀ¼
    If SFO.folderExists(OutFolder) = False Then
        MkDir OutFolder
    End If
' �ڵ�ǰ�ļ��󴴽�Sheet
'    Call SSSRename("Դ��")
    ActiveSheet.Name = "Դ��"
    Call SSSheetAdd("�ܷ�")
    Call SSSheetAdd("����")
    Call SSSheetAdd("��ѧ")
    Call SSSheetAdd("Ӣ��")
    Call SSSheetAdd("����")
    Call SSSheetAdd("��ѧ")
    Call SSSheetAdd("����")
    Call SSSheetAdd("����")
    Call SSSheetAdd("��ʷ")
    Call SSSheetAdd("����")
' ȡ�òμ������ĸ���ѧУ
    Call GetSchool
' д���׼����
    Dim sht As Worksheet
    Dim shtImax, shtJmax As Integer
    SSlimTemp = SSLimOne
    m = 8: n = 1
SSStart:
    For Each sht In ActiveWorkbook.Sheets
        If InStr(sht.Name, "Դ��") = 0 Then
            For k = 3 To UBound(SSlimTemp, 2)
                If InStr(SSlimTemp(2, k), sht.Name) > 0 Then
                    shtImax = sht.Cells(90000, 1).End(xlUp).Row
                    sht.Cells(shtImax, m) = Empty
                    sht.Cells(3, m) = SSlimTemp(2, k)
                    For i = 4 To shtImax
                       For j = 3 To UBound(SSlimTemp, 1) - 1
                        If InStr(SSlimTemp(j, 1), sht.Cells(i, 1)) > 0 Then
                            sht.Cells(i, m) = SSlimTemp(j, k)
                            sht.Cells(shtImax, m) = sht.Cells(shtImax, m) + SSlimTemp(j, k)
                        End If
                       Next
                    Next
                End If
            Next
        End If
    Next
    If n = 1 Then
        SSlimTemp = SSLimTwo
        m = 3: n = 2
        GoTo SSStart:
    End If
'
    Call ͳ�Ʒ�����
'
    Dim OutOk As Integer
    Application.DisplayAlerts = False
    For Each sht In ActiveWorkbook.Sheets
        Select Case sht.Name
            Case Is = "�ܷ�"
                OutOk = 1
            Case Is = "����"
                OutOk = 1
            Case Is = "��ѧ"
                OutOk = 1
            Case Is = "Ӣ��"
                OutOk = 1
            Case Is = "����"
                OutOk = 1
            Case Is = "��ѧ"
                OutOk = 1
            Case Is = "����"
                OutOk = 1
            Case Is = "����"
                OutOk = 1
            Case Is = "��ʷ"
                OutOk = 1
            Case Is = "����"
                OutOk = 1
            Case Else
                OutOk = 0
        End Select
        If OutOk = 1 Then
            sht.Select
            Call ��ɫ
            Call �ӱ߿�
            sht.Copy
            ActiveWorkbook.SaveAs Filename:=OutFolder & "\" & sht.Name & "�������߶Աȱ�", FileFormat:=xlNormal  '�����������ΪEXCELĬ�ϸ�ʽ
            ActiveWorkbook.Close
        End If
    Next
    Application.DisplayAlerts = True
        MsgBox "�����ɼ�ͳ�����!"
End Sub
Sub SSSRename(ReName)
    Dim sht As Worksheet
    Dim shtOk As Integer
    shtOk = 0
    For Each sht In ActiveWorkbook.Sheets
        If InStr(sht.Name, AddName) > 0 Then
            shtOk = 1
        End If
    Next
    If shtOk = 0 Then
        ActiveSheet.Name = AddName
    End If
End Sub

Sub SSSheetAdd(AddName)
    Dim sht As Worksheet
    Dim shtOk As Integer
    shtOk = 0
    For Each sht In ActiveWorkbook.Sheets
        If InStr(sht.Name, AddName) > 0 Then
            shtOk = 1
        End If
    Next
    If shtOk = 0 Then
        Sheets.Add After:=ActiveSheet
        ActiveSheet.Name = AddName
    End If
    Sheets(AddName).Cells(1, 1) = AddName & "�������߶Աȱ�"
    Sheets(AddName).Cells(2, 3) = "���ж���"
    Sheets(AddName).Cells(3, 4) = "����"
    Sheets(AddName).Cells(2, 5) = "��������"
    Sheets(AddName).Cells(3, 6) = "����"
    Sheets(AddName).Cells(3, 7) = "���߲�"
    Sheets(AddName).Cells(2, 8) = "����һ��"
    Sheets(AddName).Cells(3, 9) = "����"
    Sheets(AddName).Cells(2, 10) = "����һ��"
    Sheets(AddName).Cells(3, 11) = "����"
    Sheets(AddName).Cells(3, 12) = "һ�߲�"
    Sheets(AddName).Cells(3, 1) = "ѧУ"
    Sheets(AddName).Cells(3, 2) = "����"
End Sub

Sub AbsorbSchool(AbSht)
    Dim i, j, k, p, q, Imax, Jmax As Integer
    Dim YBook As Variant
    Imax = Sheets("Դ��").Cells(90000, 1).End(xlUp).Row
    Jmax = Sheets("Դ��").Cells(2, 1000).End(xlToLeft).Column
    YBook = Sheets("Դ��").Range(Sheets("Դ��").Cells(1, 1), Sheets("Դ��").Cells(Imax, Jmax))
    Sheets(AbSht).Cells(4, 1) = Sheets("Դ��").Cells(3, 2)
    k = 4
AbsorbBegin:
    For i = 3 To Imax
        If InStr(YBook(i, 2), Sheets(AbSht).Cells(k, 1)) > 0 And Len(YBook(i, 2)) > 0 Then
            YBook(i, 2) = Empty
        End If
    Next
    q = 3
    Do Until Len(YBook(q, 2)) > 0 Or q = Imax
        q = q + 1
    Loop
    If q > 3 And q < Imax Then
        k = k + 1
        Sheets(AbSht).Cells(k, 1) = YBook(q, 2)
        GoTo AbsorbBegin:
    End If
    k = k + 1
    Sheets(AbSht).Cells(k, 1) = "�ϼ�"
End Sub
Sub GetSchool()
    Dim sht As Worksheet
    Call AbsorbSchool("�ܷ�")
    For Each sht In ActiveWorkbook.Sheets
        If InStr(sht.Name, "Դ��") = 0 And InStr(sht.Name, "�ܷ�") = 0 Then
            For i = 4 To Sheets("�ܷ�").Cells(90000, 1).End(xlUp).Row
                sht.Cells(i, 1) = Sheets("�ܷ�").Cells(i, 1)
            Next
        End If
    Next
End Sub
Sub ����X(Xian, Col)
Dim Imax, Jmax As Integer
Imax = Sheets(Xian).Cells(90000, 1).End(xlUp).Row
Jmax = Sheets(Xian).Cells(1, 200).End(xlToLeft).Column
Dim i, j, k, p, q As Integer
Dim outab(1 To 30, 1 To 2) As Variant
Dim intab As Variant
intab = Sheets(Xian).Range(Sheets(Xian).Cells(4, 1), Sheets(Xian).Cells(Imax, 22))
PaiMingStart:
p = p + 1
For i = 1 To Imax - 4
    If CInt(intab(i, Col)) > CInt(outab(p, 2)) And intab(i, Col) > 0 Then
        outab(p, 1) = intab(i, 1)
        outab(p, 2) = intab(i, Col)
    End If
Next
For i = 1 To Imax - 4
    If outab(p, 1) = intab(i, 1) Then
        intab(i, Col) = 0
    End If
Next
If p < Imax - 4 Then
    GoTo PaiMingStart:
End If
For i = 4 To Imax - 1
    For j = 1 To Imax - 4
        If Sheets(Xian).Cells(i, 1) = outab(j, 1) Then
            Sheets(Xian).Cells(i, Col + 1) = j
        End If
    Next
Next
End Sub

Sub ����()
    For r = 3 To 21 Step 2
     Call ����X("һ��", r)
     Call ����X("����", r)
    Next
End Sub

Sub �ֿ�����X(SortName, SortHang, SortLie)
    Dim Imax, Jmax As Integer
    Sheets(SortName).Select
    Imax = Sheets(SortName).Cells(90000, 1).End(xlUp).Row
    Jmax = Sheets(SortName).Cells(2, 200).End(xlToLeft).Column
    ActiveWorkbook.Worksheets(SortName).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(SortName).Sort.SortFields.Add2 Key:=Range(Cells(SortHang, SortLie), Cells(Imax, SortLie)) _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(SortName).Sort
        .SetRange Range(Cells(SortHang, 1), Cells(Imax, Jmax))
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub �ֿ�����XX(FKSource, FKName)
' �������Ը��ֺ�ĳɼ�ͳ��
    Dim FKLie
    Select Case FKName
        Case Is = "�ܷ�"
            FKLie = 6
        Case Is = "����"
            FKLie = 10
        Case Is = "��ѧ"
            FKLie = 14
        Case Is = "Ӣ��"
            FKLie = 18
        Case Is = "����"
            FKLie = 23
        Case Is = "��ѧ"
            FKLie = 28
        Case Is = "����"
            FKLie = 33
        Case Is = "����"
            FKLie = 38
        Case Is = "��ʷ"
            FKLie = 43
        Case Is = "����"
            FKLie = 48
    End Select
    Call �ֿ�����X(FKSource, 3, FKLie)
    Call ͳ�Ʒ�����XX(FKSource, FKLie, FKName)
End Sub
Sub ͳ�Ʒ�����()
    Call �ֿ�����XX("Դ��", "�ܷ�")
    Call �ֿ�����XX("Դ��", "����")
    Call �ֿ�����XX("Դ��", "��ѧ")
    Call �ֿ�����XX("Դ��", "Ӣ��")
    Call �ֿ�����XX("Դ��", "����")
    Call �ֿ�����XX("Դ��", "��ѧ")
    Call �ֿ�����XX("Դ��", "����")
    Call �ֿ�����XX("Դ��", "����")
    Call �ֿ�����XX("Դ��", "��ʷ")
    Call �ֿ�����XX("Դ��", "����")
End Sub

Sub ͳ�Ʒ�����X(Yuan, YuanLie, MuBiao, MBLie)
Dim CurScore As Single
Dim CurRow, CurUp, CurDown As Integer
Dim Imax, Jmax As Integer
Dim CurImax, CurJmax As Integer
Dim YuanDoc, MuBiaoDoc As Variant
Imax = Sheets(Yuan).Cells(90000, 1).End(xlUp).Row
Jmax = Sheets(Yuan).Cells(2, 200).End(xlToLeft).Column
YuanDoc = Sheets(Yuan).Range(Sheets(Yuan).Cells(1, 1), Sheets(Yuan).Cells(Imax, Jmax))
CurImax = Sheets(MuBiao).Cells(90000, 1).End(xlUp).Row
CurJmax = Sheets(MuBiao).Cells(3, 200).End(xlToLeft).Column
MuBiaoDoc = Sheets(MuBiao).Range(Sheets(MuBiao).Cells(1, 1), Sheets(MuBiao).Cells(Imax, Jmax))
CurRow = Sheets(MuBiao).Cells(13, MBLie) + 2
CurScore = Sheets(Yuan).Cells(CurRow, YuanLie)
CurUp = 0: CurDown = 0
Do While YuanDoc(CurRow - CurUp, YuanLie) = YuanDoc(CurRow, YuanLie)
    CurUp = CurUp + 1
Loop
Do While YuanDoc(CurRow + CurDown, YuanLie) = YuanDoc(CurRow, YuanLie)
    CurDown = CurDown + 1
Loop
If CurUp < CurDown Then
    CurScore = YuanDoc(CurRow - CurUp, YuanLie)
Else
    CurScore = YuanDoc(CurRow, YuanLie)
End If
    Sheets(MuBiao).Cells(3, MBLie + 2) = Sheets(MuBiao).Name & "��" & CurScore & ")"
For i = 4 To CurImax - 1
    MuBiaoDoc(i, MBLie + 2) = 0
    MuBiaoDoc(i, 2) = 0
    For j = 3 To Imax
        If InStr(YuanDoc(j, 2), MuBiaoDoc(i, 1)) > 0 And YuanDoc(j, YuanLie) >= CurScore Then
           MuBiaoDoc(i, MBLie + 2) = MuBiaoDoc(i, MBLie + 2) + 1
        End If
        If InStr(YuanDoc(j, 2), MuBiaoDoc(i, 1)) > 0 Then
            MuBiaoDoc(i, 2) = MuBiaoDoc(i, 2) + 1
        End If
    Next
    Sheets(MuBiao).Cells(i, MBLie + 2) = MuBiaoDoc(i, MBLie + 2)
    Sheets(MuBiao).Cells(i, 2) = MuBiaoDoc(i, 2)
Next
For i = 4 To CurImax - 1
 Sheets(MuBiao).Cells(CurImax, 2) = Sheets(MuBiao).Cells(CurImax, 2) + MuBiaoDoc(i, 2)
 Sheets(MuBiao).Cells(CurImax, MBLie + 2) = Sheets(MuBiao).Cells(CurImax, MBLie + 2) + MuBiaoDoc(i, MBLie + 2)
 Sheets(MuBiao).Cells(i, MBLie + 4) = MuBiaoDoc(i, MBLie + 2) - MuBiaoDoc(i, MBLie)
Next
 Sheets(MuBiao).Cells(CurImax, MBLie + 4) = Sheets(MuBiao).Cells(CurImax, MBLie + 2) - MuBiaoDoc(CurImax, MBLie)
End Sub
Sub ͳ�Ʒ�����XX(Source, SouLie, KeMu)
   Call ͳ�Ʒ�����X(Source, SouLie, KeMu, 3)
   Call ͳ�Ʒ�����X(Source, SouLie, KeMu, 8)
   Call ����X(KeMu, 3)
   Call ����X(KeMu, 5)
   Call ����X(KeMu, 8)
   Call ����X(KeMu, 10)
End Sub

Sub �ӱ߿�()
'
    Range("A1:L1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("C2:D2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("E2:F2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("H2:I2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("J2:K2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Selection.CurrentRegion.Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
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
    Range("A1:L1").Select
    With Selection.Font
        .Name = "΢���ź�"
        .Size = 14
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
End Sub
Sub ��ɫ()
    Dim i, j, k As Integer
    Imax = Cells(90000, 1).End(xlUp).Row
    i = 3: j = 65535
ColorAdd:
    Range(Cells(3, i), Cells(Imax, i + 1)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = j
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Select Case i
        Case Is = 3
            i = 8: j = 65535
            GoTo ColorAdd:
        Case Is = 8
            i = 5: j = 15773696
            GoTo ColorAdd:
        Case Is = 5
            i = 10: j = 15773696
            GoTo ColorAdd:
    End Select
End Sub

