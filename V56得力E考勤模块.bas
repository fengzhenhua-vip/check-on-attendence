Attribute VB_Name = "V56����E����ģ��"
' ����E����ϵͳv55
' ���ߣ�����
' ����: 2021��4��25��--2021��4��28��
' ���ԣ���Ե���E+���ڻ����ɵ�ǩ�����ݣ����л��ܼ���ɫ����
' ��������δ����
' ˵������У׼���ɵ������������ļ�ֱ������ΪCorrect ,�������ĺô����ڿ��Ը�����ϰ�仯��ʱ����У׼����������ɺ���õ�ʱ�伸�����Ժ��Բ��ƣ��ͷ�����������������������

Public DeLiOriginal As Variant
Sub ����E����()
    Application.ScreenUpdating = False
    Call ������������
    Call ����EԤ����
    Call ��׼ͨ�ÿ���(DeLiOriginal)
    Application.ScreenUpdating = True
End Sub

Sub ����EԤ����()
    Sheets(1).name = NameOriginal
    Sheets(NameOriginal).Select
    Call ����Eɾ���ո�
    Call ����Eʱ��ˢ��
    Call ����EԴ������
    Call ����E��ȡ��Ч����
End Sub
Sub ����EԴ������()
' ȡ��ǰ���кϲ��ĵ�Ԫ��
    Rows("1:2").Select
    Selection.UnMerge
    Dim i As Integer
'��ͷ��A
    For i = 1 To 13
        Cells(1, i + 13) = Cells(1, i)
        Cells(1, i) = "A" & Cells(1, i)
    Next
'����������
    Columns("C:C").Select
    ActiveWorkbook.Worksheets(NameOriginal).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(NameOriginal).Sort.SortFields.Add Key:=Range("C1"), SortOn _
        :=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(NameOriginal).Sort
        .SetRange Range("A1:Z1928")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
'��ͷȥA
    For i = 14 To 26
        Cells(1, i - 13) = Cells(1, i)
        Cells(1, i) = ""
    Next
End Sub
Sub ����Eʱ��ˢ��()
'
' ��Ե���e+ϵͳ���е�ʱ������ ���������ݵ�ָ������ǰ��Ӧ����ת������Ӧ��ʱ���ʽ

    Columns("L:M").NumberFormatLocal = TimeFormat
    Dim i As Integer
    For i = 2 To Sheets(NameOriginal).Range("a65536").End(xlUp).Row
     Cells(i, 12) = Cells(i, 12).Value
     Cells(i, 13) = Cells(i, 13).Value
    Next
End Sub
Sub ����Eɾ���ո�()
' ɾ�����пո�,��EF����ʱ���ʽʱ�����пո�Ҳ�ᱻ��ȫɾ��������ʱ���ϵĴ�������Ҫ�ָ�
    Range("A1").Select
    Cells.Replace What:=" ", Replacement:="", LookAt:=xlPart, SearchOrder:= _
       xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Cells.Replace What:="AM", Replacement:=" AM", LookAt:=xlPart, SearchOrder:= _
       xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Cells.Replace What:="PM", Replacement:=" PM", LookAt:=xlPart, SearchOrder:= _
       xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
End Sub
Sub ����E��ȡ��Ч����()
' ���ɱ�׼ͨ�ÿ��������ʽ
    Dim DeLiSource As Variant
    Dim DeRmax As Integer
    Dim DeCmax As Integer
    Dim i, j As Integer
    DeRmax = Sheets(NameOriginal).Range("a65536").End(xlUp).Row
    DeCmax = 13
    Columns("F:F").NumberFormatLocal = DateFormat
    For i = 1 To DeRmax
     Cells(i, 6) = Cells(i, 6).Value
    Next
    DeLiSource = Range(Cells(1, 3), Cells(DeRmax, DeCmax)).Value
    ReDim DeLiOriginal(1 To DeRmax, 1 To 15) As Variant
    j = 0
    For i = 1 To DeRmax
        j = j + 1
        DeLiOriginal(j, 1) = DeLiSource(i, 1)
        DeLiOriginal(j, 2) = DeLiSource(i, 4)
        DeLiOriginal(j, 3) = DeLiSource(i, 7)
        DeLiOriginal(j, 4) = DeLiSource(i, 5)
        DeLiOriginal(j, 5) = DeLiSource(i, 10)
        DeLiOriginal(j, 6) = DeLiSource(i, 11)
    Next
End Sub

