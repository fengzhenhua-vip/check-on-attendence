Attribute VB_Name = "V66����E����ģ��"
' ����E����ϵͳV66
' ���ߣ�����
' ����: 2021��4��25��--2021��6��12��
' ���ԣ���Ե���E+���ڻ����ɵ�ǩ�����ݣ����л��ܼ���ɫ����
' ��������δ����
' ˵������У׼���ɵ������������ļ�ֱ������ΪCorrect ,�������ĺô����ڿ��Ը�����ϰ�仯��ʱ����У׼����������ɺ���õ�ʱ�伸�����Ժ��Բ��ƣ��ͷ�����������������������
' ��־����2021��6��11�յ�12�գ��Ա�׼ͨ�ÿ���ģ���������������ǿ�ļ����ԣ�ͬʱ�����Ч�ʣ����Դ˴ζԵ���EԤ����ϵͳ����������������Ӽ�������ͬ���汾��ΪV66.
' ��־������Trim����������ȥ���ַ������߿ո񣩣�������Mid����������ȡ���ַ�����V66

Public DeLiOriginal As Variant
Public DeRmax As Integer
Public DeCmax As Integer

Sub ����E����()
    Application.ScreenUpdating = False
    Call ������������
    Call GetDeliOriginal
    Call ��׼ͨ�ÿ���(DeLiOriginal)
    Application.ScreenUpdating = True
End Sub

Sub DeLiEplusSort()
' ȡ��ǰ���кϲ��ĵ�Ԫ��
    Rows("1:2").UnMerge
'�����������ڡ���ι�ͬ����
    ActiveWorkbook.Worksheets(NameOriginal).Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets(NameOriginal).Sort.SortFields. _
        Add Key:=Range("C2:C" & DeRmax), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets(NameOriginal).Sort.SortFields. _
        Add Key:=Range("F2:F" & DeRmax), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets(NameOriginal).Sort.SortFields. _
        Add Key:=Range("I2:I" & DeRmax), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(NameOriginal).Sort
        .SetRange Range("A1:Q" & DeRmax)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub GetDeliOriginal()
' ����Sheets��
    Sheets(1).name = NameOriginal
    Sheets(NameOriginal).Select
' ȡ��Դ������������������
    DeRmax = Sheets(NameOriginal).Range("a65536").End(xlUp).Row
    DeCmax = 13
    Call DeLiEplusSort                                                                                        '��Դ������
' ���ɱ�׼ͨ�ÿ��������ʽ
    Dim DeLiSource As Variant
    Dim i, j As Integer
    DeLiSource = Range(Cells(1, 3), Cells(DeRmax, DeCmax)).Value
    ReDim DeLiOriginal(1 To DeRmax, 1 To 15) As Variant
    j = 1
    DeLiOriginal(j, 1) = DeLiSource(1, 1)                                                                     'ͨ�ó������и����Ĳ��֣����Դ˴�����ʡ�Զ�û��Ӱ��
    DeLiOriginal(j, 2) = DeLiSource(1, 4)
    DeLiOriginal(j, 3) = DeLiSource(1, 7)
    DeLiOriginal(j, 4) = DeLiSource(1, 5)
    DeLiOriginal(j, 5) = DeLiSource(1, 10)
    DeLiOriginal(j, 6) = DeLiSource(1, 11)
    For i = 2 To DeRmax
        If DeLiSource(i, 1) > 0 Then                                                                          '��������Ϊ�յ���
            j = j + 1
            DeLiOriginal(j, 1) = VBA.Trim(Mid(DeLiSource(i, 1), 1, 1)) & VBA.Trim(Mid(DeLiSource(i, 1), 2))   '�򵥺�ȥ�����������м�Ŀո�,����ʹ��ͳһ�Ŀո�����
            DeLiOriginal(j, 2) = CDate(DeLiSource(i, 4))
            DeLiOriginal(j, 3) = DeLiSource(i, 7)
            DeLiOriginal(j, 4) = DeLiSource(i, 5)
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' �������ý����ʱ���ʽδ����ʱ�Զ�ת��Ϊʱ���ʽ,���������׼ʱ����10�ĸ�17�η��������ڱȽ�ʱ    '
' ���Ի��ͳ����Ϣ����ͳ�ƽ����ɫʱ��ʱ��ڵ��ϻ�������С����Խ���׼������Ϣ�еı�׼ʱ��ڵ��Ϊ     '
' 59�룬��CDate��õ�ʱ�����00��������ʴ�ʱ������ϸ�ͳ�ƽ�����������������������V66�����޸ġ� '
'                                                                                                      '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If DeLiSource(i, 10) > 0 Then
                DeLiOriginal(j, 5) = CDate(DeLiSource(i, 10))
            End If
            If DeLiSource(i, 11) > 0 Then
                DeLiOriginal(j, 6) = CDate(DeLiSource(i, 11))
            End If
        End If
    Next
End Sub

