Attribute VB_Name = "TJS_V10"
Sub ִ��ͳ��()
    Call ͳ�Ʒ�����
'
    Dim sht As Worksheet
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
            sht.Copy
            ActiveWorkbook.SaveAs Filename:=OutFolder & "\" & sht.Name & "�������߶Աȱ�", FileFormat:=xlNormal  '�����������ΪEXCELĬ�ϸ�ʽ
            ActiveWorkbook.Close
        End If
    Next
    Application.DisplayAlerts = True
        MsgBox "�����ɼ�ͳ�����!"
End Sub


Sub �����ϴ�����X(Yuan, KeMu, KeMuLie)
 Dim Imax, Jmax As Integer
 Dim i, j, k As Integer
 Dim KeMuCol As Integer
 Dim YuanSJ As Variant
 Imax = Sheets(Yuan).Cells(90000, 1).End(xlUp).Row
 Jmax = Sheets(Yuan).Cells(3, 200).End(xlToLeft).Column
 YuanSJ = Sheets(Yuan).Range(Sheets(Yuan).Cells(1, 1), Sheets(Yuan).Cells(Imax, Jmax))
 For j = 3 To Jmax
    If InStr(YuanSJ(3, j), KeMu) > 0 Then
        KeMuCol = j
    End If
 Next
 For i = 3 To Imax
    For k = 3 To Imax
        If Sheets(KeMu).Cells(i, 1) = YuanSJ(k, 1) Then
            Sheets(KeMu).Cells(i, KeMuLie) = YuanSJ(k, KeMuCol)
        End If
    Next
 Next
End Sub
Sub �����ϴ�����XX(XueKe)
    Sheets(XueKe).Cells(1, 1) = XueKe & "ͳ�ƶԱȱ�"
    Call �����ϴ�����X("һ��", XueKe, 8)
    Call �����ϴ�����X("����", XueKe, 3)
End Sub

Sub �����ϴ�����()
    Call �����ϴ�����XX("�ܷ�")
    Call �����ϴ�����XX("����")
    Call �����ϴ�����XX("��ѧ")
    Call �����ϴ�����XX("Ӣ��")
    Call �����ϴ�����XX("����")
    Call �����ϴ�����XX("��ѧ")
    Call �����ϴ�����XX("����")
    Call �����ϴ�����XX("����")
    Call �����ϴ�����XX("��ʷ")
    Call �����ϴ�����XX("����")
End Sub
