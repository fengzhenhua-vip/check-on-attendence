Attribute VB_Name = "��׼����ģ��"
Public Const STRowMax As Single = 65536 '��office2007��׼�趨
Sub ��׼VBA����(ShName, RowB, ICol, OCol)
'�汾�� V1.0
'���ߣ�����
'ʱ�䣺2022��1��14��17��20
'���ܣ�����ĳһ����ֵ������С����������׼������
'�ĸ���������Ϊ��������������ʼ�У������У������
    Dim ARR As Variant
    Dim TempValue, TempMin As Single
    Dim i, j, k, p, q, RowE As Integer
    Dim ARROut(1 To STRowMax, 1 To 2) As Variant
    If ICol = OCol Then
        MsgBox "�������������������ͬ ��"                   '���������������ͬʱ���˳����򣬲�������
        End
    End If
    RowE = Sheets(ShName).Cells(STRowMax, ICol).End(xlUp).Row
    If RowB >= RowE Then
        MsgBox "�����������С�ڵ��ڿ�ʼ���� ��"
        End
    End If
    ARR = Sheets(ShName).Range(Sheets(ShName).Cells(RowB, ICol), Sheets(ShName).Cells(RowE, ICol))
    TempMin = Application.WorksheetFunction.Small(ARR, 1) - 1
    k = 1
ArrBegin:
    q = 0
    TempValue = Application.WorksheetFunction.Large(ARR, 1) ' �˴�������excel���������Ҳ�����Լ�����for next�������ֵ
    For i = 1 To UBound(ARR, 1)
        If ARR(i, 1) = TempValue Then
            ARR(i, 1) = TempMin: ARROut(i, 1) = k: q = q + 1
        End If
    Next
    k = k + q
    If k <= UBound(ARR, 1) Then
        GoTo ArrBegin:
    End If
    Sheets(ShName).Range(Sheets(ShName).Cells(RowB, OCol), Sheets(ShName).Cells(RowE, OCol)) = ARROut
End Sub
Sub ��׼VBA��ȡ(ShName, RowB, ICol, OShName, ORowB, OCol)
' �汾��V1.0
' ���ߣ�����
' ʱ�䣺2022��1��14��17:29
    Dim ARR As Variant
    Dim i, j, k, p, q, RowE As Integer
    Dim ARROut(1 To STRowMax, 1 To 2) As Variant
    Dim sht As Worksheet
    j = 0: q = 0
    For Each sht In ActiveWorkbook.Sheets
        If sht.Name = ShName Then
            j = 1
        End If
        If sht.Name = OShName Then
            q = 1
        End If
    Next
    If j = 0 Then
        MsgBox ShName & "�����ڣ�����д��ȷ�ġ����뱨������ ��"
        End
    End If
    If q = 0 Then
        MsgBox OShName & "�����ڣ�����д��ȷ�ġ������������ ��"
        End
    End If
    RowE = Sheets(ShName).Cells(STRowMax, ICol).End(xlUp).Row
    If ShName = OShName Then
        If ICol = OCol Then
            MsgBox "���������������ͬ ��"                   '���������������ͬʱ���˳����򣬲�������
            End
        End If
    End If
    If RowB >= RowE Then
        MsgBox "��������С�ڵ��ڿ�ʼ���� ��"
        End
    End If
    ARR = Sheets(ShName).Range(Sheets(ShName).Cells(RowB, ICol), Sheets(ShName).Cells(RowE, ICol))
    k = 1: p = 0
TQBegin:
    p = p + 1
    ARROut(p, 1) = ARR(k, 1)
    For i = k To UBound(ARR, 1)
        If ARR(i, 1) = ARROut(p, 1) Then
            ARR(i, 1) = Empty
        End If
    Next
    Do While Len(ARR(k, 1)) = 0 And k < UBound(ARR, 1)
        k = k + 1
    Loop
    If k <= UBound(ARR, 1) And Len(ARR(k, 1)) > 0 Then
        GoTo TQBegin:
    End If
    Sheets(OShName).Range(Sheets(OShName).Cells(ORowB, OCol), Sheets(OShName).Cells(ORowB + p - 1, OCol)) = ARROut
'    MsgBox Application.WorksheetFunction.Large(ARROut, 1)
End Sub
Sub ��׼VBAƴ������(IShName, IRowB, ICol)
' �汾��V1.0
' ���ߣ�����
' ʱ�䣺2022��1��14��19:19
' ���ܣ�ֻ��һ�а�ƴ�����򣬲��Ƕ�����չ������
    Dim RowE As Integer
    RowE = Sheets(IShName).Cells(STRowMax, ICol).End(xlUp).Row
    ActiveWorkbook.Worksheets(IShName).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(IShName).Sort.SortFields.Add2 Key:=Sheets(IShName).Range(Sheets(IShName).Cells(IRowB, ICol), Sheets(IShName).Cells(RowE, ICol)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(IShName).Sort
        .SetRange Sheets(IShName).Range(Sheets(IShName).Cells(IRowB, ICol), Sheets(IShName).Cells(RowE, ICol))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub ��׼VBA������()
    ��׼VBA�칫ϵͳ.Show
End Sub
Sub test1()
    Call ��׼VBA����("Sheet1", 1, 2, 3)
End Sub
Sub test2()
    Call ��׼VBA��ȡ("ɽ��", 3, 2, "Out", 1, 4)
End Sub
Sub test3()
   Call ��׼VBAƴ������("Out", 2, 2)
End Sub
