Attribute VB_Name = "��׼��VBA��ģ��"
Public Arr, ArrOut As Variant 'Ĭ��ʹ�õ���ʱ����,�ڳ����п���ʹ��Redim������¶���

Sub ��׼VBA����(ShName, RowB, ICol, OCol)
'�汾�� V2.0
'���ߣ�����
'ʱ�䣺2022��1��14��17��20
'���ܣ�����ĳһ����ֵ������С����������׼������
'�ĸ���������Ϊ��������������ʼ�У������У������
    Dim Arr As Variant
    Dim TempValue, TempMin As Double
    Dim i, j, k, p, q, RowE, RowEX As Integer
    If ICol = OCol Then
        MsgBox "�������������������ͬ ��"                   '���������������ͬʱ���˳����򣬲�������
        End
    End If
    RowEX = Sheets(ShName).Range("A1").SpecialCells(xlLastCell).Row + 1 'ȡ�����һ��
    RowE = Sheets(ShName).Cells(RowEX, ICol).End(xlUp).Row
    ReDim ArrOut(1 To RowE, 1 To 2) As Variant
    If RowB >= RowE Then
        MsgBox "�����������С�ڵ��ڿ�ʼ���� ��"
        End
    End If
    Arr = Sheets(ShName).Range(Sheets(ShName).Cells(RowB, ICol), Sheets(ShName).Cells(RowE, ICol))
    TempMin = CDbl(Application.WorksheetFunction.Small(Arr, 1)) - CDbl(Application.WorksheetFunction.Large(Arr, 1)) - 1
    k = 1
ArrBegin:
    q = 0
    TempValue = CDbl(Application.WorksheetFunction.Large(Arr, 1)) ' �˴�������excel���������Ҳ�����Լ�����for next�������ֵ
    For i = 1 To UBound(Arr, 1)
        If Arr(i, 1) = TempValue Then
            Arr(i, 1) = TempMin: ArrOut(i, 1) = k: q = q + 1
        End If
    Next
    k = k + q
    If k <= UBound(Arr, 1) Then
        GoTo ArrBegin:
    End If
    Sheets(ShName).Range(Sheets(ShName).Cells(RowB, OCol), Sheets(ShName).Cells(RowE, OCol)) = ArrOut
End Sub
Sub ��׼VBA��ȡ(ShName, RowB, ICol, OShName, ORowB, OCol)
' �汾��V1.0
' ���ߣ�����
' ʱ�䣺2022��1��14��17:29
    Dim Arr As Variant
    Dim i, j, k, p, q, RowE As Integer
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
    RowE = Sheets(ShName).Range("A1").SpecialCells(xlLastCell).Row 'ȡ�����һ��
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
    Arr = Sheets(ShName).Range(Sheets(ShName).Cells(RowB, ICol), Sheets(ShName).Cells(RowE, ICol))
    ReDim ArrOut(1 To RowE, 1 To 2) As Variant
    k = 1: p = 0
TQBegin:
    p = p + 1
    ArrOut(p, 1) = Arr(k, 1)
    For i = k To UBound(Arr, 1)
        If Arr(i, 1) = ArrOut(p, 1) Then
            Arr(i, 1) = Empty
        End If
    Next
    Do While Len(Arr(k, 1)) = 0 And k < UBound(Arr, 1)
        k = k + 1
    Loop
    If k <= UBound(Arr, 1) And Len(Arr(k, 1)) > 0 Then
        GoTo TQBegin:
    End If
    Sheets(OShName).Range(Sheets(OShName).Cells(ORowB, OCol), Sheets(OShName).Cells(ORowB + p - 1, OCol)) = ArrOut
End Sub
Sub ��׼VBA����(IShName, IRowB, ICol, ShunXu)
' �汾��V1.0
' ���ߣ�����
' ʱ�䣺2022��1��14��19:19
' ���ܣ�ֻ��һ�а�ƴ�����򣬲��Ƕ�����չ������
    Dim RowE, RowEX As Integer
    Dim ShengJiang As String
    If ShunXu = 1 Then
        ShengJiang = xlDescending   '����1
    Else
        ShengJiang = xlAscending    '����0
    End If
    RowEX = Sheets(IShName).Range("A1").SpecialCells(xlLastCell).Row + 1 'ȡ�����һ��
    RowE = Sheets(IShName).Cells(RowEX, ICol).End(xlUp).Row
    ActiveWorkbook.Worksheets(IShName).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(IShName).Sort.SortFields.Add2 Key:=Sheets(IShName).Range(Sheets(IShName).Cells(IRowB, ICol), Sheets(IShName).Cells(RowE, ICol)), _
        SortOn:=xlSortOnValues, Order:=ShengJiang, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(IShName).Sort
        .SetRange Sheets(IShName).Range(Sheets(IShName).Cells(IRowB, ICol), Sheets(IShName).Cells(RowE, ICol))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub ��׼VBA���(IShName, IRowB, ICol1, ICol2, OShName, ORowB, OCol1, OCol2)
' �汾��V1.0
' ���ߣ�����
' ʱ�䣺2022��1��15��14:32
' ���ܣ�ֻ��һ�а�ƴ�����򣬲��Ƕ�����չ������
    Dim IArr, OArr As Variant
    Dim IColMin, IColMax, OColMin, OColMax As Integer
    Dim IRowE, ORowE As Integer
    Dim i, j, k, m, n, p, q As Integer
    If ICol1 < ICol2 Then
        IColMin = ICol1: IColMax = ICol2
    ElseIf ICol1 > ICol2 Then
        IColMin = ICol2: IColMax = ICol1
    Else
        MsgBox IShName & ICol1 & "��" & ICol2 & "��ͬ������������ ��"
    End If
    IRowE = Sheets(IShName).Range("A1").SpecialCells(xlLastCell).Row 'ȡ�����һ��
    IArr = Sheets(IShName).Range(Sheets(IShName).Cells(IRowB, IColMin), Sheets(IShName).Cells(IRowE, IColMax))
    If OCol1 < OCol2 Then
        OColMin = OCol1: OColMax = OCol2
    ElseIf OCol1 > OCol2 Then
        OColMin = OCol2: OColMax = OCol1
    Else
        MsgBox OShName & OCol1 & "��" & OCol2 & "��ͬ������������ ��"
    End If
    ORowE = Sheets(OShName).Range("A1").SpecialCells(xlLastCell).Row 'ȡ�����һ��
    OArr = Sheets(OShName).Range(Sheets(OShName).Cells(ORowB, OColMin), Sheets(OShName).Cells(ORowE, OColMax))
    m = ICol1 - IColMin + 1: n = OCol1 - OColMin + 1
    p = ICol2 - IColMin + 1: q = OCol2 - OColMin + 1
    For i = 1 To UBound(OArr, 1)
        For j = 1 To UBound(IArr, 1)
            If OArr(i, n) = IArr(j, m) Then
                OArr(i, q) = IArr(j, p)
            End If
        Next
    Next
    Sheets(OShName).Range(Sheets(OShName).Cells(ORowB, OColMin), Sheets(OShName).Cells(ORowE, OColMax)) = OArr
End Sub
Sub ��׼VBA����������(IShName, IRowB, ICol1, ICol2, OShName, ORowB, OCol1, OCol2, SMin, SMax)
' �汾��V1.0
' ���ߣ�����
' ʱ�䣺2022��1��15��14:32
' ���ܣ�ֻ��һ�а�ƴ�����򣬲��Ƕ�����չ������
    Dim IArr, OArr As Variant
    Dim IColMin, IColMax, OColMin, OColMax As Integer
    Dim IRowE, ORowE As Integer
    Dim i, j, k, m, n, p, q As Integer
    If ICol1 < ICol2 Then
        IColMin = ICol1: IColMax = ICol2
    ElseIf ICol1 > ICol2 Then
        IColMin = ICol2: IColMax = ICol1
    Else
        MsgBox IShName & ICol1 & "��" & ICol2 & "��ͬ������������ ��"
    End If
    IRowE = Sheets(IShName).Range("A1").SpecialCells(xlLastCell).Row 'ȡ�����һ��
    IArr = Sheets(IShName).Range(Sheets(IShName).Cells(IRowB, IColMin), Sheets(IShName).Cells(IRowE, IColMax))
    If OCol1 < OCol2 Then
        OColMin = OCol1: OColMax = OCol2
    ElseIf OCol1 > OCol2 Then
        OColMin = OCol2: OColMax = OCol1
    Else
        MsgBox OShName & OCol1 & "��" & OCol2 & "��ͬ������������ ��"
    End If
    ORowE = Sheets(OShName).Range("A1").SpecialCells(xlLastCell).Row 'ȡ�����һ��
    OArr = Sheets(OShName).Range(Sheets(OShName).Cells(ORowB, OColMin), Sheets(OShName).Cells(ORowE, OColMax))
    m = ICol1 - IColMin + 1: n = OCol1 - OColMin + 1
    p = ICol2 - IColMin + 1: q = OCol2 - OColMin + 1
    For i = 1 To UBound(OArr, 1)
        OArr(i, q) = Empty
        For j = 1 To UBound(IArr, 1)
            If InStr(OArr(i, n), IArr(j, m)) > 0 Then
                If CDbl(SMin) <= CDbl(IArr(j, p)) And CDbl(IArr(j, p)) <= CDbl(SMax) Then
                    OArr(i, q) = OArr(i, q) + 1
                End If
            End If
        Next
    Next
    Sheets(OShName).Range(Sheets(OShName).Cells(ORowB, OColMin), Sheets(OShName).Cells(ORowE, OColMax)) = OArr
    Sheets(OShName).Select
End Sub
Public Function ScoreLine(ShName, NumLine, RowB, ICol)
' �汾��V1.0
' ���ߣ�����
' ʱ�䣺2022��1��15��14:32
' ���ܣ�ȡ��ĳһ�з����ߣ�Ŀǰ�ǰ�ƽԭ�ص�һ��ѧ�ı���ȡ�á�����һ��ĳ�Ƶ������ķ���Ϊ��׼�������������������������Ƚϣ��Խ����߷�����Ϊ��ǰ���Է�����
    Dim Arr, ArrBak As Variant
    Dim RowE As Integer
    Dim i, j, k As Integer
    Dim UpNum, DownNum As Integer
    RowE = Sheets(ShName).Range("A1").SpecialCells(xlLastCell).Row 'ȡ�����һ��
    ArrBak = Sheets(ShName).Range(Sheets(ShName).Cells(RowB, ICol), Sheets(ShName).Cells(RowE, ICol))
    Call ��׼VBA����(ShName, RowB, ICol, 1)
    Arr = Sheets(ShName).Range(Sheets(ShName).Cells(RowB, ICol), Sheets(ShName).Cells(RowE, ICol))
    Sheets(ShName).Range(Sheets(ShName).Cells(RowB, ICol), Sheets(ShName).Cells(RowE, ICol)) = ArrBak
    i = 0
UpBegin:
    If NumLine > i Then
        If Arr(NumLine - i, 1) = Arr(NumLine, 1) Then
            i = i + 1
            GoTo UpBegin:
        Else
            UpNum = i
        End If
    Else
        UpNum = i
    End If
    i = 0
DownBegin:
    If NumLine > i Then
        If Arr(NumLine + i, 1) = Arr(NumLine, 1) Then
            i = i + 1
            GoTo DownBegin:
        Else
            DownNum = i
        End If
    Else
        DownNum = i
    End If
    If UpNum < DownNum Then
        ScoreLine = Arr(NumLine - UpNum, 1)
    Else
        ScoreLine = Arr(NumLine, 1)
     End If
End Function
Sub ��׼VBA��ֹ�����(CFName)
    Dim sht As Worksheet
    Dim CFPath, CFFolder As String
    Set SFO = CreateObject("Scripting.FileSystemObject")
    CFPath = ActiveWorkbook.Path
    CFFolder = CFPath & "\" & CFName & "��ֽ��"
    If SFO.folderExists(CFFolder) = False Then
        MkDir CFFolder
    End If
    Application.DisplayAlerts = False
    For Each sht In ActiveWorkbook.Sheets
        sht.Select: sht.Copy
        ActiveWorkbook.SaveAs Filename:=CFFolder & "\" & sht.Name & CFName, FileFormat:=xlOpenXMLWorkbook '�����������ΪEXCELĬ��2007��ʽ 2003��ʽxlNormal
        ActiveWorkbook.Close
    Next
    Application.DisplayAlerts = True
        MsgBox ActiveWorkbook.Name & "������ !"
End Sub
Sub ��׼VBA�ϲ�������()
    Dim fpath, fname As String
    Dim CurFil, OtherFil As String
    Dim OArr As Variant
    Dim CurBook As Workbook
    Dim Imax, Jmax As Integer
    Dim i, j, k, m, n, p, q, r As Integer
    Dim sht, osht As Worksheet
    Dim ActiveShtName As String
    For Each sht In ActiveWorkbook.Sheets
        m = Sheets(sht.Name).Range("A1").SpecialCells(xlLastCell).Row 'ȡ�����һ��
        n = Sheets(sht.Name).Range("A1").SpecialCells(xlLastCell).Column 'ȡ�����һ��
        If m > 1 Or n > 1 Then
            MsgBox "��ǰ�������ǿգ������´���һ���յĹ���������ִ�кϲ����������� ��"
            End
        End If
    Next
    fpath = ActiveWorkbook.Path
    CurFil = ActiveWorkbook.Name
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    fname = Dir(fpath & "\*.xl*")
    i = 1: ReDim Arr(i) As Variant
    Arr(i) = fname
    Do While fname <> ""
        fname = Dir
        If fname = "" Then
            Exit Do
        End If
        i = i + 1
        ReDim Preserve Arr(i) As Variant
        Arr(i) = fname
    Loop
    For p = 1 To i
        If Arr(p) <> CurFil Then
            Set CurBook = GetObject(fpath & "\" & Arr(p))
            If InStr(Arr(p), "xlsx") > 0 Then
                OtherFil = Left(Arr(p), Len(Arr(p)) - 5)
            Else
                OtherFil = Left(Arr(p), Len(Arr(p)) - 4)
            End If
            For Each osht In CurBook.Sheets
                Imax = CurBook.Sheets(osht.Name).Range("A1").SpecialCells(xlLastCell).Row    'ȡ�����һ��
                Jmax = CurBook.Sheets(osht.Name).Range("A1").SpecialCells(xlLastCell).Column 'ȡ�����һ��
                If Imax > 1 Or Jmax > 1 Then
                    ActiveShtName = OtherFil & "(" & osht.Name & ")"
                    OArr = CurBook.Sheets(osht.Name).Range(CurBook.Sheets(osht.Name).Cells(1, 1), CurBook.Sheets(osht.Name).Cells(Imax, Jmax))
                    q = 0
                    For Each sht In ActiveWorkbook.Sheets
                        If sht.Name = ActiveShtName Then
                            q = 1
                        End If
                    Next
                    If q = 0 Then
                        Sheets.Add After:=ActiveSheet
                        ActiveSheet.Name = ActiveShtName
                    End If
                    Sheets(ActiveShtName).Range(Sheets(ActiveShtName).Cells(1, 1), Sheets(ActiveShtName).Cells(Imax, Jmax)) = OArr
                End If
            Next
        End If
    Next
    For Each sht In ActiveWorkbook.Sheets
        m = sht.Range("A1").SpecialCells(xlLastCell).Row    'ȡ�����һ��
        n = sht.Range("A1").SpecialCells(xlLastCell).Column 'ȡ�����һ��
        If m = 1 And n = 1 Then
            If Len(sht.Cells(1, 1)) = 0 Then
                sht.Delete
            End If
        End If
    Next
    Application.ScreenUpdating = True
    ActiveWorkbook.Save
    Application.DisplayAlerts = True
End Sub
Sub ��׼VBA���кϲ�������()
    Dim sht As Worksheet
    Dim Hfile, Hpath As String
    Dim i, j, k As Integer
    Dim Imax, Jmax As Integer
    Dim IBegin, IEnd As Integer
    Dim Arr As Variant
    Dim CurBook As Workbook
    Set SFO = CreateObject("Scripting.FileSystemObject")
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Hpath = ActiveWorkbook.Path
    Set CurBook = GetObject(Hpath & "\" & ActiveWorkbook.Name)
    If InStr(ActiveWorkbook.Name, ".xlsx") Then
        Hfile = Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 5) & "(���кϲ�)" & ".xlsx"
    Else
        Hfile = Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 4) & "(���кϲ�)" & ".xlsx"
    End If
    Hfile = Hpath & "\" & Hfile
    If SFO.fileExists(Hfile) = False Then
        Application.DisplayAlerts = False
        Workbooks.Add.SaveAs Filename:=Hfile
        Sheets(1).Name = "���кϲ�"
        For Each sht In CurBook.Sheets
            Imax = CurBook.Sheets(sht.Name).Range("A1").SpecialCells(xlLastCell).Row    'ȡ�����һ��
            Jmax = CurBook.Sheets(sht.Name).Range("A1").SpecialCells(xlLastCell).Column 'ȡ�����һ��
            Arr = CurBook.Sheets(sht.Name).Range(CurBook.Sheets(sht.Name).Cells(1, 1), CurBook.Sheets(sht.Name).Cells(Imax, Jmax))
            IBegin = Sheets("���кϲ�").Range("A1").SpecialCells(xlLastCell).Row
            i = Sheets("���кϲ�").Range("A1").SpecialCells(xlLastCell).Column
            If i = 1 And IBegin = 1 Then
                If Len(Sheets("���кϲ�").Cells(1, 1)) = 0 Then
                    IBegin = 0
                End If
            End If
            IBegin = IBegin + 1
            IEnd = IBegin + Imax - 1
            Range(Cells(IBegin, 1), Cells(IEnd, Jmax)) = Arr
        Next
    Else
        MsgBox Hfile & "�Ѿ����ڣ���ɾ��������ִ�кϲ����� ��"
    End If
    Application.ScreenUpdating = True
    ActiveWorkbook.Save
    Application.DisplayAlerts = True
End Sub
Sub ��׼VBA���кϲ�������()
    Dim sht As Worksheet
    Dim Hfile, Hpath As String
    Dim i, j, k As Integer
    Dim Imax, Jmax As Integer
    Dim IBegin, IEnd As Integer
    Dim Arr As Variant
    Dim CurBook As Workbook
    Set SFO = CreateObject("Scripting.FileSystemObject")
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Hpath = ActiveWorkbook.Path
    Set CurBook = GetObject(Hpath & "\" & ActiveWorkbook.Name)
    If InStr(ActiveWorkbook.Name, ".xlsx") Then
        Hfile = Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 5) & "(���кϲ�)" & ".xlsx"
    Else
        Hfile = Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 4) & "(���кϲ�)" & ".xlsx"
    End If
    Hfile = Hpath & "\" & Hfile
    If SFO.fileExists(Hfile) = False Then
        Application.DisplayAlerts = False
        Workbooks.Add.SaveAs Filename:=Hfile
        Sheets(1).Name = "���кϲ�"
        For Each sht In CurBook.Sheets
            Imax = CurBook.Sheets(sht.Name).Range("A1").SpecialCells(xlLastCell).Row    'ȡ�����һ��
            Jmax = CurBook.Sheets(sht.Name).Range("A1").SpecialCells(xlLastCell).Column 'ȡ�����һ��
            Arr = CurBook.Sheets(sht.Name).Range(CurBook.Sheets(sht.Name).Cells(1, 1), CurBook.Sheets(sht.Name).Cells(Imax, Jmax))
            IBegin = Sheets("���кϲ�").Range("A1").SpecialCells(xlLastCell).Column
            i = Sheets("���кϲ�").Range("A1").SpecialCells(xlLastCell).Row
            If i = 1 And IBegin = 1 Then
                If Len(Sheets("���кϲ�").Cells(1, 1)) = 0 Then
                    IBegin = 0
                End If
            End If
            IBegin = IBegin + 1
            IEnd = IBegin + Jmax - 1
            Range(Cells(1, IBegin), Cells(Imax, IEnd)) = Arr
        Next
    Else
        MsgBox Hfile & "�Ѿ����ڣ���ɾ��������ִ�кϲ����� ��"
    End If
    Application.ScreenUpdating = True
    ActiveWorkbook.Save
    Application.DisplayAlerts = True
End Sub
Sub ��׼VBAת��()
' �汾��V2.0
' ��־��ԭ�����������鷽ʽֱ�ӱ任�ģ�����Ϊ����߼�����ֱ�ӵ�����Excel�ĺ������������Ӽ�����
    Dim Imax, Jmax, Zmax As Integer
    Dim Arr As Variant
    Imax = ActiveSheet.Range("A1").SpecialCells(xlLastCell).Row         'ȡ�����һ��
    Jmax = ActiveSheet.Range("A1").SpecialCells(xlLastCell).Column      'ȡ�����һ��
    If Imax > Jmax Then
        Zmax = Imax
    Else
        Zmax = Jmax
    End If
    Arr = Range(Cells(1, 1), Cells(Zmax, Zmax))
    Range(Cells(1, 1), Cells(Zmax, Zmax)) = Application.Transpose(Arr)  'Transposeȡת��
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
   Call ��׼VBA����("Sheet2", 1, 1, 0)
End Sub
Sub test4()
  Call ��׼VBA���("Sheet1", 2, 6, 7, "Sheet2", 1, 1, 3)
End Sub
Sub test5()
  Call ��׼VBA����������("Sheet1", 3, 2, 5, "Sheet2", 2, 1, 2, 0, 800)
End Sub
Sub test6()
    k = ScoreLine("Sheet1", 300, 2, 14)
    MsgBox k
End Sub
Sub test7()
    Call ��׼VBA��ֹ�����("��ʦ��A��")
End Sub
