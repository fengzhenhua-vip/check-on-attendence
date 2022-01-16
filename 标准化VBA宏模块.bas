Attribute VB_Name = "��׼��VBA��ģ��"
Public Const STRowMax As Double = 65536 '��office2007��׼�趨1048576
Public Const STColMax As Double = 256 '��office2007��׼�趨16384
Public Const CheckLine As Double = 100  '100�У��У���Ӧ�ð�ͨ����������������ǰ��������ֵ�ˣ�����Ϊ�����Ч�ʣ����ô�ֵ�㹻��

Sub ��׼VBA����(ShName, RowB, ICol, OCol)
'�汾�� V1.0
'���ߣ�����
'ʱ�䣺2022��1��14��17��20
'���ܣ�����ĳһ����ֵ������С����������׼������
'�ĸ���������Ϊ��������������ʼ�У������У������
    Dim Arr As Variant
    Dim TempValue, TempMin As Double
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
    Arr = Sheets(ShName).Range(Sheets(ShName).Cells(RowB, ICol), Sheets(ShName).Cells(RowE, ICol))
    TempMin = CDbl(Application.WorksheetFunction.Small(Arr, 1) - 1)
    k = 1
ArrBegin:
    q = 0
    TempValue = CDbl(Application.WorksheetFunction.Large(Arr, 1)) ' �˴�������excel���������Ҳ�����Լ�����for next�������ֵ
    For i = 1 To UBound(Arr, 1)
        If Arr(i, 1) = TempValue Then
            Arr(i, 1) = TempMin: ARROut(i, 1) = k: q = q + 1
        End If
    Next
    k = k + q
    If k <= UBound(Arr, 1) Then
        GoTo ArrBegin:
    End If
    Sheets(ShName).Range(Sheets(ShName).Cells(RowB, OCol), Sheets(ShName).Cells(RowE, OCol)) = ARROut
End Sub
Sub ��׼VBA��ȡ(ShName, RowB, ICol, OShName, ORowB, OCol)
' �汾��V1.0
' ���ߣ�����
' ʱ�䣺2022��1��14��17:29
    Dim Arr As Variant
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
    Arr = Sheets(ShName).Range(Sheets(ShName).Cells(RowB, ICol), Sheets(ShName).Cells(RowE, ICol))
    k = 1: p = 0
TQBegin:
    p = p + 1
    ARROut(p, 1) = Arr(k, 1)
    For i = k To UBound(Arr, 1)
        If Arr(i, 1) = ARROut(p, 1) Then
            Arr(i, 1) = Empty
        End If
    Next
    Do While Len(Arr(k, 1)) = 0 And k < UBound(Arr, 1)
        k = k + 1
    Loop
    If k <= UBound(Arr, 1) And Len(Arr(k, 1)) > 0 Then
        GoTo TQBegin:
    End If
    Sheets(OShName).Range(Sheets(OShName).Cells(ORowB, OCol), Sheets(OShName).Cells(ORowB + p - 1, OCol)) = ARROut
'    MsgBox Application.WorksheetFunction.Large(ARROut, 1)
End Sub
Sub ��׼VBA����(IShName, IRowB, ICol, ShunXu)
' �汾��V1.0
' ���ߣ�����
' ʱ�䣺2022��1��14��19:19
' ���ܣ�ֻ��һ�а�ƴ�����򣬲��Ƕ�����չ������
    Dim RowE As Integer
    Dim ShengJiang As String
    If ShunXu = 1 Then
        ShengJiang = xlDescending   '����
    Else
        ShengJiang = xlAscending    '����
    End If
    RowE = Sheets(IShName).Cells(STRowMax, ICol).End(xlUp).Row
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
    IRowE = Sheets(IShName).Cells(STRowMax, ICol1).End(xlUp).Row
    IArr = Sheets(IShName).Range(Sheets(IShName).Cells(IRowB, IColMin), Sheets(IShName).Cells(IRowE, IColMax))
    If OCol1 < OCol2 Then
        OColMin = OCol1: OColMax = OCol2
    ElseIf OCol1 > OCol2 Then
        OColMin = OCol2: OColMax = OCol1
    Else
        MsgBox OShName & OCol1 & "��" & OCol2 & "��ͬ������������ ��"
    End If
    ORowE = Sheets(OShName).Cells(STRowMax, OCol1).End(xlUp).Row
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
    IRowE = Sheets(IShName).Cells(STRowMax, ICol1).End(xlUp).Row
    IArr = Sheets(IShName).Range(Sheets(IShName).Cells(IRowB, IColMin), Sheets(IShName).Cells(IRowE, IColMax))
    If OCol1 < OCol2 Then
        OColMin = OCol1: OColMax = OCol2
    ElseIf OCol1 > OCol2 Then
        OColMin = OCol2: OColMax = OCol1
    Else
        MsgBox OShName & OCol1 & "��" & OCol2 & "��ͬ������������ ��"
    End If
    ORowE = Sheets(OShName).Cells(STRowMax, OCol1).End(xlUp).Row
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
    RowE = Sheets(ShName).Cells(STRowMax, ICol).End(xlUp).Row
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
        ActiveWorkbook.SaveAs Filename:=CFFolder & "\" & sht.Name & CFName, FileFormat:=xlNormal  '�����������ΪEXCELĬ�ϸ�ʽ
        ActiveWorkbook.Close
    Next
    Application.DisplayAlerts = True
        MsgBox ActiveWorkbook.Name & "������ !"
End Sub
Sub ��׼VBA�ϲ�������()
    Dim fpath, fname As String
    Dim Arr(STRowMax) As String
    Dim CurFil, OtherFil As String
    Dim OArr As Variant
    Dim CurBook As Workbook
    Dim Imax, Jmax As Integer
    Imax = 1000: Jmax = 100
    Dim i, j, k, m, n, p, q As Integer
    Dim sht, osht As Worksheet
    Dim ActiveShtName As String
    For Each sht In ActiveWorkbook.Sheets
        n = 0
        For m = 1 To 10
            If Len(Sheets(sht.Name).Cells(m, 1)) > 0 Then
                n = 1
            End If
        Next
        If n = 1 Then
                 MsgBox "��ǰ�������ǿգ������´���һ���յĹ���������ִ�кϲ����������� ��"
                 End
        End If
    Next
    fpath = ActiveWorkbook.Path
    CurFil = ActiveWorkbook.Name
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    fname = Dir(fpath & "\*.xl*")
    i = i + 1
    Arr(i) = fname
    Do While fname <> ""
        fname = Dir
        If fname = "" Then
            Exit Do
        End If
        i = i + 1
        Arr(i) = fname
    Loop
    For p = 1 To i
        If Arr(p) <> CurFil Then
            Set CurBook = GetObject(fpath & "\" & Arr(p))
            If InStr(Arr(p), "xlsx") > 0 Then
                OtherFil = Left(Arr(1), Len(Arr(1)) - 5)
            Else
                OtherFil = Left(Arr(1), Len(Arr(1)) - 4)
            End If
            For Each osht In CurBook.Sheets
                k = 0
                For j = 1 To 10   ' ̽���һ��ǰ10�У������Ϊ�գ�����Ϊ�˹������ǿյģ����ϲ����±���
                    If Len(CurBook.Sheets(osht.Name).Cells(j, 1)) > 0 Then
                        k = 1
                    End If
                Next
                If k = 1 Then
                    ActiveShtName = Left(Arr(p), Len(Arr(1)) - 5) & "(" & osht.Name & ")"
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
        n = 0
        For m = 1 To 10
            If Len(Sheets(sht.Name).Cells(m, 1)) > 0 Then
                n = 1
            End If
        Next
        If n = 0 Then
                 sht.Delete
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
        Sheets(1).Name = "�ϲ�"
        For Each sht In CurBook.Sheets
            Imax = CurBook.Sheets(sht.Name).Cells(STRowMax, 1).End(xlUp).Row                 '��ʱ����һ�к͵�һ�л�ȡ���Χ
            Jmax = CurBook.Sheets(sht.Name).Cells(1, STColMax).End(xlToLeft).Column
            For j = 2 To CheckLine
                If Imax < CurBook.Sheets(sht.Name).Cells(STRowMax, j).End(xlUp).Row Then
                    Imax = CurBook.Sheets(sht.Name).Cells(STRowMax, j).End(xlUp).Row
                End If
            Next
            For i = 2 To CheckLine
                If Jmax < CurBook.Sheets(sht.Name).Cells(i, STColMax).End(xlToLeft).Column Then
                    Jmax = CurBook.Sheets(sht.Name).Cells(i, STColMax).End(xlToLeft).Column
                End If
            Next
            Arr = CurBook.Sheets(sht.Name).Range(CurBook.Sheets(sht.Name).Cells(1, 1), CurBook.Sheets(sht.Name).Cells(Imax, Jmax))
            IBegin = Cells(STRowMax, 1).End(xlUp).Row
            For j = 2 To CheckLine
                If IBegin < Cells(STRowMax, j).End(xlUp).Row Then
                    IBegin = Cells(STRowMax, j).End(xlUp).Row
                End If
            Next
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
        Sheets(1).Name = "�ϲ�"
        For Each sht In CurBook.Sheets
            Imax = CurBook.Sheets(sht.Name).Cells(STRowMax, 1).End(xlUp).Row
            Jmax = CurBook.Sheets(sht.Name).Cells(1, STColMax).End(xlToLeft).Column                 '��ʱ����һ�к͵�һ�л�ȡ���Χ
            For j = 2 To CheckLine
                If Imax < CurBook.Sheets(sht.Name).Cells(STRowMax, j).End(xlUp).Row Then
                    Imax = CurBook.Sheets(sht.Name).Cells(STRowMax, j).End(xlUp).Row
                End If
            Next
            For i = 2 To CheckLine
                If Jmax < CurBook.Sheets(sht.Name).Cells(i, STColMax).End(xlToLeft).Column Then
                    Jmax = CurBook.Sheets(sht.Name).Cells(i, STColMax).End(xlToLeft).Column
                End If
            Next
            Arr = CurBook.Sheets(sht.Name).Range(CurBook.Sheets(sht.Name).Cells(1, 1), CurBook.Sheets(sht.Name).Cells(Imax, Jmax))
            IBegin = Cells(1, STColMax).End(xlToLeft).Column
            For j = 2 To CheckLine
                If IBegin < Cells(j, STColMax).End(xlToLeft).Column Then
                    IBegin = Cells(j, STColMax).End(xlToLeft).Column
                End If
            Next
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
    Dim i, j, k As Integer
    Dim Imax, Jmax, Zmax As Integer
    Dim Arr As Variant
    Dim Brr(1 To CheckLine, 1 To CheckLine) As Variant
    Imax = Sheets(ActiveSheet.Name).Cells(STRowMax, 1).End(xlUp).Row
    Jmax = Sheets(ActiveSheet.Name).Cells(1, STColMax).End(xlToLeft).Column                '��ʱ����һ�к͵�һ�л�ȡ���Χ
    For j = 2 To CheckLine
        If Imax < Sheets(ActiveSheet.Name).Cells(STRowMax, j).End(xlUp).Row Then
            Imax = Sheets(ActiveSheet.Name).Cells(STRowMax, j).End(xlUp).Row
        End If
    Next
    For i = 2 To CheckLine
        If Jmax < Sheets(ActiveSheet.Name).Cells(i, STColMax).End(xlToLeft).Column Then
            Jmax = Sheets(ActiveSheet.Name).Cells(i, STColMax).End(xlToLeft).Column
        End If
    Next
    If Imax > Jmax Then
        Zmax = Imax
    Else
        Zmax = Jmax
    End If
    Arr = Range(Cells(1, 1), Cells(Zmax, Zmax))
    For i = 1 To Zmax
        For j = 1 To Zmax
            Brr(i, j) = Arr(j, i)
        Next
    Next
    Range(Cells(1, 1), Cells(Zmax, Zmax)) = Brr
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
