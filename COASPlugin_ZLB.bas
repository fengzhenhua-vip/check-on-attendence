Attribute VB_Name = "COASPlugin_ZLB"
Sub COASPlugin_ZLB()
Dim ZLBPath, ZLBDir, ZLBFile, ZLBName, ZLBFolder, COACfgPath As String
Dim ZLBYue, ZLBZhou, ZLBRi As Integer
Dim ZLBBook As Workbook
Dim i, j As Integer
Dim ZLBRmax, ZLBCmax, ZLBStep As Integer
Dim ZLBScore As Integer
ZLBScore = 20
Dim ZLBSource As Variant
COACfgPath = CreateObject("WScript.Shell").specialfolders("Desktop") & "\考勤系统"                          '默认设置为桌面
ZLBPath = ActiveWorkbook.Path
'
Set SFO = CreateObject("Scripting.FileSystemObject")
If InStr(ActiveWorkbook.Name, "异常班主任") > 0 Then
    Application.ScreenUpdating = False
    ZLBRmax = Cells(1000, 3).End(xlUp).Row
    ZLBCmax = Cells(1, 26).End(xlToLeft).Column
    ZLBSource = Sheets(1).Range(Sheets(1).Cells(1, 1), Sheets(1).Cells(ZLBRmax, ZLBCmax))
    ZLBStep = 2
    Do While Len(ZLBSource(ZLBStep, 16)) = 0
        ZLBStep = ZLBStep + 1
    Loop
    ZLBStep = ZLBStep - 1
'基本信息设置
    ZLBYue = CInt(Mid(ActiveWorkbook.Name, 24, 2))
    ZLBRi = CInt(Mid(ActiveWorkbook.Name, 27, 2))
    ZLBZhou = Fix(ZLBRi / 7)
    ZLBName = "高三文理部" & ZLBYue & "月份第" & ZLBZhou & "周量化"
    ZLBFile = ZLBPath & "\" & ZLBName & ".xlsx"
    If SFO.fileExists(ZLBFile) = False Then
        FileCopy COACfgPath & "\考勤系统配置\周量化模板.xlsx", ZLBFile
    End If
    Workbooks.Open Filename:=ZLBFile
    Set ZLBBook = GetObject(ZLBFile)
    ZLBBook.Sheets(1).Cells(1, 1) = ZLBName
' 加入班主任打分
    For i = 3 To 38
        For j = 2 To ZLBRmax - ZLBStep
            If Len(ZLBSource(j, 1)) > 0 And InStr(ZLBBook.Sheets(1).Cells(i, 2), ZLBSource(j, 1)) > 0 Then
                ZLBBook.Sheets(1).Cells(i, 3) = ZLBScore - ZLBSource(j + ZLBStep, 16)
            End If
        Next
    Next
    For i = 3 To 38
        If Len(ZLBBook.Sheets(1).Cells(i, 3)) = 0 Then
            ZLBBook.Sheets(1).Cells(i, 3) = ZLBScore
        End If
    Next
'加入班组评分
    Dim fpath, fName As String
    Dim fBook As Workbook
    Dim fCol As Integer
    fpath = ActiveWorkbook.Path & "\周量化报表打印"
'录入课堂加分
    fName = Dir(fpath & "\班级*")
    Set fBook = GetObject(fpath & "\" & fName)
    fCol = fBook.Sheets(1).Cells(1, 100).End(xlToLeft).Column
    For i = 3 To 38
        ZLBBook.Sheets(1).Cells(i, 8) = fBook.Sheets(1).Cells(i, fCol)
    Next
' 录入卫生
    fName = Dir(fpath & "\高三卫生*")
    Set fBook = GetObject(fpath & "\" & fName)
    fCol = fBook.Sheets(1).Cells(4, 100).End(xlToLeft).Column
    For i = 3 To 38
        ZLBBook.Sheets(1).Cells(i, 4) = fBook.Sheets(1).Cells(i + 2, fCol)
    Next
' 录入激情诵读
    fName = Dir(fpath & "\激情*")
    Set fBook = GetObject(fpath & "\" & fName)
    fCol = fBook.Sheets(1).Cells(2, 100).End(xlToLeft).Column
    For i = 3 To 38
        ZLBBook.Sheets(1).Cells(i, 6) = fBook.Sheets(1).Cells(i, fCol)
    Next
' 录入作业展
    fName = Dir(fpath & "\作业展*")
    Set fBook = GetObject(fpath & "\" & fName)
    fCol = fBook.Sheets(1).Cells(3, 100).End(xlToLeft).Column
    For i = 3 To 38
        ZLBBook.Sheets(1).Cells(i, 5) = fBook.Sheets(1).Cells(i + 1, fCol)
    Next
    Application.ScreenUpdating = True
 ' 关闭及退出
     Application.DisplayAlerts = False                                                                                          '关闭所有工作薄
      ZLBBook.Close savechanges:=True
      Workbooks.Close
     Application.DisplayAlerts = True
     Application.Quit
     Shell "explorer.exe " & ZLBPath, vbNormalFocus
End If
End Sub

