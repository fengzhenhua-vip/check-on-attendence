Attribute VB_Name = "TJS_V10"
Sub 执行统计()
    Call 统计分数线
'
    Dim sht As Worksheet
    Dim OutOk As Integer
    Application.DisplayAlerts = False
    For Each sht In ActiveWorkbook.Sheets
        Select Case sht.Name
            Case Is = "总分"
                OutOk = 1
            Case Is = "语文"
                OutOk = 1
            Case Is = "数学"
                OutOk = 1
            Case Is = "英语"
                OutOk = 1
            Case Is = "物理"
                OutOk = 1
            Case Is = "化学"
                OutOk = 1
            Case Is = "生物"
                OutOk = 1
            Case Is = "政治"
                OutOk = 1
            Case Is = "历史"
                OutOk = 1
            Case Is = "地理"
                OutOk = 1
            Case Else
                OutOk = 0
        End Select
        If OutOk = 1 Then
            sht.Copy
            ActiveWorkbook.SaveAs Filename:=OutFolder & "\" & sht.Name & "联考上线对比表", FileFormat:=xlNormal  '将工作簿另存为EXCEL默认格式
            ActiveWorkbook.Close
        End If
    Next
    Application.DisplayAlerts = True
        MsgBox "联考成绩统计完毕!"
End Sub


Sub 导入上次数据X(Yuan, KeMu, KeMuLie)
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
Sub 导入上次数据XX(XueKe)
    Sheets(XueKe).Cells(1, 1) = XueKe & "统计对比表"
    Call 导入上次数据X("一线", XueKe, 8)
    Call 导入上次数据X("二线", XueKe, 3)
End Sub

Sub 导入上次数据()
    Call 导入上次数据XX("总分")
    Call 导入上次数据XX("语文")
    Call 导入上次数据XX("数学")
    Call 导入上次数据XX("英语")
    Call 导入上次数据XX("物理")
    Call 导入上次数据XX("化学")
    Call 导入上次数据XX("生物")
    Call 导入上次数据XX("政治")
    Call 导入上次数据XX("历史")
    Call 导入上次数据XX("地理")
End Sub
