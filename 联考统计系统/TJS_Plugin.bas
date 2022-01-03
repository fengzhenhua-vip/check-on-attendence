Attribute VB_Name = "TJS_Plugin"
Sub 非复习部上线人数()
Dim i, j, temp As Integer
Dim Imax, Jmax As Integer
Dim YiNum, ErNum As Integer
Imax = Sheets("CJ").Cells(90000, 1).End(xlUp).Row
Jmax = Sheets("CJ").Cells(1, 200).End(xlToLeft).Column
For i = 2 To Imax
    If InStr(Sheets("CJ").Cells(i, 5), "平原一中") > 0 And IsNumeric(Sheets("CJ").Cells(i, 4)) Then
        If Sheets("CJ").Cells(i, 4) < 43 Then
        temp = temp + 1
        End If
    End If
Next
Sheets("去除复习部").Cells(4, 2) = temp
'
Dim YiXian, ErXian As Single
YiXian = 39: ErXian = 71.5: YiNum = Empty: ErNum = Empty
For i = 2 To Imax
    If InStr(Sheets("CJ").Cells(i, 5), "平原一中") > 0 And IsNumeric(Sheets("CJ").Cells(i, 4)) Then
        If Sheets("CJ").Cells(i, 4) < 43 And Sheets("CJ").Cells(i, 38) >= YiXian Then
            YiNum = YiNum + 1
        End If
        If Sheets("CJ").Cells(i, 4) < 43 And Sheets("CJ").Cells(i, 34) >= ErXian Then
            ErNum = ErNum + 1
        End If
    End If
Next
Sheets("去除复习部").Cells(15, 21) = YiNum
'Sheets("去除复习部").Cells(9, 19) = ErNum
End Sub
Sub 非复人数()
Dim i, j, temp As Integer
Dim Imax, Jmax As Integer
Dim YiNum, ErNum As Integer
Imax = Sheets("CJ").Cells(90000, 1).End(xlUp).Row
Jmax = Sheets("CJ").Cells(1, 200).End(xlToLeft).Column
For i = 2 To Imax
    If InStr(Sheets("CJ").Cells(i, 5), "平原一中") > 0 And IsNumeric(Sheets("CJ").Cells(i, 4)) Then
        If 0 < Sheets("CJ").Cells(i, 4) And Sheets("CJ").Cells(i, 4) < 40 Then
        temp = temp + 1
        End If
    End If
Next
Sheets("去除复习部").Cells(4, 2) = temp
End Sub
Sub 统计总人数X(hang)
Dim i, j, temp As Integer
Dim Imax, Jmax As Integer
j = hang
Imax = Sheets("山东").Cells(90000, 1).End(xlUp).Row
Jmax = Sheets("山东").Cells(1, 200).End(xlToLeft).Column
temp = Empty
For i = 3 To Imax
    If InStr(Sheets("山东").Cells(i, 2), Sheets("总分").Cells(j, 1)) > 0 Then
         temp = temp + 1
    End If
Next
Sheets("总分").Cells(j, 2) = temp
End Sub
Sub 统计总人数()
    For q = 4 To 12
        Call 统计总人数X(q)
    Next
End Sub

Sub 过滤学校X(Xian)
Dim Imax, Jmax As Integer
Imax = Sheets(Xian).Cells(90000, 1).End(xlUp).Row
Jmax = Sheets(Xian).Cells(1, 200).End(xlToLeft).Column
Dim i, j, k, p As Integer
Dim gltab(1 To 30, 1 To 22) As Variant
p = 0: Erase gltab
p = p + 1
For k = 1 To 22
    gltab(p, k) = Sheets(Xian).Cells(3, k)
Next
For i = 4 To Imax
    For j = 4 To 12
        If InStr(Sheets(Xian).Cells(i, 1), Sheets("总分").Cells(j, 1)) > 0 Then
            p = p + 1
            For k = 1 To 22
                gltab(p, k) = Sheets(Xian).Cells(i, k)
            Next
        End If
    Next
Next
For j = 2 To 22
    If InStr(gltab(1, j), "名次") = 0 Then
        For i = 2 To p
            gltab(p + 1, j) = gltab(p + 1, j) + CInt(gltab(i, j))
        Next
    End If
Next
gltab(p + 1, 1) = "合计"
Sheets(Xian).Range(Sheets(Xian).Cells(3, 1), Sheets(Xian).Cells(30, 22)) = gltab
End Sub
Sub 过滤学校()
Call 过滤学校X("一线")
Call 过滤学校X("二线")
End Sub
