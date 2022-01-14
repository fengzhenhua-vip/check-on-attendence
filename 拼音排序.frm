VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 拼音排序 
   Caption         =   "拼音排序V1.0"
   ClientHeight    =   1710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3135
   OleObjectBlob   =   "拼音排序.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "拼音排序"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Click()

End Sub

Private Sub 排序列_Change()

End Sub

Private Sub 起始行_Change()

End Sub

Private Sub 输入排序列_Click()

End Sub

Private Sub 输入起始行_Click()

End Sub

Private Sub 退出_Click()
    End
End Sub

Private Sub 执行拼音排序_Click()
Dim i, j As Integer
    If 起始行.Value = "" Then
      MsgBox "请输入起始行 ！"
    End If
    If 排序列.Value = "" Then
      MsgBox "请输入排序列 ！"
    End If
    If 起始行.Value <> "" And 排序列.Value <> "" Then
        i = CInt(起始行.Value)
        j = CInt(排序列.Value)
        Call 标准VBA拼音排序(ActiveSheet.Name, i, j)
    End If
End Sub
