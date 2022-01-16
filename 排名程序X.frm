VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 排名程序X 
   Caption         =   "排名程序V1.0"
   ClientHeight    =   1935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3270
   OleObjectBlob   =   "排名程序X.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "排名程序X"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub 名次列_Change()

End Sub

Private Sub 排名列_Change()

End Sub

Private Sub 起始行_Change()

End Sub

Private Sub 退出_Click()
    End
End Sub

Private Sub 输入名次列_Click()

End Sub

Private Sub 执行排名_Click()
    Dim i, j, k As Integer
    If 起始行.Value = "" Then
      MsgBox "请输入起始行 ！"
    End If
    If 排名列.Value = "" Then
      MsgBox "请输入排名列 ！"
    End If
    If 名次列.Value = "" Then
      MsgBox "请输入名次列 ！"
    End If
    If 起始行.Value <> "" And 排名列.Value <> "" And 名次列.Value <> "" Then
        i = CInt(起始行.Value)
        j = CInt(排名列.Value)
        k = CInt(名次列.Value)
        Call 标准VBA排名(ActiveSheet.Name, i, j, k)
    End If
End Sub
