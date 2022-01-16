VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 提取学校X 
   Caption         =   "提取数据"
   ClientHeight    =   2070
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5445
   OleObjectBlob   =   "提取学校X.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "提取学校X"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub 执行提取_Click()
    Dim i, j, k, m, n, p As Integer
    If 起始行I.Value = "" Then
      MsgBox "请填写输入起始行 ！"
    End If
    If 提取列I.Value = "" Then
      MsgBox "请填写输入提取列 ！"
    End If
    If 起始行O.Value = "" Then
      MsgBox "请填写输出起始行 ！"
    End If
    If 提取列O.Value = "" Then
      MsgBox "请填写输出提取列 ！"
    End If
    If 起始行I.Value <> "" And 起始行O.Value <> "" And 提取列I.Value <> "" And 提取列O.Value <> "" Then
        i = CInt(起始行I.Value)
        j = CInt(提取列I.Value)
        m = CInt(起始行O.Value)
        n = CInt(提取列O.Value)
        Call 标准VBA提取(报表名I.Value, i, j, 报表名O.Value, m, n)
        Sheets(报表名O.Value).Select
    End If
End Sub
