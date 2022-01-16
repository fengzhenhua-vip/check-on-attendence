VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 标准拆分薄X 
   Caption         =   "拆分工作薄V1.0"
   ClientHeight    =   1455
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2790
   OleObjectBlob   =   "标准拆分薄X.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "标准拆分薄X"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub 执行拆分_Click()
    Dim HouZhui As String
    If InStr(ActiveWorkbook.Name, "PERSONAL.XLSB") > 0 Then
       GoTo zuihou:
    End If
    HouZhui = 后辍.Value
    Call 标准VBA拆分工作薄(HouZhui)
zuihou:
End Sub

