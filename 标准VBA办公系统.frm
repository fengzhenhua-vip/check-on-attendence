VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 标准VBA办公系统 
   Caption         =   "标准VBA工具箱V1.0"
   ClientHeight    =   1950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4815
   OleObjectBlob   =   "标准VBA办公系统.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "标准VBA办公系统"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'名称：标准VBA办公工具箱
'版本：V1.0
'作者：冯振华
'日期：2022年1月14日22:54


Private Sub 标准填表_Click()
 标准填表X.Show
End Sub

Private Sub 拆工作薄_Click()
    标准拆分薄X.Show
End Sub

Private Sub 单列排序_Click()
 标准排序X.Show
End Sub

Private Sub 合工作薄_Click()
 标准合并薄X.Show
End Sub

Private Sub 合工作表_Click()
 标准合并表X.Show
End Sub

Private Sub 交换行列_Click()
 Call 标准VBA转置
End Sub

Private Sub 人数统计_Click()
 人数统计X.Show
End Sub

Private Sub 生成排名_Click()
  排名程序X.Show
End Sub

Private Sub 提取数据_Click()
 提取学校X.Show
End Sub
