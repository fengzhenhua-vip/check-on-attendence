VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ��ȡѧУX 
   Caption         =   "��ȡ����"
   ClientHeight    =   2070
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5445
   OleObjectBlob   =   "��ȡѧУX.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "��ȡѧУX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ִ����ȡ_Click()
    Dim i, j, k, m, n, p As Integer
    If ��ʼ��I.Value = "" Then
      MsgBox "����д������ʼ�� ��"
    End If
    If ��ȡ��I.Value = "" Then
      MsgBox "����д������ȡ�� ��"
    End If
    If ��ʼ��O.Value = "" Then
      MsgBox "����д�����ʼ�� ��"
    End If
    If ��ȡ��O.Value = "" Then
      MsgBox "����д�����ȡ�� ��"
    End If
    If ��ʼ��I.Value <> "" And ��ʼ��O.Value <> "" And ��ȡ��I.Value <> "" And ��ȡ��O.Value <> "" Then
        i = CInt(��ʼ��I.Value)
        j = CInt(��ȡ��I.Value)
        m = CInt(��ʼ��O.Value)
        n = CInt(��ȡ��O.Value)
        Call ��׼VBA��ȡ(������I.Value, i, j, ������O.Value, m, n)
        Sheets(������O.Value).Select
    End If
End Sub
