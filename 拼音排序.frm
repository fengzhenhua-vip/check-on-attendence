VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ƴ������ 
   Caption         =   "ƴ������V1.0"
   ClientHeight    =   1710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3135
   OleObjectBlob   =   "ƴ������.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "ƴ������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Click()

End Sub

Private Sub ������_Change()

End Sub

Private Sub ��ʼ��_Change()

End Sub

Private Sub ����������_Click()

End Sub

Private Sub ������ʼ��_Click()

End Sub

Private Sub �˳�_Click()
    End
End Sub

Private Sub ִ��ƴ������_Click()
Dim i, j As Integer
    If ��ʼ��.Value = "" Then
      MsgBox "��������ʼ�� ��"
    End If
    If ������.Value = "" Then
      MsgBox "������������ ��"
    End If
    If ��ʼ��.Value <> "" And ������.Value <> "" Then
        i = CInt(��ʼ��.Value)
        j = CInt(������.Value)
        Call ��׼VBAƴ������(ActiveSheet.Name, i, j)
    End If
End Sub
