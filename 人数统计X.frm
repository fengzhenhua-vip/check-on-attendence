VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ����ͳ��X 
   Caption         =   "����ͳ��"
   ClientHeight    =   3225
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5640
   OleObjectBlob   =   "����ͳ��X.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "����ͳ��X"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ִ��ͳ��_Click()
    Dim i, j, k, m, n, p, q, r, s As Integer
    Dim InName, OutName As String
    If ���뱨����.Value = "" Then
      MsgBox "����д���뱨���� ��"
    End If
    If ������ʼ��.Value = "" Then
      MsgBox "����д������ʼ�� ��"
    End If
    If ����������.Value = "" Then
      MsgBox "������д�������� ��"
    End If
    If ����������.Value = "" Then
      MsgBox "����д���������� ��"
    End If
    If ���������.Value = "" Then
      MsgBox "����д��������� ��"
    End If
    If �����ʼ��.Value = "" Then
      MsgBox "����д�����ʼ�� ��"
    End If
    If ���������.Value = "" Then
      MsgBox "������д�������� ��"
    End If
    If ���������.Value = "" Then
      MsgBox "����д��������� ��"
    End If
    If ��ͷ�.Value = "" Then
      MsgBox "����д��ͷ� ��"
    End If
    If ��߷�.Value = "" Then
      MsgBox "����д��߷� ��"
    End If
    If ���뱨����.Value <> "" And ������ʼ��.Value <> "" And ����������.Value <> "" And ����������.Value <> "" Then
        If ���������.Value <> "" And �����ʼ��.Value <> "" And ���������.Value <> "" And ���������.Value <> "" Then
            InName = ���뱨����.Value: OutName = ���������.Value
            i = CInt(������ʼ��.Value): m = CInt(�����ʼ��.Value)
            j = CInt(����������.Value): n = CInt(���������.Value)
            k = CInt(����������.Value): p = CInt(���������.Value)
            r = CDbl(��ͷ�.Value): s = CDbl(��߷�.Value)
            Call ��׼VBA����������(InName, i, j, k, OutName, m, n, p, r, s)
        End If
    End If
End Sub
