Attribute VB_Name = "Module1"
'�O���[�o���ϐ�
Public ws As Worksheet

Option Explicit

Sub ShowUserForm()

    UserForm1.Show




End Sub

'�I�u�W�F�N�g�ϐ����
Public Sub setObj()
    Set ws = ThisWorkbook.Worksheets("�����}�X�^")
End Sub
'�I�u�W�F�N�g�ϐ��J��
Public Sub releaseObj()
    Set ws = Nothing
End Sub
