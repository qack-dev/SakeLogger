Attribute VB_Name = "Module1"
'�O���[�o���ϐ�
Public ws As Worksheet
Public lastCell As Range


'�O���[�o���萔
Public Const idCol As Integer = 1 'ID��
Public Const nameCol As Integer = 2 '�����̖��O��
Public Const kindsCol As Integer = 3 '��ޗ�
Public Const alcoholCol As Integer = 4 '�x����
Public Const fullCol As Integer = 5 '���J���d�ʗ�
Public Const empCol As Integer = 6 '��d�ʗ�

Option Explicit

Sub ShowUserForm()

    UserForm1.Show




End Sub

'�I�u�W�F�N�g�ϐ����
Public Sub setObj()
    Set ws = ThisWorkbook.Worksheets("�����}�X�^")
    'ws.Rows.Count �ŃV�[�g�̍ő�s�����擾���A�������� End(xlUp) �Ńf�[�^�̂���ŏI�Z����T���B
    Set lastCell = ws.Cells(ws.Rows.Count, nameCol).End(xlUp)
End Sub
'�I�u�W�F�N�g�ϐ��J��
Public Sub releaseObj()
    Set ws = Nothing
    Set lastCell = Nothing
End Sub
