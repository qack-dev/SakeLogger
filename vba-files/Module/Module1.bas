Attribute VB_Name = "Module1"
'�O���[�o���ϐ�
Public wsMaster As Worksheet
Public wsLog As Worksheet
Public lastCell As Range


'�O���[�o���萔
'�����}�X�^�V�[�g
Public Const idCol As Integer = 1 'ID��
Public Const nameCol As Integer = 2 '�����̖��O��
Public Const kindsCol As Integer = 3 '��ޗ�
Public Const alcoholCol As Integer = 4 '�x����
Public Const fullCol As Integer = 5 '���J���d�ʗ�
Public Const empCol As Integer = 6 '��d�ʗ�
'�����L�^�V�[�g
Public Const logDateCol As Integer = 1 '������
Public Const logNameCol As Integer = 2 '�����̖��O��
Public Const logNowCol As Integer = 3 '���ݏd�ʗ�
Public Const logPureAlcCol As Integer = 4 '���A���ʗ�
Public Const logDrunkCol As Integer = 5 '���񂾗ʗ�
Public Const logComCol As Integer = 6 '�R�����g��
Public Const logIdCol As Integer = 7 '���̃V�[�g��ID�̗�
Public Const logNameIdCol As Integer = 8 '������ID��


Option Explicit

'�I�u�W�F�N�g�ϐ����
Public Sub setObj()
    Set wsMaster = ThisWorkbook.Worksheets("�����}�X�^")
    Set wsLog = Sheets("�����L�^")
    'wsMaster.RowsMaster.Count �ŃV�[�g�̍ő�s�����擾���A�������� End(xlUp) �Ńf�[�^�̂���ŏI�Z����T���B
    Set lastCell = wsMaster.Cells(wsMaster.Rows.Count, nameCol).End(xlUp)
End Sub
'�I�u�W�F�N�g�ϐ��J��
Public Sub releaseObj()
    Set wsMaster = Nothing
    Set lastCell = Nothing
    Set wsLog = Nothing
End Sub

Sub ShowUserForm()

    UserForm1.Show




End Sub

