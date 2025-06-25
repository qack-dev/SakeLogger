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

Option Explicit

'�I�u�W�F�N�g�ϐ����
Public Sub setObj()
    Set wsMaster = ThisWorkbook.Worksheets("�����}�X�^")
    Set wsLog = ThisWorkbook.Worksheets("�����L�^")
    'wsMaster.RowsMaster.Count �ŃV�[�g�̍ő�s�����擾���A�������� End(xlUp) �Ńf�[�^�̂���ŏI�Z����T���B
    Set lastCell = wsMaster.Cells(wsMaster.Rows.Count, nameCol).End(xlUp)
End Sub
'�I�u�W�F�N�g�ϐ��J��
Public Sub releaseObj()
    Set wsMaster = Nothing
    Set wsLog = Nothing
    Set lastCell = Nothing
End Sub

Public Sub ShowUserForm()
    Call setObj
    wsMaster.Activate
    frmSakeLogger.Show
    Call releaseObj



End Sub

'���K�\����'yyyy/mm/dd'�`�����`�F�b�N���A�����t�Ƃ��đÓ�������
Public Function IsYyyyMmDdFormat_RegEx(ByVal target As String) As Boolean
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")

    ' �p�^�[����ݒ�
    ' ^         : ������̐擪
    ' \d{4}     : 4���̐��� (�N)
    ' /         : ��؂蕶���̃X���b�V��
    ' \d{2}     : 2���̐��� (��)
    ' /         : ��؂蕶���̃X���b�V��
    ' \d{2}     : 2���̐��� (��)
    ' $         : ������̖���
    regEx.Pattern = "^\d{4}/\d{2}/\d{2}$"

    If regEx.Test(target) And IsDate(target) Then
        IsYyyyMmDdFormat_RegEx = True
    Else
        IsYyyyMmDdFormat_RegEx = False
    End If
End Function

'���񂾗�(drankWeight)�Ə��A���R�[����(pureAlcohol)�̌v�Z
Public Function CalcAlcoholInfo(sakeName As String, nowWeight As Double, _
                         ByRef drankWeight As Double, ByRef pureAlcohol As Double) As Boolean
    Dim abv As Double, fullWeight As Double, emptyWeight As Double
    Dim prevWeight As Double
    Dim i As Long, lastRow As Long
    Dim found As Boolean

    On Error GoTo ErrHandler
    
    ' �}�X�^����x��(ABV)�E�d�ʎ擾
    found = False
    For i = 2 To lastCell.Row
        If wsMaster.Cells(i, idCol).Value & "." & wsMaster.Cells(i, nameCol).Value = sakeName Then
            abv = wsMaster.Cells(i, alcoholCol).Value     ' �x��
            fullWeight = wsMaster.Cells(i, fullCol).Value ' ���J���d��
            If wsMaster.Cells(i, empCol).Value = "" Then
                MsgBox "���̎��͋�{�g���d�ʂ����o�^�ł��B" & vbCrLf & _
                       "���ݏI�������{�g���d�ʂ���͂��Ă��������B", vbExclamation
                CalcAlcoholInfo = False
            Else
                emptyWeight = wsMaster.Cells(i, empCol).Value ' ��{�g���d��
                ' ���̓`�F�b�N
                If nowWeight > fullWeight Or nowWeight < emptyWeight Then
                    MsgBox "���݂̏d�����s���ł��B", vbExclamation
                    Exit Function
                End If
            End If
            found = True
            Exit For
        End If
    Next i

    If Not found Then
        MsgBox "�����}�X�^�ɂ��̂�����������܂���B", vbCritical
        CalcAlcoholInfo = False
        Exit Function
    End If

    ' ���򏈗��iOptionButton�̏�ԂŁj
    If frmSakeLogger.optNewOpen.Value = True Then
        ' �V�i�J��
        drankWeight = fullWeight - nowWeight

    ElseIf frmSakeLogger.optContinued.Value = True Then
        ' �p�����p�F�O��d�ʎ擾
        lastRow = wsLog.Cells(wsLog.Rows.Count, logIdCol).End(xlUp).Row
        found = False
        For i = lastRow To 2 Step -1
            If wsLog.Cells(i, logNameCol).Value = sakeName Then
                prevWeight = wsLog.Cells(i, logNowCol).Value
                found = True
                Exit For
            End If
        Next i

        If Not found Then
            MsgBox "���̂����̋L�^���܂����݂��܂���B" & vbCrLf & _
                   "�w�V�i���J�����x��I��ł��������B", vbExclamation
            CalcAlcoholInfo = False
            Exit Function
        End If

        drankWeight = prevWeight - nowWeight

    Else
        MsgBox "�w�V�i���J�����x�܂��́w�r���̂��������񂾁x��I��ł��������B", vbExclamation
        CalcAlcoholInfo = False
        Exit Function
    End If

    ' ���A���R�[���ʌv�Z
    pureAlcohol = drankWeight * (abv / 100) * 0.8

    CalcAlcoholInfo = True
    Exit Function

ErrHandler:
    MsgBox "�v�Z���ɃG���[���������܂����B" & vbCrLf & Err.Description, vbCritical
    CalcAlcoholInfo = False
End Function


