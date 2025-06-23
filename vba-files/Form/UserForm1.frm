VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "��ʂ�o�^"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6060
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'�O���[�o���萔
Private Const idCol As Integer = 1 'ID��
Private Const nameCol As Integer = 2 '�����̖��O��
Private Const kindsCol As Integer = 3 '��ޗ�
Private Const alcoholCol As Integer = 4 '�x����
Private Const fullCol As Integer = 5 '���J���d�ʗ�
Private Const empCol As Integer = 6 '��d�ʗ�

Option Explicit

Private Sub btnCalc_Click()
    Dim i As Long, lastRow As Long
    Dim sakeName As String
    Dim abv As Double, fullWeight As Double, emptyWeight As Double
    Dim nowWeight As Double, drankWeight As Double, pureAlcohol As Double

    Call setObj
    sakeName = cmbSake.Value

    If sakeName = "" Then
        MsgBox "������I�����Ă�������", vbExclamation
        Exit Sub
    End If

    If Not IsNumeric(txtNowWeight.Value) Then
        MsgBox "���݂̏d���𐳂������͂��Ă�������", vbExclamation
        Exit Sub
    End If

    nowWeight = CDbl(txtNowWeight.Value)

    lastRow = ws.Cells(ws.Rows.Count, nameCol).End(xlUp).Row
    For i = 2 To lastRow
        If ws.Cells(i, nameCol).Value = sakeName Then
            ' �K�v�ȏ����擾
            abv = ws.Cells(i, alcoholCol).Value         ' �x��
            fullWeight = ws.Cells(i, fullCol).Value  ' ���J���d��

            If ws.Cells(i, empCol).Value = "" Then
                MsgBox "���̎��͋�{�g���d�ʂ����o�^�ł��B" & vbCrLf & _
                       "���ݏI�������{�g���d�ʂ���͂��Ă��������B", vbExclamation
                Exit Sub
            End If

            emptyWeight = ws.Cells(i, empCol).Value ' ��{�g���d��

            ' ���̓`�F�b�N
            If nowWeight > fullWeight Or nowWeight < emptyWeight Then
                MsgBox "���݂̏d�����s���ł��B", vbExclamation
                Exit Sub
            End If

            ' ���񂾏d�ʌv�Z
            drankWeight = fullWeight - nowWeight

            ' ���A���R�[���ʌv�Z�i�A���R�[���̔�d = 0.8�j
            pureAlcohol = drankWeight * (abv / 100) * 0.8

            ' ���ʂ�\���i�����_1���j
            lblResult.Caption = "���A���R�[����: " & Format(pureAlcohol, "0.0") & " g"
            Exit Sub
        End If
    Next i
    Call releaseObj
End Sub

Private Sub cmbSake_Change()
    Dim targetRow As Long
    Dim lastRow As Long
    Dim i As Long
    Call setObj

    lastRow = ws.Cells(ws.Rows.Count, nameCol).End(xlUp).Row

    For i = 2 To lastRow
        If ws.Cells(i, idCol).Value & "." & ws.Cells(i, nameCol).Value = cmbSake.Value Then
            ' �Ώۂ̍s����������������擾
            lblABV.Caption = "�x���F" & ws.Cells(i, alcoholCol).Value & " %"
            lblFullW.Caption = "���J���d�ʁF" & ws.Cells(i, fullCol).Value & " g"
            
            If ws.Cells(i, empCol).Value = "" Then
                lblEmptyW.Caption = "��{�g���d�ʁF���o�^"
                lblAlert.Caption = "!!!���ݏI��������{�g���d�ʂ���͂��Ă�������!!!"
            Else
                lblEmptyW.Caption = "��{�g���d�ʁF" & ws.Cells(i, empCol).Value & " g"
                lblAlert.Caption = ""
            End If
            
            Exit For
        End If
    Next i
    Call releaseObj
End Sub

Private Sub UserForm_Initialize()
    '�ϐ��錾
    Dim i As Integer
    Dim lastCell As Range
    Dim dataRange As Range
    Call setObj
    'ws.Rows.Count �ŃV�[�g�̍ő�s�����擾���A�������� End(xlUp) �Ńf�[�^�̂���ŏI�Z����T���B
    Set lastCell = ws.Cells(ws.Rows.Count, nameCol).End(xlUp)

    '�f�[�^��2�s�ڈȍ~�ɑ��݂���ꍇ�̂ݏ��������s
    If lastCell.Row >= 2 Then
        '2�s�ڂ���ŏI�s�܂ł͈̔͂��擾
        'Set dataRange = ws.Range(Cells(2, nameCol), lastCell)
        For i = 2 To lastCell.Row
            cmbSake.AddItem ws.Cells(i, idCol).Value & "." & ws.Cells(i, nameCol).Value
        Next i
        'ComboBox�̃��X�g�ɔ͈͂̒l��ݒ�
        'Me.cmbSake.List = dataRange.Value
    End If
    Call releaseObj
End Sub
