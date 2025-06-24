VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSakeLogger 
   Caption         =   "��ʂ�o�^"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7530
   OleObjectBlob   =   "frmSakeLogger.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmSakeLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCalc_Click()
    Dim i As Long, lastRow As Long
    Dim sakeName As String
    Dim abv As Double, fullWeight As Double, emptyWeight As Double
    Dim nowWeight As Double, drankWeight As Double, pureAlcohol As Double

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

    lastRow = lastCell.Row
    For i = 2 To lastRow
        If wsMaster.Cells(i, idCol).Value & "." & wsMaster.Cells(i, nameCol).Value = sakeName Then
            ' �K�v�ȏ����擾
            abv = wsMaster.Cells(i, alcoholCol).Value         ' �x��
            fullWeight = wsMaster.Cells(i, fullCol).Value  ' ���J���d��

            If wsMaster.Cells(i, empCol).Value = "" Then
                MsgBox "���̎��͋�{�g���d�ʂ����o�^�ł��B" & vbCrLf & _
                       "���ݏI�������{�g���d�ʂ���͂��Ă��������B", vbExclamation
            Else
                emptyWeight = wsMaster.Cells(i, empCol).Value ' ��{�g���d��
                ' ���̓`�F�b�N
                If nowWeight > fullWeight Or nowWeight < emptyWeight Then
                    MsgBox "���݂̏d�����s���ł��B", vbExclamation
                    Exit Sub
                End If
            End If
            ' ���񂾏d�ʌv�Z
            drankWeight = fullWeight - nowWeight
            ' ���A���R�[���ʌv�Z�i�A���R�[���̔�d = 0.8�j
            pureAlcohol = drankWeight * (abv / 100) * 0.8
            ' ���ʂ�\���i�����_1���j
            lblResult.Caption = "���A���R�[���ʁF" & Format(pureAlcohol, "0.0") & " g"
            Exit Sub
        End If
    Next i
End Sub

Private Sub cmbSake_Change()
    Dim targetRow As Long
    Dim lastRow As Long
    Dim i As Long
    
    lastRow = lastCell.Row
    
    With wsMaster
        For i = 2 To lastRow
            If .Cells(i, idCol).Value & "." & .Cells(i, nameCol).Value = cmbSake.Value Then
                ' �Ώۂ̍s����������������擾
                lblABV.Caption = "�x���F" & .Cells(i, alcoholCol).Value & " %"
                lblFullW.Caption = "���J���d�ʁF" & .Cells(i, fullCol).Value & " g"
                
                If .Cells(i, empCol).Value = "" Then
                    lblEmptyW.Caption = "��{�g���d�ʁF���o�^"
                    lblAlert.Caption = "!!!���ݏI��������{�g���d�ʂ���͂��Ă�������!!!"
                Else
                    lblEmptyW.Caption = "��{�g���d�ʁF" & .Cells(i, empCol).Value & " g"
                    lblAlert.Caption = ""
                End If
                
                Exit For
            End If
        Next i
    End With
End Sub

Private Sub btnSave_Click()
    Dim sakeName As String
    Dim abv As Double, fullWeight As Double, emptyWeight As Double
    Dim nowWeight As Double, drankWeight As Double, pureAlcohol As Double
    Dim i As Long
    Dim iLog As Long
    Dim lastLogRow As Long
    Dim prevWeight As Double
    Dim found As Boolean

    ' --- ���̓`�F�b�N ---
    If cmbSake.Value = "" Then
        MsgBox "����I��ł�������", vbExclamation
        Exit Sub
    End If

    If Not IsYyyyMmDdFormat_RegEx(txtDate.Value) Then
        MsgBox "���񂾓��ɂ�'yyyy/mm/dd'�`���œ��͂��Ă�������", vbExclamation
        Exit Sub
    End If

    If Not IsNumeric(txtNowWeight.Value) Then
        MsgBox "���݂̏d������͂��Ă�������", vbExclamation
        Exit Sub
    End If

    sakeName = cmbSake.Value
    nowWeight = CDbl(txtNowWeight.Value)

    ' --- �}�X�^������擾 ---
    For i = 2 To lastCell.Row
        If wsMaster.Cells(i, idCol).Value & "." & wsMaster.Cells(i, nameCol).Value = sakeName Then
            abv = wsMaster.Cells(i, alcoholCol).Value
            fullWeight = wsMaster.Cells(i, fullCol).Value
            If wsMaster.Cells(i, empCol).Value = "" Then
                MsgBox "��{�g���d�ʂ������͂ł�", vbExclamation
            Else
                emptyWeight = wsMaster.Cells(i, empCol).Value ' ��{�g���d��
                ' ���̓`�F�b�N
                If nowWeight > fullWeight Or nowWeight < emptyWeight Then
                    MsgBox "���݂̏d�����s���ł��B", vbExclamation
                    Exit Sub
                End If
            End If
        End If
    Next i

    ' --- ���񂾗ʁE���A���ʂ̌v�Z���@�𕪊� ---
    If optNewOpen.Value = True Then
        ' �V�i���J�����ꍇ�F���J���d�ʂ���v�Z
        drankWeight = fullWeight - nowWeight
    ElseIf optContinued.Value = True Then
        ' �p�����p�F�ߋ��̋L�^���璼�߂̏d�ʂ�����
        found = False
        lastLogRow = wsLog.Cells(wsLog.Rows.Count, logIdCol).End(xlUp).Row

        ' �������ɑk���ē������̒��߂̏d�ʂ�T��
        For iLog = lastLogRow To 2 Step -1
            If wsLog.Cells(iLog, logNameCol).Value = sakeName Then
                prevWeight = wsLog.Cells(iLog, logNowCol).Value
                found = True
                Exit For
            End If
        Next iLog

        If Not found Then
            MsgBox "���̂����̋L�^���܂����݂��܂���B" & vbCrLf & _
                   "�w�V�i���J�����x��I��ł��������B", vbExclamation
            Exit Sub
        End If

        drankWeight = prevWeight - nowWeight
    Else
        MsgBox "�V�i���p������I��ł��������B", vbExclamation
        Exit Sub
    End If

    ' --- ���A���R�[���ʂ̌v�Z�i���ʁj ---
    pureAlcohol = drankWeight * (abv / 100) * 0.8

    ' --- �����ݒ�̕ύX ---
    wsLog.Cells(lastLogRow + 1, logNowCol).NumberFormat = "0.0"
    wsLog.Cells(lastLogRow + 1, logPureAlcCol).NumberFormat = "0.0"
    wsLog.Cells(lastLogRow + 1, logDrunkCol).NumberFormat = "0.0"

    ' --- ���O�ɋL�^���� ---
    wsLog.Cells(lastLogRow + 1, logDateCol).Value = txtDate.Value          ' ����
    wsLog.Cells(lastLogRow + 1, logNameCol).Value = sakeName               ' ��
    wsLog.Cells(lastLogRow + 1, logNowCol).Value = nowWeight              ' ���ݏd��
    wsLog.Cells(lastLogRow + 1, logPureAlcCol).Value = Round(pureAlcohol, 1)  ' ���A����(g)
    wsLog.Cells(lastLogRow + 1, logDrunkCol).Value = Round(drankWeight, 1)  ' ���񂾗�(g)
    wsLog.Cells(lastLogRow + 1, logIdCol).Value = lastLogRow
    
    MsgBox "�L�^��ۑ����܂����I", vbInformation

    ' --- ���͗������Z�b�g�i�C�Ӂj ---
    txtNowWeight.Value = ""
    lblResult.Caption = ""
End Sub

Private Sub UserForm_Initialize()
    '�ϐ��錾
    Dim i As Long
    '�f�[�^��2�s�ڈȍ~�ɑ��݂���ꍇ�̂ݏ��������s
    If lastCell.Row >= 2 Then
        '2�s�ڂ���ŏI�s�܂ł͈̔͂��擾
        For i = 2 To lastCell.Row
            cmbSake.AddItem wsMaster.Cells(i, idCol).Value & "." & wsMaster.Cells(i, nameCol).Value
        Next i
    End If
End Sub
