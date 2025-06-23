VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "��ʂ�o�^"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6915
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm1"
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

    lastRow = lastCell.Row
    For i = 2 To lastRow
        'If ws.Cells(i, nameCol).Value = sakeName Then
        If ws.Cells(i, idCol).Value & "." & ws.Cells(i, nameCol).Value = sakeName Then
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
            lblResult.Caption = "���A���R�[���ʁF" & Format(pureAlcohol, "0.0") & " g"
            Exit Sub
        End If
    Next i
    Call releaseObj
End Sub

Private Sub btnSave_Click()

End Sub

Private Sub cmbSake_Change()
    Dim targetRow As Long
    Dim lastRow As Long
    Dim i As Long
    Call setObj

    lastRow = lastCell.Row
    
    With ws
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

    Call releaseObj
End Sub

Private Sub UserForm_Initialize()
    '�ϐ��錾
    Dim i As Long
    Call setObj
    '�f�[�^��2�s�ڈȍ~�ɑ��݂���ꍇ�̂ݏ��������s
    If lastCell.Row >= 2 Then
        '2�s�ڂ���ŏI�s�܂ł͈̔͂��擾
        For i = 2 To lastCell.Row
            cmbSake.AddItem ws.Cells(i, idCol).Value & "." & ws.Cells(i, nameCol).Value
        Next i
    End If
    Call releaseObj
End Sub
