VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'�O���[�o���萔
Private Const inputCol As Integer = 2 '�����̖��O��

Option Explicit

Private Sub UserForm_Initialize()
    '�ϐ��錾
    Dim ws As Worksheet
    Dim lastCell As Range
    Dim dataRange As Range
    '�I�u�W�F�N�g�i�[
    Set ws = ThisWorkbook.Worksheets("�����}�X�^")
    'ws.Rows.Count �ŃV�[�g�̍ő�s�����擾���A�������� End(xlUp) �Ńf�[�^�̂���ŏI�Z����T���B
    Set lastCell = ws.Cells(ws.Rows.Count, inputCol).End(xlUp)

    '�f�[�^��2�s�ڈȍ~�ɑ��݂���ꍇ�̂ݏ��������s
    If lastCell.Row >= 2 Then
        '2�s�ڂ���ŏI�s�܂ł͈̔͂��擾
        Set dataRange = ws.Range(Cells(2, inputCol), lastCell)
        'ComboBox�̃��X�g�ɔ͈͂̒l��ݒ�
        Me.ComboBox1.List = dataRange.Value
    End If
End Sub
