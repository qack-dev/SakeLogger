Attribute VB_Name = "M_SakeForm"
Option Explicit

Private masterSheet As Worksheet
Private logSheet As Worksheet
Private lastMasterDataCell As Range

' ���[�U�[�t�H�[����\�����郁�C���v���V�[�W��
Public Sub ShowSakeLoggerForm()
    If Not InitializeObjects() Then Exit Sub
    
    ' �t�H�[���\���O�Ƀ}�X�^�V�[�g���A�N�e�B�u��
    masterSheet.Activate
    
    Load frmSakeLogger
    frmSakeLogger.Show
    Unload frmSakeLogger
    
    ' �t�H�[����������A���O�V�[�g���A�N�e�B�u�����A���`
    logSheet.Activate
    Dim lastLogCell As Range
    Set lastLogCell = logSheet.Cells(logSheet.Rows.Count, COL_LOG_ID).End(xlUp)
    If lastLogCell.Row > 1 Then
        Call M_SheetUtils.FormatTable(logSheet.Range(logSheet.Cells(1, COL_LOG_ID), lastLogCell.Offset(0, COL_LOG_COMMENT - 1)), True)
    End If
    
    ReleaseObjects
End Sub

' �I�u�W�F�N�g�ϐ���������
Private Function InitializeObjects() As Boolean
    On Error GoTo ErrorHandler
    Set masterSheet = ThisWorkbook.Worksheets(SHEET_MASTER)
    Set logSheet = ThisWorkbook.Worksheets(SHEET_LOG)
    Set lastMasterDataCell = masterSheet.Cells(masterSheet.Rows.Count, COL_MASTER_ID).End(xlUp)
    InitializeObjects = True
    Exit Function

ErrorHandler:
    MsgBox "�������Ɏ��s���܂����B�V�[�g�����ύX����Ă��Ȃ����m�F���Ă��������B" & vbCrLf & _
           "�G���[: " & Err.Description, vbCritical
    InitializeObjects = False
End Function

' �I�u�W�F�N�g�ϐ������
Private Sub ReleaseObjects()
    Set masterSheet = Nothing
    Set logSheet = Nothing
    Set lastMasterDataCell = Nothing
End Sub

' frmSakeLogger����Ăяo�������J�v���V�[�W��
Public Function GetMasterSheet() As Worksheet
    Set GetMasterSheet = masterSheet
End Function

Public Function GetLogSheet() As Worksheet
    Set GetLogSheet = logSheet
End Function

Public Function GetLastMasterDataCell() As Range
    Set GetLastMasterDataCell = lastMasterDataCell
End Function