Attribute VB_Name = "M_Constants"
Option Explicit

' --- �V�[�g�� ---
Public Const SHEET_MASTER As String = "�����}�X�^"
Public Const SHEET_LOG As String = "�����L�^"
Public Const SHEET_SUMMARY As String = "�W�v"
Public Const SHEET_HOLIDAY As String = "�j���}�X�^"

' --- �����}�X�^�V�[�g��C���f�b�N�X ---
Public Const COL_MASTER_ID As Long = 1
Public Const COL_MASTER_NAME As Long = 2
Public Const COL_MASTER_KIND As Long = 3
Public Const COL_MASTER_ALCOHOL As Long = 4
Public Const COL_MASTER_FULL_WEIGHT As Long = 5
Public Const COL_MASTER_EMPTY_WEIGHT As Long = 6

' --- �����L�^�V�[�g��C���f�b�N�X ---
Public Const COL_LOG_ID As Long = 1
Public Const COL_LOG_DATE As Long = 2
Public Const COL_LOG_NAME As Long = 3
Public Const COL_LOG_CURRENT_WEIGHT As Long = 4
Public Const COL_LOG_PURE_ALCOHOL As Long = 5
Public Const COL_LOG_DRANK_WEIGHT As Long = 6
Public Const COL_LOG_COMMENT As Long = 7

' --- �W�v�V�[�g��C���f�b�N�X ---
Public Const COL_SUMMARY_DATE As Long = 1
Public Const COL_SUMMARY_PURE_ALCOHOL As Long = 2
Public Const COL_SUMMARY_TOTAL As Long = 6
Public Const COL_SUMMARY_MONTHLY_TOTAL As Long = 7
Public Const COL_SUMMARY_HELPER_START_DATE As Long = 8
Public Const COL_SUMMARY_HELPER_END_DATE As Long = 9