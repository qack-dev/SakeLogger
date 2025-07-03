Attribute VB_Name = "M_SakeLogics"
Option Explicit

' ���񂾗ʂƏ��A���R�[���ʂ��v�Z����
Public Function CalculateAlcoholInfo(ByVal sakeName As String, ByVal currentWeight As Double, ByRef drankWeight As Double, ByRef pureAlcohol As Double) As Boolean
    Dim abv As Double, fullWeight As Double, emptyWeight As Double
    Dim previousWeight As Double
    Dim i As Long
    Dim isFound As Boolean
    Dim masterSheet As Worksheet
    Dim logSheet As Worksheet

    On Error GoTo ErrorHandler

    Set masterSheet = M_SakeForm.GetMasterSheet()
    Set logSheet = M_SakeForm.GetLogSheet()

    ' �}�X�^����x���E�d�ʏ����擾
    isFound = False
    For i = 2 To M_SakeForm.GetLastMasterDataCell().Row
        If masterSheet.Cells(i, COL_MASTER_ID).Value & "." & masterSheet.Cells(i, COL_MASTER_NAME).Value = sakeName Then
            abv = masterSheet.Cells(i, COL_MASTER_ALCOHOL).Value
            fullWeight = masterSheet.Cells(i, COL_MASTER_FULL_WEIGHT).Value
            
            If IsEmpty(masterSheet.Cells(i, COL_MASTER_EMPTY_WEIGHT).Value) Or masterSheet.Cells(i, COL_MASTER_EMPTY_WEIGHT).Value = "" Then
                MsgBox "����: ���̂����͋�{�g���d�ʂ����o�^�ł��B��d�ʂ�0g�Ƃ��Čv�Z�𑱍s���܂��B", vbInformation
                emptyWeight = 0
            Else
                emptyWeight = masterSheet.Cells(i, COL_MASTER_EMPTY_WEIGHT).Value
            End If

            If currentWeight > fullWeight Or currentWeight < emptyWeight Then
                MsgBox "���݂̏d�ʂ̒l���s���ł��i���^���d�ʂ𒴂��Ă��邩�A��d�ʂ�������Ă��܂��j�B", vbExclamation
                CalculateAlcoholInfo = False
                Exit Function
            End If
            
            isFound = True
            Exit For
        End If
    Next i

    If Not isFound Then
        MsgBox "�����}�X�^�ɊY�����邨����������܂���B", vbCritical
        CalculateAlcoholInfo = False
        Exit Function
    End If

    ' ���񂾗ʂ��v�Z
    If frmSakeLogger.optNewOpen.Value Then
        drankWeight = fullWeight - currentWeight
    ElseIf frmSakeLogger.optContinued.Value Then
        previousWeight = GetPreviousWeight(sakeName, logSheet)
        If previousWeight = -1 Then
            MsgBox "���̂����̉ߋ��̋L�^��������܂���B�u�V�K�J���v��I�����Ă��������B", vbExclamation
            CalculateAlcoholInfo = False
            Exit Function
        End If
        drankWeight = previousWeight - currentWeight
    Else
        MsgBox "�u�V�K�J���v�܂��́u��������v��I�����Ă��������B", vbExclamation
        CalculateAlcoholInfo = False
        Exit Function
    End If

    ' ���A���R�[���ʂ��v�Z (��d: 0.8)
    pureAlcohol = drankWeight * (abv / 100) * 0.8

    CalculateAlcoholInfo = True
    Exit Function

ErrorHandler:
    MsgBox "�v�Z���ɃG���[���������܂����B" & vbCrLf & Err.Description, vbCritical
    CalculateAlcoholInfo = False
End Function

' �w�肳�ꂽ�����̑O��̏d�ʂ��擾����
Private Function GetPreviousWeight(ByVal sakeName As String, ByVal logSheet As Worksheet) As Double
    Dim lastRow As Long
    Dim i As Long

    GetPreviousWeight = -1 ' ������Ȃ������ꍇ�̃f�t�H���g�l
    lastRow = logSheet.Cells(logSheet.Rows.Count, COL_LOG_ID).End(xlUp).Row

    For i = lastRow To 2 Step -1
        If logSheet.Cells(i, COL_LOG_NAME).Value = sakeName Then
            GetPreviousWeight = logSheet.Cells(i, COL_LOG_CURRENT_WEIGHT).Value
            Exit Function
        End If
    Next i
End Function

' ���t������ 'yyyy/mm/dd' �`�������؂���
Public Function IsValidDateFormat(ByVal dateString As String) As Boolean
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = "^\d{4}/\d{2}/\d{2}$"
    
    If regEx.Test(dateString) And IsDate(dateString) Then
        IsValidDateFormat = True
    Else
        IsValidDateFormat = False
    End If
End Function
