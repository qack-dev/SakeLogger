Attribute VB_Name = "M_SakeLogics"
Option Explicit

' ���񂾗ʂƏ��A���R�[���ʂ��v�Z
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

    ' �}�X�^�[����x���E�d�ʏ����擾
    isFound = False
    For i = 2 To M_SakeForm.GetLastMasterDataCell().Row
        If masterSheet.Cells(i, COL_MASTER_ID).Value & "." & masterSheet.Cells(i, COL_MASTER_NAME).Value = sakeName Then
            abv = masterSheet.Cells(i, COL_MASTER_ALCOHOL).Value
            fullWeight = masterSheet.Cells(i, COL_MASTER_FULL_WEIGHT).Value
            
            If IsEmpty(masterSheet.Cells(i, COL_MASTER_EMPTY_WEIGHT).Value) Or masterSheet.Cells(i, COL_MASTER_EMPTY_WEIGHT).Value = "" Then
                MsgBox "����: ���̂����͋󂫗e��d�ʂ����o�^�ł��B�󂫗e��d�ʂ�0g�Ƃ��Čv�Z�𑱍s���܂��B", vbInformation
                emptyWeight = 0
            Else
                emptyWeight = masterSheet.Cells(i, COL_MASTER_EMPTY_WEIGHT).Value
            End If

            If currentWeight > fullWeight Or currentWeight < emptyWeight Then
                MsgBox "���݂̏d�ʂ̒l���s���ł��i���^�����d�ʂ𒴂��Ă��邩�A�󂫗e��d�ʂ�������Ă��܂��j�B", vbExclamation
                CalculateAlcoholInfo = False
                Exit Function
            End If
            
            isFound = True
            Exit For
        End If
    Next i

    If Not isFound Then
        MsgBox "�����}�X�^�[�ɊY�����邨����������܂���ł����B", vbCritical
        CalculateAlcoholInfo = False
        Exit Function
    End If

    ' ���񂾗ʂ��v�Z
    If frmSakeLogger.optNewOpen.Value Then
        drankWeight = fullWeight - currentWeight
    ElseIf frmSakeLogger.optContinued.Value Then
        previousWeight = GetPreviousWeight(sakeName, logSheet)
        If previousWeight = -1 Then
            MsgBox "���̂����̒��O�̋L�^��������܂���ł����B�u�V�K�J���v��I�����Ă��������B", vbExclamation
            CalculateAlcoholInfo = False
            Exit Function
        End If
        drankWeight = previousWeight - currentWeight
    Else
        MsgBox "�u�V�K�J���v�܂��́u�p���v��I�����Ă��������B", vbExclamation
        CalculateAlcoholInfo = False
        Exit Function
    End If

    ' ���A���R�[���ʂ��v�Z (���x: 0.8)
    pureAlcohol = drankWeight * (abv / 100) * 0.8

    CalculateAlcoholInfo = True
    Exit Function

ErrorHandler:
    MsgBox "�v�Z���ɃG���[���������܂����B" & vbCrLf & Err.Description, vbCritical
    CalculateAlcoholInfo = False
End Function

' �w�肳�ꂽ�����̒��O�̏d�ʂ��擾����
Public Function GetPreviousWeight(ByVal sakeName As String, ByVal logSheet As Worksheet) As Double
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

' �t�H�[������̓��͂Ɋ�Â��Č��݂̏d�ʂ��v�Z���鋤�ʊ֐�
Public Function CalculateCurrentWeightFromInput( _
    ByVal frm As Object, _
    ByVal fullWeight As Double, _
    ByVal emptyWeight As Double, _
    ByVal previousWeight As Double, _
    ByRef currentWeight As Double _
) As Boolean
    Dim inputCount As Integer
    Dim tempWeight As Double
    
    On Error GoTo ErrorHandler
    
    inputCount = 0
    If Trim(frm.txtNowWeight.Value) <> "" Then inputCount = inputCount + 1
    If Trim(frm.txtNowPercent.Value) <> "" Then inputCount = inputCount + 1
    If Trim(frm.txtNowNum.Value) <> "" Then inputCount = inputCount + 1
    
    ' 1. ���̓\�[�X�̓���ƃo���f�[�V����
    If inputCount = 0 Then
        MsgBox "���݂̏d�ʁA�c��(%)�A�܂��͔t������͂��Ă��������B", vbExclamation
        CalculateCurrentWeightFromInput = False
        Exit Function
    ElseIf inputCount > 1 Then
        MsgBox "���͉ӏ���1�����ɂ��Ă��������B", vbExclamation
        CalculateCurrentWeightFromInput = False
        Exit Function
    End If
    
    ' 3. �v�Z���W�b�N
    If Trim(frm.txtNowWeight.Value) <> "" Then
        If Not IsNumeric(frm.txtNowWeight.Value) Then
            MsgBox "���݂̏d�ʂ͐��l����͂��Ă��������B", vbExclamation
            CalculateCurrentWeightFromInput = False
            Exit Function
        End If
        currentWeight = CDbl(frm.txtNowWeight.Value)
    ElseIf Trim(frm.txtNowPercent.Value) <> "" Then
        If Not IsNumeric(frm.txtNowPercent.Value) Then
            MsgBox "�c��(%)�͐��l����͂��Ă��������B", vbExclamation
            CalculateCurrentWeightFromInput = False
            Exit Function
        End If
        tempWeight = CDbl(frm.txtNowPercent.Value)
        If tempWeight < 0 Or tempWeight > 100 Then
            MsgBox "�c��(%)��0����100�͈̔͂œ��͂��Ă��������B", vbExclamation
            CalculateCurrentWeightFromInput = False
            Exit Function
        End If
        currentWeight = emptyWeight + ((fullWeight - emptyWeight) * (tempWeight / 100))
    ElseIf Trim(frm.txtNowNum.Value) <> "" Then
        If Not IsNumeric(frm.txtNowNum.Value) Then
            MsgBox "�t���͐��l����͂��Ă��������B", vbExclamation
            CalculateCurrentWeightFromInput = False
            Exit Function
        End If
        If previousWeight = -1 Then ' previousWeight���擾�ł��Ȃ������ꍇ
            MsgBox "�p���L�^�̃f�[�^��������܂���B�V�K�J����I�����邩�A�ʂ̓��͕��@���g�p���Ă��������B", vbExclamation
            CalculateCurrentWeightFromInput = False
            Exit Function
        End If
        
        Dim drankAmount As Double
        ' �J���҃���: ���́u���񂾗ʁv�̌v�Z���́A���e�ʑS�̂ɔt�����|����Ƃ�����������ȃ��W�b�N�ł��B
        ' �Ⴆ�΁AtxtNowNum���u1�v�̏ꍇ�A���e�ʑS�̂����񂾂Ƃ݂Ȃ���܂��B
        ' �����u1�t������̗ʁv����`����Ă���ꍇ�́A���̗ʂ�txtNowNum���|������������I��������܂���B
        ' ���݂̃��W�b�N�́AtxtNowNum���u�����v�Ƃ��ċ@�\���Ă���悤�Ɍ����܂��B
        drankAmount = (fullWeight - emptyWeight) * CDbl(frm.txtNowNum.Value)
        currentWeight = previousWeight - drankAmount
    End If
    
    CalculateCurrentWeightFromInput = True
    Exit Function

ErrorHandler:
    MsgBox "���͒l�̏������ɃG���[���������܂����B" & vbCrLf & Err.Description, vbCritical
    CalculateCurrentWeightFromInput = False
End Function
