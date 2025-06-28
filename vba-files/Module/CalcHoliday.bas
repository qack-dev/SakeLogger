Attribute VB_Name = "CalcHoliday"
Option Explicit

'================================================================================
' ���C���v���V�[�W���F�w�肵���N�̏j�����V�[�g�ɒǋL����
'================================================================================
Sub AddHolidaysForYear()
    Dim yearVal As Long
    Dim wsMaster As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim holidayDate As Date
    Dim holidayName As String
    Dim checkRange As Range
    
    '--- ���[�U�[�ɓ��͂𑣂� ---
    On Error Resume Next
    yearVal = Application.InputBox("�j�����X�g�ɒǉ��������N�i����j����͂��Ă��������B", "�N �w��", Year(Date))
    If yearVal = 0 Then Exit Sub ' �L�����Z�����ꂽ�ꍇ
    On Error GoTo 0
    
    '--- �����ݒ� ---
    Set wsMaster = ThisWorkbook.Sheets("�j���}�X�^") ' �����g�̃V�[�g���ɍ��킹�Ă�������
    
    '--- ���̔N�̏j�������Ƀ��X�g�ɂȂ����`�F�b�N�i�ȈՔŁj---
    If Application.WorksheetFunction.CountIf(wsMaster.Columns("A"), ">=" & DateSerial(yearVal, 1, 1)) > 0 And _
       Application.WorksheetFunction.CountIf(wsMaster.Columns("A"), "<=" & DateSerial(yearVal, 12, 31)) > 0 Then
        If MsgBox(yearVal & "�N�̏j���͊��ɒǉ�����Ă���\��������܂��B���s���܂����H", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If

    '--- �j�����v�Z���A�V�[�g�ɏ����o�� ---
    Application.ScreenUpdating = False
    
    For i = 1 To 12 ' 1������12���܂�
        holidayName = GetHolidayName(DateSerial(yearVal, i, 1), holidayDate) ' �����ő�\�`�F�b�N
    Next i
    
    ' 1�N���̓��t�����[�v���ďj������Ə�������
    For i = 1 To 366
        holidayDate = DateSerial(yearVal, 1, i)
        If Year(holidayDate) <> yearVal Then Exit For ' �N���ς������I��
        
        holidayName = GetHolidayName(holidayDate)
        
        If holidayName <> "" Then
            ' ���ɓ������t���Ȃ����m�F
            Set checkRange = wsMaster.Columns("A").Find(What:=holidayDate, LookIn:=xlFormulas, LookAt:=xlWhole)
            
            If checkRange Is Nothing Then ' ���t���Ȃ���ΒǋL
                lastRow = wsMaster.Cells(wsMaster.Rows.Count, "A").End(xlUp).Row + 1
                wsMaster.Cells(lastRow, "A").Value = holidayDate
                wsMaster.Cells(lastRow, "B").Value = holidayName
                wsMaster.Cells(lastRow, "C").Value = yearVal & "�N �����v�Z"
            End If
        End If
    Next i
    
    ' �V�[�g����t���ɕ��בւ�
    With wsMaster.Sort
        .SortFields.Clear
        .SortFields.Add key:=wsMaster.Range("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange wsMaster.UsedRange
        .Header = xlYes ' 1�s�ڂ����o���̏ꍇ
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Application.ScreenUpdating = True
    MsgBox yearVal & "�N�̏j����ǉ����܂����B�@�����ɂ��ύX�͎蓮�ŏC�����Ă��������B"
End Sub


'================================================================================
' �⏕�֐��Q
'================================================================================

' �w�肳�ꂽ���t�̏j������Ԃ��i�j���łȂ���΋󕶎��j
Private Function GetHolidayName(d As Date, Optional ByRef holidayDate As Date) As String
    Dim y As Long, m As Long, n As Long
    y = Year(d)
    m = Month(d)
    n = Day(d)
    Dim wd As Long
    wd = Weekday(d)
    
    ' 2000�N�ȍ~��ΏۂƂ���
    If y < 2000 Then GetHolidayName = "": Exit Function
    
    Dim resultName As String
    holidayDate = d ' ������
    
    '--- �Œ�j�� ---
    Select Case m
        Case 1: If n = 1 Then resultName = "����"
        Case 2: If n = 11 Then resultName = "�����L�O�̓�"
        Case 4: If n = 29 Then resultName = "���a�̓�"
        Case 5:
            If n = 3 Then
                resultName = "���@�L�O��"
            ElseIf n = 4 Then
                resultName = "�݂ǂ�̓�"
            ElseIf n = 5 Then
                resultName = "���ǂ��̓�"
            End If
        Case 8: If n = 11 Then resultName = "�R�̓�"
        Case 11:
            If n = 3 Then
                resultName = "�����̓�"
            ElseIf n = 23 Then
                resultName = "�ΘJ���ӂ̓�"
            End If
        Case 12: If y >= 2019 And n = 23 Then resultName = "" ' ��c�a����
                 If y >= 2020 And n = 23 Then resultName = "" ' ��c�a����
                 If y <= 2018 And n = 23 Then resultName = "�V�c�a����"
                 If y >= 2020 And Day(GetHolidayName(DateSerial(y, 2, 23))) = 23 Then resultName = "�V�c�a����" '2/23���j��
    End Select

    ' �V�c�a���� (2020�N�ȍ~)
    If y >= 2020 And m = 2 And n = 23 Then resultName = "�V�c�a����"

    '--- �n�b�s�[�}���f�[ ---
    If GetHappyMonday(y, m, 2) = n And m = 1 Then resultName = "���l�̓�"
    If GetHappyMonday(y, m, 3) = n And m = 7 Then resultName = "�C�̓�"
    If GetHappyMonday(y, m, 3) = n And m = 9 Then resultName = "�h�V�̓�"
    If GetHappyMonday(y, m, 2) = n And m = 10 Then resultName = "�X�|�[�c�̓�" '�̈�̓��������

    '--- �t���̓��E�H���̓� (�ȈՌv�Z��) ---
    If n = CInt(20.8431 + 0.242194 * (y - 1980) - CInt((y - 1980) / 4)) And m = 3 Then resultName = "�t���̓�"
    If n = CInt(23.2488 + 0.242194 * (y - 1980) - CInt((y - 1980) / 4)) And m = 9 Then resultName = "�H���̓�"

    '--- �U�֋x�� ---
    If resultName = "" And wd = 2 Then '���j���ŏj���łȂ��ꍇ
        If GetHolidayName(d - 1) <> "" Then '�O�����j���̏ꍇ
            resultName = "�U�֋x��"
        End If
    End If
    
    '--- �����̋x�� ---
    If resultName = "" And GetHolidayName(d - 1) <> "" And GetHolidayName(d + 1) <> "" Then
        resultName = "�����̋x��"
    End If

    GetHolidayName = resultName
End Function

' ��n���j���̓��t��Ԃ�
Private Function GetHappyMonday(y As Long, m As Long, weekNum As Long) As Long
    Dim firstDay As Date
    firstDay = DateSerial(y, m, 1)
    GetHappyMonday = ((weekNum - 1) * 7) + 1 + (9 - Weekday(firstDay, vbMonday)) Mod 7
End Function

