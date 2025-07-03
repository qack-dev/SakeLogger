Attribute VB_Name = "CalcHoliday"
Option Explicit

'================================================================================
' �O���[�o���萔��`
'================================================================================
Private Const COL_DATE As Long = 1      ' A��: ���t
Private Const COL_NAME As Long = 2      ' B��: �j����
Private Const COL_NOTES As Long = 3     ' C��: ���l


'================================================================================
' ���C���v���V�[�W���F���[�U�[�Ƃ̑Θb�ƃV�[�g�ւ̏������݂�S��
'================================================================================
Sub AddHolidaysForYear()
    Dim yearVal As Long
    Dim wsHoliday As Worksheet
    Dim holidays As Object ' Scripting.Dictionary
    Dim holidayDate As Variant
    Dim lastRow As Long
    Dim checkRange As Range
    
    On Error Resume Next
    yearVal = Application.InputBox("�j�����X�g�ɒǉ��������N�i����j����͂��Ă��������B", "�N �w��", Year(Date))
    If yearVal = 0 Then Exit Sub ' �L�����Z����
    On Error GoTo 0
    
    Set wsHoliday = ThisWorkbook.Sheets("�j���}�X�^")
    
    If Application.WorksheetFunction.CountIfs(wsHoliday.Columns(COL_DATE), ">=" & DateSerial(yearVal, 1, 1), wsHoliday.Columns(COL_DATE), "<=" & DateSerial(yearVal, 12, 31)) > 0 Then
        If MsgBox(yearVal & "�N�̏j���͊��ɒǉ�����Ă���\��������܂��B�d���𖳎����Ď��s���܂����H", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If

    Application.ScreenUpdating = False
    
    Set holidays = GenerateHolidayList(yearVal)
    
    For Each holidayDate In holidays.Keys
        Set checkRange = wsHoliday.Columns(COL_DATE).Find(What:=CDate(holidayDate), LookIn:=xlFormulas, LookAt:=xlWhole)
        
        If checkRange Is Nothing Then
            lastRow = wsHoliday.Cells(wsHoliday.Rows.Count, COL_DATE).End(xlUp).Row + 1
            With wsHoliday.Cells(lastRow, COL_DATE)
                .Value = CDate(holidayDate)
                .NumberFormatLocal = "yyyy/mm/dd"
            End With
            wsHoliday.Cells(lastRow, COL_NAME).Value = holidays(holidayDate)
            wsHoliday.Cells(lastRow, COL_NOTES).Value = yearVal & "�N �����v�Z"
        End If
    Next holidayDate
    
    If wsHoliday.Cells(wsHoliday.Rows.Count, COL_DATE).End(xlUp).Row > 1 Then
        With wsHoliday.Sort
            .SortFields.Clear
            .SortFields.Add key:=wsHoliday.Columns(COL_DATE), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange wsHoliday.UsedRange
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With
    End If
    
    If lastRow > 0 Then
        Call M_SheetUtils.FormatTable(wsHoliday.Range(wsHoliday.Cells(1, COL_DATE), wsHoliday.Cells(lastRow, COL_NOTES)), True)
    End If
    
    Set holidays = Nothing
    Application.ScreenUpdating = True
    MsgBox yearVal & "�N�̏j����ǉ����܂����B�K�v�ɉ����Ď蓮�ŕύX�E�폜���Ă��������B"
End Sub


'================================================================================
' �⏕�֐��Q
'================================================================================

Private Function GenerateHolidayList(ByVal y As Long) As Object
    Dim hList As Object
    Set hList = CreateObject("Scripting.Dictionary")
    
    Dim d As Variant
    Dim tempDate As Date
    
    '--- Step 1: �@���Œ�߂�ꂽ�u�Œ���v�̏j�����܂����X�g�A�b�v ---
    For Each d In Array( _
        DateSerial(y, 1, 1), _
        GetHappyMonday(y, 1, 2), _
        DateSerial(y, 2, 11), _
        IIf(y >= 2020, DateSerial(y, 2, 23), 0), _
        GetShunbun(y), _
        DateSerial(y, 4, 29), _
        DateSerial(y, 5, 3), _
        DateSerial(y, 5, 4), _
        DateSerial(y, 5, 5), _
        IIf(y >= 2003, GetHappyMonday(y, 7, 3), 0), _
        IIf(y >= 2016, DateSerial(y, 8, 11), 0), _
        IIf(y >= 2003, GetHappyMonday(y, 9, 3), 0), _
        GetShubun(y), _
        IIf(y >= 2000, GetHappyMonday(y, 10, 2), 0), _
        DateSerial(y, 11, 3), _
        DateSerial(y, 11, 23), _
        IIf(y <= 2018, DateSerial(y, 12, 23), 0) _
    )
        If d > 0 Then ' �[���ȊO�̗L���ȓ��t�̂�
            If Not hList.Exists(CDate(d)) Then
                hList.Add CDate(d), GetPrimaryHolidayName(CDate(d))
            End If
        End If
    Next d

    '--- Step 2: �U�֋x����ǉ� ---
    Dim originalDates As Variant
    originalDates = hList.Keys
    
    For Each d In originalDates
        If Weekday(d) = vbSunday Then
            tempDate = d + 1
            Do While hList.Exists(tempDate)
                tempDate = tempDate + 1
            Loop
            hList.Add tempDate, "�U�֋x��"
        End If
    Next d
    
    '--- Step 3: �����̋x����ǉ� ---
    originalDates = hList.Keys
    For Each d In originalDates
        If hList.Exists(d + 2) And Not hList.Exists(d + 1) Then
            hList.Add d + 1, "�����̋x��"
        End If
    Next d

    Set GenerateHolidayList = hList
End Function


Private Function GetPrimaryHolidayName(ByVal d As Date) As String
    Dim y As Long: y = Year(d)
    Dim m As Long: m = Month(d)
    Dim n As Long: n = Day(d)
    
    Select Case m
        Case 1
            If n = 1 Then GetPrimaryHolidayName = "����" Else GetPrimaryHolidayName = "���l�̓�"
        Case 2
            If n = 11 Then GetPrimaryHolidayName = "�����L�O�̓�" Else GetPrimaryHolidayName = "�V�c�a����"
        Case 3
            GetPrimaryHolidayName = "�t���̓�"
        Case 4
            GetPrimaryHolidayName = "���a�̓�"
        Case 5
            If n = 3 Then GetPrimaryHolidayName = "���@�L�O��"
            If n = 4 Then GetPrimaryHolidayName = "�݂ǂ�̓�"
            If n = 5 Then GetPrimaryHolidayName = "���ǂ��̓�"
        Case 7
            GetPrimaryHolidayName = "�C�̓�"
        Case 8
            GetPrimaryHolidayName = "�R�̓�"
        Case 9
            If Day(d) = Day(GetShubun(y)) Then GetPrimaryHolidayName = "�H���̓�" Else GetPrimaryHolidayName = "�h�V�̓�"
        Case 10
            GetPrimaryHolidayName = "�X�|�[�c�̓�"
        Case 11
            If n = 3 Then GetPrimaryHolidayName = "�����̓�" Else GetPrimaryHolidayName = "�ΘJ���ӂ̓�"
        Case 12
            GetPrimaryHolidayName = "�V�c�a����"
    End Select
End Function


Private Function GetHappyMonday(ByVal y As Long, ByVal m As Long, ByVal weekNum As Long) As Date
    GetHappyMonday = DateSerial(y, m, (weekNum - 1) * 7 + 1 + (8 - Weekday(DateSerial(y, m, 1), vbMonday)) Mod 7)
End Function

Private Function GetShunbun(ByVal y As Long) As Date
    Dim d As Integer
    d = Int(20.8431 + 0.242194 * (y - 1980) - Int((y - 1980) / 4))
    GetShunbun = DateSerial(y, 3, d)
End Function

Private Function GetShubun(ByVal y As Long) As Date
    Dim d As Integer
    d = Int(23.2488 + 0.242194 * (y - 1980) - Int((y - 1980) / 4))
    GetShubun = DateSerial(y, 9, d)
End Function
