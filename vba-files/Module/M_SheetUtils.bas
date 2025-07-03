Attribute VB_Name = "M_SheetUtils"
Option Explicit

' �W�v�V�[�g���X�V����
Public Sub UpdateSummarySheet()
    Dim logSheet As Worksheet, summarySheet As Worksheet
    Dim summaryData As Object ' Scripting.Dictionary
    Dim lastRow As Long, i As Long
    Dim logDate As String, alcoholAmount As Double

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    Set logSheet = ThisWorkbook.Worksheets(SHEET_LOG)
    Set summarySheet = ThisWorkbook.Worksheets(SHEET_SUMMARY)
    Set summaryData = CreateObject("Scripting.Dictionary")

    ' �������O���W�v
    lastRow = logSheet.Cells(logSheet.Rows.Count, COL_LOG_DATE).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "�������O�Ƀf�[�^������܂���B", vbInformation
        GoTo CleanUp
    End If

    Dim logDataArray As Variant
    logDataArray = logSheet.Range(logSheet.Cells(2, COL_LOG_DATE), logSheet.Cells(lastRow, COL_LOG_PURE_ALCOHOL)).Value

    ' �f�[�^�������ɏW�v
    For i = 1 To UBound(logDataArray, 1)
        logDate = Format(logDataArray(i, 1), "yyyy/mm/dd")
        alcoholAmount = logDataArray(i, COL_LOG_PURE_ALCOHOL - COL_LOG_DATE + 1)

        If summaryData.Exists(logDate) Then
            summaryData(logDate) = summaryData(logDate) + alcoholAmount
        Else
            summaryData.Add logDate, alcoholAmount
        End If
    Next i

    ' �W�v�V�[�g�ɏ����o��
    summarySheet.Cells.ClearContents
    summarySheet.Range(summarySheet.Cells(1, COL_SUMMARY_DATE), summarySheet.Cells(1, COL_SUMMARY_PURE_ALCOHOL)).Value = Array("���t", "���A���R�[����")

    i = 2
    Dim key As Variant
    For Each key In summaryData.Keys
        summarySheet.Cells(i, COL_SUMMARY_DATE).NumberFormat = "yyyy/mm/dd"
        summarySheet.Cells(i, COL_SUMMARY_PURE_ALCOHOL).NumberFormat = "0.0"
        summarySheet.Cells(i, COL_SUMMARY_DATE).Value = CDate(key)
        summarySheet.Cells(i, COL_SUMMARY_PURE_ALCOHOL).Value = Round(summaryData(key), 1)
        i = i + 1
    Next key

    ' ���t�Ń\�[�g
    If i > 2 Then
        summarySheet.Range(summarySheet.Cells(2, COL_SUMMARY_DATE), summarySheet.Cells(i - 1, COL_SUMMARY_PURE_ALCOHOL)).Sort _
            Key1:=summarySheet.Cells(2, COL_SUMMARY_DATE), Order1:=xlAscending, Header:=xlNo
    End If

    Call FormatTable(summarySheet.Range(summarySheet.Cells(1, COL_SUMMARY_DATE), summarySheet.Cells(i - 1, COL_SUMMARY_PURE_ALCOHOL)), False)

    MsgBox "�W�v�V�[�g���X�V���܂����B", vbInformation

CleanUp:
    Application.ScreenUpdating = True
    Set logSheet = Nothing
    Set summarySheet = Nothing
    Set summaryData = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "�W�v�V�[�g�̍X�V���ɃG���[���������܂����B" & vbCrLf & Err.Description, vbCritical
    GoTo CleanUp
End Sub

' �O���t���쐬�E�X�V����
Public Sub CreateOrUpdateGraph()
    Dim summarySheet As Worksheet
    Dim chartObj As ChartObject
    Dim lastRow As Long

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    Set summarySheet = ThisWorkbook.Worksheets(SHEET_SUMMARY)
    summarySheet.Activate

    lastRow = summarySheet.Cells(summarySheet.Rows.Count, COL_SUMMARY_DATE).End(xlUp).Row
    If lastRow < 2 Then Exit Sub ' �f�[�^���Ȃ��ꍇ�͏I��

    ' �����̃O���t���폜
    For Each chartObj In summarySheet.ChartObjects
        chartObj.Delete
    Next chartObj

    ' �V�����O���t���쐬
    Set chartObj = summarySheet.ChartObjects.Add(Left:=210, Top:=45, Width:=500, Height:=300)

    With chartObj.Chart
        .ChartType = xlColumnClustered
        .SetSourceData Source:=summarySheet.Range(summarySheet.Cells(1, COL_SUMMARY_DATE), summarySheet.Cells(lastRow, COL_SUMMARY_PURE_ALCOHOL))
        .HasTitle = True
        .ChartTitle.Text = "���� ���A���R�[���ێ�� (g)"
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "���t"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "���A���R�[���� (g)"
    End With

    MsgBox "�O���t���쐬�E�X�V���܂����B", vbInformation

CleanUp:
    Application.ScreenUpdating = True
    Set summarySheet = Nothing
    Set chartObj = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "�O���t�쐬���ɃG���[���������܂����B" & vbCrLf & Err.Description, vbCritical
    GoTo CleanUp
End Sub

' ���v����ǉ�����
Public Sub AddTotalFields()
    Dim summarySheet As Worksheet
    Dim lastRow As Long
    Dim formulaString As String

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    Set summarySheet = ThisWorkbook.Worksheets(SHEET_SUMMARY)
    summarySheet.Activate

    lastRow = summarySheet.Cells(summarySheet.Rows.Count, COL_SUMMARY_PURE_ALCOHOL).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    ' �w�b�_��ݒ�
    summarySheet.Cells(1, COL_SUMMARY_TOTAL).Value = "�݌v���v"
    summarySheet.Cells(1, COL_SUMMARY_MONTHLY_TOTAL).Value = "�������v"

    ' ������ݒ� (R1C1�`��)
    summarySheet.Cells(2, COL_SUMMARY_TOTAL).NumberFormat = "0.0"
    summarySheet.Cells(2, COL_SUMMARY_TOTAL).FormulaR1C1 = "=SUM(R2C" & COL_SUMMARY_PURE_ALCOHOL & ":R" & lastRow & "C" & COL_SUMMARY_PURE_ALCOHOL & ")"

    ' �������v�p�̃w���p�[�Z��
    summarySheet.Cells(1, COL_SUMMARY_HELPER_START_DATE).Value = "����1��"
    summarySheet.Cells(1, COL_SUMMARY_HELPER_END_DATE).Value = "��������"
    summarySheet.Cells(2, COL_SUMMARY_HELPER_START_DATE).NumberFormat = "yyyy/mm/dd"
    summarySheet.Cells(2, COL_SUMMARY_HELPER_END_DATE).NumberFormat = "yyyy/mm/dd"
    summarySheet.Cells(2, COL_SUMMARY_HELPER_START_DATE).FormulaR1C1 = "=DATE(YEAR(TODAY()), MONTH(TODAY()), 1)"
    summarySheet.Cells(2, COL_SUMMARY_HELPER_END_DATE).FormulaR1C1 = "=EOMONTH(RC[-1], 0)"

    ' �������v�̐���
    formulaString = "=SUMIFS(" & _
        "R2C" & COL_SUMMARY_PURE_ALCOHOL & ":R" & lastRow & "C" & COL_SUMMARY_PURE_ALCOHOL & "," & _
        "R2C" & COL_SUMMARY_DATE & ":R" & lastRow & "C" & COL_SUMMARY_DATE & "," & _
        """>=""&R2C" & COL_SUMMARY_HELPER_START_DATE & "," & _
        "R2C" & COL_SUMMARY_DATE & ":R" & lastRow & "C" & COL_SUMMARY_DATE & "," & _
        """<=""&R2C" & COL_SUMMARY_HELPER_END_DATE & ")"

    summarySheet.Cells(2, COL_SUMMARY_MONTHLY_TOTAL).NumberFormat = "0.0"
    summarySheet.Cells(2, COL_SUMMARY_MONTHLY_TOTAL).FormulaR1C1 = formulaString

    Call FormatTable(summarySheet.Range(summarySheet.Cells(1, COL_SUMMARY_TOTAL), summarySheet.Cells(2, COL_SUMMARY_HELPER_END_DATE)), False)

    MsgBox "���v����ǉ����܂����B", vbInformation

CleanUp:
    Application.ScreenUpdating = True
    Set summarySheet = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "���v���̒ǉ����ɃG���[���������܂����B" & vbCrLf & Err.Description, vbCritical
    GoTo CleanUp
End Sub

' �e�[�u���̏�����ݒ肷��
Public Sub FormatTable(ByVal targetRange As Range, ByVal hasFilter As Boolean)
    If hasFilter Then
        With targetRange.Parent.AutoFilter
            If .FilterMode Then .ShowAllData
        End With
        targetRange.Parent.AutoFilterMode = False
        targetRange.Rows(1).AutoFilter
    End If
    
    targetRange.EntireColumn.AutoFit
    targetRange.Borders.LineStyle = xlContinuous
End Sub
