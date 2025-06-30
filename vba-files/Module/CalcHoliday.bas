Attribute VB_Name = "CalcHoliday"
Option Explicit

'================================================================================
' グローバル定数宣言
'================================================================================
Private Const COL_DATE As Long = 1      ' A列: 日付を格納する列
Private Const COL_NAME As Long = 2      ' B列: 祝日名を格納する列
Private Const COL_NOTES As Long = 3     ' C列: 備考を格納する列


'================================================================================
' メインプロシージャ：ユーザーとの対話とシートへの書き込みを担当
'================================================================================
Sub AddHolidaysForYear()
    Dim yearVal As Long
    Dim wsHoliday As Worksheet
    Dim holidays As Object ' Scripting.Dictionary
    Dim holidayDate As Variant
    Dim lastRow As Long
    Dim checkRange As Range
    
    On Error Resume Next
    yearVal = Application.InputBox("祝日リストに追加したい年（西暦）を入力してください。", "年 指定", Year(Date))
    If yearVal = 0 Then Exit Sub ' キャンセル時
    On Error GoTo 0
    
    Set wsHoliday = ThisWorkbook.Sheets("祝日マスタ")
    
    If Application.WorksheetFunction.CountIfs(wsHoliday.Columns(COL_DATE), ">=" & DateSerial(yearVal, 1, 1), wsHoliday.Columns(COL_DATE), "<=" & DateSerial(yearVal, 12, 31)) > 0 Then
        If MsgBox(yearVal & "年の祝日は既に追加されている可能性があります。重複を避けて続行しますか？", vbQuestion + vbYesNo) = vbNo Then
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
            wsHoliday.Cells(lastRow, COL_NOTES).Value = yearVal & "年 自動計算"
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
        Call shape(wsHoliday.Range(Cells(1, COL_DATE), Cells(lastRow, COL_NOTES)), True)
    End If
    
    Set holidays = Nothing
    Application.ScreenUpdating = True
    MsgBox yearVal & "年の祝日を追加しました。法改正による変更は手動で修正してください。"
End Sub


'================================================================================
' 補助関数群
'================================================================================

Private Function GenerateHolidayList(ByVal y As Long) As Object
    Dim hList As Object
    Set hList = CreateObject("Scripting.Dictionary")
    
    ' ★コンパイルエラー修正：変数をDate型からVariant型に変更
    Dim d As Variant
    Dim tempDate As Date
    
    '--- Step 1: 振替休日を除く「本来の祝日」をすべてリストアップ ---
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
        If d > 0 Then ' 日付が有効な場合のみ
            If Not hList.Exists(CDate(d)) Then
                hList.Add CDate(d), GetPrimaryHolidayName(CDate(d))
            End If
        End If
    Next d

    '--- Step 2: 振替休日を追加 ---
    Dim originalDates As Variant
    originalDates = hList.Keys
    
    For Each d In originalDates
        If Weekday(d) = vbSunday Then
            tempDate = d + 1
            Do While hList.Exists(tempDate)
                tempDate = tempDate + 1
            Loop
            hList.Add tempDate, "振替休日"
        End If
    Next d
    
    '--- Step 3: 国民の休日を追加 ---
    originalDates = hList.Keys
    For Each d In originalDates
        If hList.Exists(d + 2) And Not hList.Exists(d + 1) Then
            hList.Add d + 1, "国民の休日"
        End If
    Next d

    Set GenerateHolidayList = hList
End Function


Private Function GetPrimaryHolidayName(ByVal d As Date) As String
    Dim y As Long: y = Year(d)
    Dim m As Long: m = Month(d)
    Dim n As Long: n = Day(d)
    
    ' ★ご要望に基づき、IF文を改行形式に整形
    Select Case m
        Case 1
            If n = 1 Then
                GetPrimaryHolidayName = "元日"
            Else
                GetPrimaryHolidayName = "成人の日"
            End If
        Case 2
            If n = 11 Then
                GetPrimaryHolidayName = "建国記念の日"
            Else
                GetPrimaryHolidayName = "天皇誕生日"
            End If
        Case 3
            GetPrimaryHolidayName = "春分の日"
        Case 4
            GetPrimaryHolidayName = "昭和の日"
        Case 5
            If n = 3 Then
                GetPrimaryHolidayName = "憲法記念日"
            ElseIf n = 4 Then
                GetPrimaryHolidayName = "みどりの日"
            ElseIf n = 5 Then
                GetPrimaryHolidayName = "こどもの日"
            End If
        Case 7
            GetPrimaryHolidayName = "海の日"
        Case 8
            GetPrimaryHolidayName = "山の日"
        Case 9
            If n = GetShubun(y) Then
                GetPrimaryHolidayName = "秋分の日"
            Else
                GetPrimaryHolidayName = "敬老の日"
            End If
        Case 10
            GetPrimaryHolidayName = "スポーツの日"
        Case 11
            If n = 3 Then
                GetPrimaryHolidayName = "文化の日"
            Else
                GetPrimaryHolidayName = "勤労感謝の日"
            End If
        Case 12
            GetPrimaryHolidayName = "天皇誕生日"
    End Select
End Function


Private Function GetHappyMonday(ByVal y As Long, ByVal m As Long, ByVal weekNum As Long) As Date
    GetHappyMonday = DateSerial(y, m, (weekNum - 1) * 7 + 1 + (8 - Weekday(DateSerial(y, m, 1), vbMonday)) Mod 7)
End Function

Private Function GetShunbun(ByVal y As Long) As Date
    Dim d As Integer
    d = Int(Val("20.8431") + Val("0.242194") * (y - 1980) - Int((y - 1980) / 4))
    GetShunbun = DateSerial(y, 3, d)
End Function

Private Function GetShubun(ByVal y As Long) As Date
    Dim d As Integer
    d = Int(Val("23.2488") + Val("0.242194") * (y - 1980) - Int((y - 1980) / 4))
    GetShubun = DateSerial(y, 9, d)
End Function

