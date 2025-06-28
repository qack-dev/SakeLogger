Attribute VB_Name = "CalcHoliday"
Option Explicit

'================================================================================
' メインプロシージャ：指定した年の祝日をシートに追記する
'================================================================================
Sub AddHolidaysForYear()
    Dim yearVal As Long
    Dim wsMaster As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim holidayDate As Date
    Dim holidayName As String
    Dim checkRange As Range
    
    '--- ユーザーに入力を促す ---
    On Error Resume Next
    yearVal = Application.InputBox("祝日リストに追加したい年（西暦）を入力してください。", "年 指定", Year(Date))
    If yearVal = 0 Then Exit Sub ' キャンセルされた場合
    On Error GoTo 0
    
    '--- 初期設定 ---
    Set wsMaster = ThisWorkbook.Sheets("祝日マスタ") ' ご自身のシート名に合わせてください
    
    '--- その年の祝日が既にリストにないかチェック（簡易版）---
    If Application.WorksheetFunction.CountIf(wsMaster.Columns("A"), ">=" & DateSerial(yearVal, 1, 1)) > 0 And _
       Application.WorksheetFunction.CountIf(wsMaster.Columns("A"), "<=" & DateSerial(yearVal, 12, 31)) > 0 Then
        If MsgBox(yearVal & "年の祝日は既に追加されている可能性があります。続行しますか？", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If

    '--- 祝日を計算し、シートに書き出す ---
    Application.ScreenUpdating = False
    
    For i = 1 To 12 ' 1月から12月まで
        holidayName = GetHolidayName(DateSerial(yearVal, i, 1), holidayDate) ' 月初で代表チェック
    Next i
    
    ' 1年分の日付をループして祝日判定と書き込み
    For i = 1 To 366
        holidayDate = DateSerial(yearVal, 1, i)
        If Year(holidayDate) <> yearVal Then Exit For ' 年が変わったら終了
        
        holidayName = GetHolidayName(holidayDate)
        
        If holidayName <> "" Then
            ' 既に同じ日付がないか確認
            Set checkRange = wsMaster.Columns("A").Find(What:=holidayDate, LookIn:=xlFormulas, LookAt:=xlWhole)
            
            If checkRange Is Nothing Then ' 日付がなければ追記
                lastRow = wsMaster.Cells(wsMaster.Rows.Count, "A").End(xlUp).Row + 1
                wsMaster.Cells(lastRow, "A").Value = holidayDate
                wsMaster.Cells(lastRow, "B").Value = holidayName
                wsMaster.Cells(lastRow, "C").Value = yearVal & "年 自動計算"
            End If
        End If
    Next i
    
    ' シートを日付順に並べ替え
    With wsMaster.Sort
        .SortFields.Clear
        .SortFields.Add key:=wsMaster.Range("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange wsMaster.UsedRange
        .Header = xlYes ' 1行目が見出しの場合
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Application.ScreenUpdating = True
    MsgBox yearVal & "年の祝日を追加しました。法改正による変更は手動で修正してください。"
End Sub


'================================================================================
' 補助関数群
'================================================================================

' 指定された日付の祝日名を返す（祝日でなければ空文字）
Private Function GetHolidayName(d As Date, Optional ByRef holidayDate As Date) As String
    Dim y As Long, m As Long, n As Long
    y = Year(d)
    m = Month(d)
    n = Day(d)
    Dim wd As Long
    wd = Weekday(d)
    
    ' 2000年以降を対象とする
    If y < 2000 Then GetHolidayName = "": Exit Function
    
    Dim resultName As String
    holidayDate = d ' 初期化
    
    '--- 固定祝日 ---
    Select Case m
        Case 1: If n = 1 Then resultName = "元日"
        Case 2: If n = 11 Then resultName = "建国記念の日"
        Case 4: If n = 29 Then resultName = "昭和の日"
        Case 5:
            If n = 3 Then
                resultName = "憲法記念日"
            ElseIf n = 4 Then
                resultName = "みどりの日"
            ElseIf n = 5 Then
                resultName = "こどもの日"
            End If
        Case 8: If n = 11 Then resultName = "山の日"
        Case 11:
            If n = 3 Then
                resultName = "文化の日"
            ElseIf n = 23 Then
                resultName = "勤労感謝の日"
            End If
        Case 12: If y >= 2019 And n = 23 Then resultName = "" ' 上皇誕生日
                 If y >= 2020 And n = 23 Then resultName = "" ' 上皇誕生日
                 If y <= 2018 And n = 23 Then resultName = "天皇誕生日"
                 If y >= 2020 And Day(GetHolidayName(DateSerial(y, 2, 23))) = 23 Then resultName = "天皇誕生日" '2/23が祝日
    End Select

    ' 天皇誕生日 (2020年以降)
    If y >= 2020 And m = 2 And n = 23 Then resultName = "天皇誕生日"

    '--- ハッピーマンデー ---
    If GetHappyMonday(y, m, 2) = n And m = 1 Then resultName = "成人の日"
    If GetHappyMonday(y, m, 3) = n And m = 7 Then resultName = "海の日"
    If GetHappyMonday(y, m, 3) = n And m = 9 Then resultName = "敬老の日"
    If GetHappyMonday(y, m, 2) = n And m = 10 Then resultName = "スポーツの日" '体育の日から改称

    '--- 春分の日・秋分の日 (簡易計算式) ---
    If n = CInt(20.8431 + 0.242194 * (y - 1980) - CInt((y - 1980) / 4)) And m = 3 Then resultName = "春分の日"
    If n = CInt(23.2488 + 0.242194 * (y - 1980) - CInt((y - 1980) / 4)) And m = 9 Then resultName = "秋分の日"

    '--- 振替休日 ---
    If resultName = "" And wd = 2 Then '月曜日で祝日でない場合
        If GetHolidayName(d - 1) <> "" Then '前日が祝日の場合
            resultName = "振替休日"
        End If
    End If
    
    '--- 国民の休日 ---
    If resultName = "" And GetHolidayName(d - 1) <> "" And GetHolidayName(d + 1) <> "" Then
        resultName = "国民の休日"
    End If

    GetHolidayName = resultName
End Function

' 第n月曜日の日付を返す
Private Function GetHappyMonday(y As Long, m As Long, weekNum As Long) As Long
    Dim firstDay As Date
    firstDay = DateSerial(y, m, 1)
    GetHappyMonday = ((weekNum - 1) * 7) + 1 + (9 - Weekday(firstDay, vbMonday)) Mod 7
End Function

