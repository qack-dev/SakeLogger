Attribute VB_Name = "Module1"
'グローバル変数
Public wsMaster As Worksheet
Public wsLog As Worksheet
Public lastCell As Range

'グローバル定数
'お酒マスタシート
Public Const idCol As Integer = 1 'ID列
Public Const nameCol As Integer = 2 'お酒の名前列
Public Const kindsCol As Integer = 3 '種類列
Public Const alcoholCol As Integer = 4 '度数列
Public Const fullCol As Integer = 5 '未開封重量列
Public Const empCol As Integer = 6 '空重量列
'飲酒記録シート
Public Const logDateCol As Integer = 1 '日時列
Public Const logNameCol As Integer = 2 'お酒の名前列
Public Const logNowCol As Integer = 3 '現在重量列
Public Const logPureAlcCol As Integer = 4 '純アル量列
Public Const logDrunkCol As Integer = 5 '飲んだ量列
Public Const logComCol As Integer = 6 'コメント列
Public Const logIdCol As Integer = 7 'このシートのIDの列
'集計シート
Public Const sumDateCol As Integer = 1 '日時列
Public Const sumPureAlcCol As Integer = 2 '純アル量列

Option Explicit

'オブジェクト変数代入
Public Sub setObj()
    Set wsMaster = ThisWorkbook.Worksheets("お酒マスタ")
    Set wsLog = ThisWorkbook.Worksheets("飲酒記録")
    'wsMaster.RowsMaster.Count でシートの最大行数を取得し、そこから End(xlUp) でデータのある最終セルを探す。
    Set lastCell = wsMaster.Cells(wsMaster.Rows.Count, nameCol).End(xlUp)
End Sub
'オブジェクト変数開放
Public Sub releaseObj()
    Set wsMaster = Nothing
    Set wsLog = Nothing
    Set lastCell = Nothing
End Sub

Public Sub ShowUserForm()
    Call setObj
    wsMaster.Activate
    frmSakeLogger.Show
    Call releaseObj



End Sub

'正規表現で'yyyy/mm/dd'形式をチェックし、かつ日付として妥当か判定
Public Function IsYyyyMmDdFormat_RegEx(ByVal target As String) As Boolean
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")

    ' パターンを設定
    ' ^         : 文字列の先頭
    ' \d{4}     : 4桁の数字 (年)
    ' /         : 区切り文字のスラッシュ
    ' \d{2}     : 2桁の数字 (月)
    ' /         : 区切り文字のスラッシュ
    ' \d{2}     : 2桁の数字 (日)
    ' $         : 文字列の末尾
    regEx.Pattern = "^\d{4}/\d{2}/\d{2}$"

    If regEx.Test(target) And IsDate(target) Then
        IsYyyyMmDdFormat_RegEx = True
    Else
        IsYyyyMmDdFormat_RegEx = False
    End If
End Function

'飲んだ量(drankWeight)と純アルコール量(pureAlcohol)の計算
Public Function CalcAlcoholInfo(sakeName As String, nowWeight As Double, _
                         ByRef drankWeight As Double, ByRef pureAlcohol As Double) As Boolean
    Dim abv As Double, fullWeight As Double, emptyWeight As Double
    Dim prevWeight As Double
    Dim i As Long, lastRow As Long
    Dim found As Boolean

    On Error GoTo ErrHandler
    
    ' マスタから度数(ABV)・重量取得
    found = False
    For i = 2 To lastCell.Row
        If wsMaster.Cells(i, idCol).Value & "." & wsMaster.Cells(i, nameCol).Value = sakeName Then
            abv = wsMaster.Cells(i, alcoholCol).Value     ' 度数
            fullWeight = wsMaster.Cells(i, fullCol).Value ' 未開封重量
            If wsMaster.Cells(i, empCol).Value = "" Then
                MsgBox "この酒は空ボトル重量が未登録です。" & vbCrLf & _
                       "飲み終えたら空ボトル重量を入力してください。", vbExclamation
                CalcAlcoholInfo = False
            Else
                emptyWeight = wsMaster.Cells(i, empCol).Value ' 空ボトル重量
                ' 入力チェック
                If nowWeight > fullWeight Or nowWeight < emptyWeight Then
                    MsgBox "現在の重さが不正です。", vbExclamation
                    Exit Function
                End If
            End If
            found = True
            Exit For
        End If
    Next i

    If Not found Then
        MsgBox "お酒マスタにこのお酒が見つかりません。", vbCritical
        CalcAlcoholInfo = False
        Exit Function
    End If

    ' 分岐処理（OptionButtonの状態で）
    If frmSakeLogger.optNewOpen.Value = True Then
        ' 新品開封
        drankWeight = fullWeight - nowWeight

    ElseIf frmSakeLogger.optContinued.Value = True Then
        ' 継続飲用：前回重量取得
        lastRow = wsLog.Cells(wsLog.Rows.Count, logIdCol).End(xlUp).Row
        found = False
        For i = lastRow To 2 Step -1
            If wsLog.Cells(i, logNameCol).Value = sakeName Then
                prevWeight = wsLog.Cells(i, logNowCol).Value
                found = True
                Exit For
            End If
        Next i

        If Not found Then
            MsgBox "このお酒の記録がまだ存在しません。" & vbCrLf & _
                   "『新品を開けた』を選んでください。", vbExclamation
            CalcAlcoholInfo = False
            Exit Function
        End If

        drankWeight = prevWeight - nowWeight

    Else
        MsgBox "『新品を開けた』または『途中のお酒を飲んだ』を選んでください。", vbExclamation
        CalcAlcoholInfo = False
        Exit Function
    End If

    ' 純アルコール量計算
    pureAlcohol = drankWeight * (abv / 100) * 0.8

    CalcAlcoholInfo = True
    Exit Function

ErrHandler:
    MsgBox "計算中にエラーが発生しました。" & vbCrLf & Err.Description, vbCritical
    CalcAlcoholInfo = False
End Function

'集計シート更新
Public Sub updateTotallingSheet()
    '変数宣言
    Dim wsLog As Worksheet, wsSum As Worksheet
    Dim dict As Object
    Dim lastRow As Long, i As Long
    Dim dt As String, alcohol As Double

    Set wsLog = ThisWorkbook.Worksheets("飲酒記録")
    Set wsSum = ThisWorkbook.Worksheets("集計")
    Set dict = CreateObject("Scripting.Dictionary")

    ' 飲酒記録から集計
    lastRow = wsLog.Cells(wsLog.Rows.Count, logDateCol).End(xlUp).Row

    For i = 2 To lastRow
        dt = Format(wsLog.Cells(i, logDateCol).Value, "yyyy/mm/dd")
        alcohol = wsLog.Cells(i, logPureAlcCol).Value

        If dict.exists(dt) Then
            dict(dt) = dict(dt) + alcohol
        Else
            dict.Add dt, alcohol
        End If
    Next i

    ' 集計シートに書き出し
    wsSum.Cells.ClearContents
    wsSum.Range(Cells(1, sumDateCol), Cells(1, sumPureAlcCol)).Value = Array("日付", "純アルコール量")

    i = 2
    Dim key As Variant
    For Each key In dict.Keys
        wsSum.Cells(i, sumDateCol).Value = key
        wsSum.Cells(i, sumPureAlcCol).Value = Round(dict(key), 1)
        i = i + 1
    Next key

    ' 並び替え（昇順）
    'wsSum.Range("A2:B" & i - 1).Sort Key1:=wsSum.Range("A2"), Order1:=xlAscending, Header:=xlNo
    wsSum.Range(Cells(2, sumDateCol), Cells(i - 1, sumPureAlcCol)).Sort Key1:=wsSum.Cells(2, sumDateCol), Order1:=xlAscending, Header:=xlNo

    MsgBox "集計シートを更新しました", vbInformation
    ' オブジェクト開放
    Set wsLog = Nothing
    Set wsSum = Nothing
    Set dict = Nothing
End Sub

' グラフを自動で作成
Public Sub makeGraph()

    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim lastRow As Long

    Set ws = ThisWorkbook.Worksheets("集計")
    ws.Activate

    ' 最終行取得（列A = 日付）
    lastRow = ws.Cells(ws.Rows.Count, sumDateCol).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "集計データがありません。", vbExclamation
        Exit Sub
    End If

    ' すでにあるグラフを削除（再生成対応）
    For Each chartObj In ws.ChartObjects
        chartObj.Delete
    Next chartObj

    ' 新しいグラフオブジェクト作成
    Set chartObj = ws.ChartObjects.Add(Left:=300, Top:=20, Width:=500, Height:=300)

    With chartObj.Chart
        .ChartType = xlColumnClustered

        ' データ範囲を R1C1 形式で設定
        .SetSourceData Source:=ws.Range(ws.Cells(2, sumDateCol), ws.Cells(lastRow, sumPureAlcCol))

        ' 軸・タイトルなどの設定
        .HasTitle = True
        .ChartTitle.Text = "日別 純アルコール摂取量 (g)"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "日付"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "純アルコール量 (g)"
    End With

    MsgBox "グラフを作成しました！", vbInformation
    
    ' オブジェクト開放
    Set ws = Nothing
    Set chartObj = Nothing
End Sub

' 累計表示
Public Sub addTotalCell()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim formulaString As String
    
    '--- 書き出し先の列を定数化 ---
    Const totalCol As Integer = 6           ' 累計合計 (F列)
    Const monthlyTotalCol As Integer = 7    ' 今月の合計 (G列)
    Const helperStartDateCol As Integer = 8 ' ヘルパーセル:月の開始日 (H列)
    Const helperEndDateCol As Integer = 9   ' ヘルパーセル:月の終了日 (I列)
    '--------------------------------------------------------------------

    Set ws = ThisWorkbook.Worksheets("集計")
    ws.Activate

    ' 最終行の取得 (純アルコール量列を基準)
    lastRow = ws.Cells(ws.Rows.Count, sumPureAlcCol).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "集計が不足しているため、累計を表示できません。", vbExclamation
        Exit Sub
    End If

    ' タイトル行
    ws.Cells(1, totalCol).Value = "累計合計"
    ws.Cells(1, monthlyTotalCol).Value = "今月の合計"

    ' 累計合計の数式 (R1C1形式で設定)
    ws.Cells(2, totalCol).FormulaR1C1 = "=SUM(R2C" & sumPureAlcCol & ":R" & lastRow & "C" & sumPureAlcCol & ")"

    ' 今月1日と月末をヘルパーセルに表示
    ws.Cells(1, helperStartDateCol).FormulaR1C1 = "=DATE(YEAR(TODAY()), MONTH(TODAY()), 1)"
    ws.Cells(1, helperEndDateCol).FormulaR1C1 = "=EOMONTH(RC[-1], 0)"

    '--- 今月合計の数式を組み立てる ---
    formulaString = "=SUMIFS(" & _
        "R2C" & sumPureAlcCol & ":R" & lastRow & "C" & sumPureAlcCol & "," & _
        "R2C" & sumDateCol & ":R" & lastRow & "C" & sumDateCol & "," & _
        """>=""" & "&R1C" & helperStartDateCol & "," & _
        "R2C" & sumDateCol & ":R" & lastRow & "C" & sumDateCol & "," & _
        """<=""" & "&R1C" & helperEndDateCol & ")"

    ' 組み立てた数式をセルに設定
    ws.Cells(2, monthlyTotalCol).FormulaR1C1 = formulaString

    MsgBox "累計セルを追加しました！", vbInformation
    
    ' オブジェクト開放
    Set ws = Nothing
End Sub
