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


