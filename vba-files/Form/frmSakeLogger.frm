VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSakeLogger 
   Caption         =   "酒量を登録"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7530
   OleObjectBlob   =   "frmSakeLogger.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSakeLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCalc_Click()
    Dim i As Long, lastRow As Long
    Dim sakeName As String
    Dim abv As Double, fullWeight As Double, emptyWeight As Double
    Dim nowWeight As Double, drankWeight As Double, pureAlcohol As Double

    sakeName = cmbSake.Value

    If sakeName = "" Then
        MsgBox "お酒を選択してください", vbExclamation
        Exit Sub
    End If

    If Not IsNumeric(txtNowWeight.Value) Then
        MsgBox "現在の重さを正しく入力してください", vbExclamation
        Exit Sub
    End If

    nowWeight = CDbl(txtNowWeight.Value)

    lastRow = lastCell.Row
    For i = 2 To lastRow
        If wsMaster.Cells(i, idCol).Value & "." & wsMaster.Cells(i, nameCol).Value = sakeName Then
            ' 必要な情報を取得
            abv = wsMaster.Cells(i, alcoholCol).Value         ' 度数
            fullWeight = wsMaster.Cells(i, fullCol).Value  ' 未開封重量

            If wsMaster.Cells(i, empCol).Value = "" Then
                MsgBox "この酒は空ボトル重量が未登録です。" & vbCrLf & _
                       "飲み終えたら空ボトル重量を入力してください。", vbExclamation
            Else
                emptyWeight = wsMaster.Cells(i, empCol).Value ' 空ボトル重量
                ' 入力チェック
                If nowWeight > fullWeight Or nowWeight < emptyWeight Then
                    MsgBox "現在の重さが不正です。", vbExclamation
                    Exit Sub
                End If
            End If
            ' 飲んだ重量計算
            drankWeight = fullWeight - nowWeight
            ' 純アルコール量計算（アルコールの比重 = 0.8）
            pureAlcohol = drankWeight * (abv / 100) * 0.8
            ' 結果を表示（小数点1桁）
            lblResult.Caption = "純アルコール量：" & Format(pureAlcohol, "0.0") & " g"
            Exit Sub
        End If
    Next i
End Sub

Private Sub cmbSake_Change()
    Dim targetRow As Long
    Dim lastRow As Long
    Dim i As Long
    
    lastRow = lastCell.Row
    
    With wsMaster
        For i = 2 To lastRow
            If .Cells(i, idCol).Value & "." & .Cells(i, nameCol).Value = cmbSake.Value Then
                ' 対象の行が見つかったら情報を取得
                lblABV.Caption = "度数：" & .Cells(i, alcoholCol).Value & " %"
                lblFullW.Caption = "未開封重量：" & .Cells(i, fullCol).Value & " g"
                
                If .Cells(i, empCol).Value = "" Then
                    lblEmptyW.Caption = "空ボトル重量：未登録"
                    lblAlert.Caption = "!!!飲み終わったら空ボトル重量を入力してください!!!"
                Else
                    lblEmptyW.Caption = "空ボトル重量：" & .Cells(i, empCol).Value & " g"
                    lblAlert.Caption = ""
                End If
                
                Exit For
            End If
        Next i
    End With
End Sub

Private Sub btnSave_Click()
    Dim sakeName As String
    Dim abv As Double, fullWeight As Double, emptyWeight As Double
    Dim nowWeight As Double, drankWeight As Double, pureAlcohol As Double
    Dim i As Long
    Dim iLog As Long
    Dim lastLogRow As Long
    Dim prevWeight As Double
    Dim found As Boolean

    ' --- 入力チェック ---
    If cmbSake.Value = "" Then
        MsgBox "酒を選んでください", vbExclamation
        Exit Sub
    End If

    If Not IsYyyyMmDdFormat_RegEx(txtDate.Value) Then
        MsgBox "飲んだ日には'yyyy/mm/dd'形式で入力してください", vbExclamation
        Exit Sub
    End If

    If Not IsNumeric(txtNowWeight.Value) Then
        MsgBox "現在の重さを入力してください", vbExclamation
        Exit Sub
    End If

    sakeName = cmbSake.Value
    nowWeight = CDbl(txtNowWeight.Value)

    ' --- マスタから情報取得 ---
    For i = 2 To lastCell.Row
        If wsMaster.Cells(i, idCol).Value & "." & wsMaster.Cells(i, nameCol).Value = sakeName Then
            abv = wsMaster.Cells(i, alcoholCol).Value
            fullWeight = wsMaster.Cells(i, fullCol).Value
            If wsMaster.Cells(i, empCol).Value = "" Then
                MsgBox "空ボトル重量が未入力です", vbExclamation
            Else
                emptyWeight = wsMaster.Cells(i, empCol).Value ' 空ボトル重量
                ' 入力チェック
                If nowWeight > fullWeight Or nowWeight < emptyWeight Then
                    MsgBox "現在の重さが不正です。", vbExclamation
                    Exit Sub
                End If
            End If
        End If
    Next i

    ' --- 飲んだ量・純アル量の計算方法を分岐 ---
    If optNewOpen.Value = True Then
        ' 新品を開けた場合：未開封重量から計算
        drankWeight = fullWeight - nowWeight
    ElseIf optContinued.Value = True Then
        ' 継続飲用：過去の記録から直近の重量を引く
        found = False
        lastLogRow = wsLog.Cells(wsLog.Rows.Count, logIdCol).End(xlUp).Row

        ' 下から上に遡って同じ酒の直近の重量を探す
        For iLog = lastLogRow To 2 Step -1
            If wsLog.Cells(iLog, logNameCol).Value = sakeName Then
                prevWeight = wsLog.Cells(iLog, logNowCol).Value
                found = True
                Exit For
            End If
        Next iLog

        If Not found Then
            MsgBox "このお酒の記録がまだ存在しません。" & vbCrLf & _
                   "『新品を開けた』を選んでください。", vbExclamation
            Exit Sub
        End If

        drankWeight = prevWeight - nowWeight
    Else
        MsgBox "新品か継続かを選んでください。", vbExclamation
        Exit Sub
    End If

    ' --- 純アルコール量の計算（共通） ---
    pureAlcohol = drankWeight * (abv / 100) * 0.8

    ' --- 書式設定の変更 ---
    wsLog.Cells(lastLogRow + 1, logNowCol).NumberFormat = "0.0"
    wsLog.Cells(lastLogRow + 1, logPureAlcCol).NumberFormat = "0.0"
    wsLog.Cells(lastLogRow + 1, logDrunkCol).NumberFormat = "0.0"

    ' --- ログに記録する ---
    wsLog.Cells(lastLogRow + 1, logDateCol).Value = txtDate.Value          ' 日時
    wsLog.Cells(lastLogRow + 1, logNameCol).Value = sakeName               ' 酒名
    wsLog.Cells(lastLogRow + 1, logNowCol).Value = nowWeight              ' 現在重量
    wsLog.Cells(lastLogRow + 1, logPureAlcCol).Value = Round(pureAlcohol, 1)  ' 純アル量(g)
    wsLog.Cells(lastLogRow + 1, logDrunkCol).Value = Round(drankWeight, 1)  ' 飲んだ量(g)
    wsLog.Cells(lastLogRow + 1, logIdCol).Value = lastLogRow
    
    MsgBox "記録を保存しました！", vbInformation

    ' --- 入力欄をリセット（任意） ---
    txtNowWeight.Value = ""
    lblResult.Caption = ""
End Sub

Private Sub UserForm_Initialize()
    '変数宣言
    Dim i As Long
    'データが2行目以降に存在する場合のみ処理を実行
    If lastCell.Row >= 2 Then
        '2行目から最終行までの範囲を取得
        For i = 2 To lastCell.Row
            cmbSake.AddItem wsMaster.Cells(i, idCol).Value & "." & wsMaster.Cells(i, nameCol).Value
        Next i
    End If
End Sub
