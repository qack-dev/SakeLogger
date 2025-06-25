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
    Dim sakeName As String
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

    If CalcAlcoholInfo(sakeName, nowWeight, drankWeight, pureAlcohol) Then
        ' 結果を表示（小数点1桁）
        lblResult.Caption = "純アルコール量：" & Format(pureAlcohol, "0.0") & " g"
    End If
End Sub

Private Sub cmbSake_Change()
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
    Dim drankWeight As Double, pureAlcohol As Double
    Dim prevWeight As Double, nowWeight As Double
    Dim lastLogRow As Long

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

    If Not CalcAlcoholInfo(sakeName, nowWeight, drankWeight, pureAlcohol) Then
        Exit Sub
    End If

    lastLogRow = wsLog.Cells(wsLog.Rows.Count, logIdCol).End(xlUp).Row + 1

    ' --- 書式設定の変更 ---
    wsLog.Cells(lastLogRow, logNowCol).NumberFormat = "0.0"
    wsLog.Cells(lastLogRow, logPureAlcCol).NumberFormat = "0.0"
    wsLog.Cells(lastLogRow, logDrunkCol).NumberFormat = "0.0"

    ' --- ログに記録する ---
    wsLog.Cells(lastLogRow, logDateCol).Value = txtDate.Value           ' 日時
    wsLog.Cells(lastLogRow, logNameCol).Value = sakeName                ' 酒名
    wsLog.Cells(lastLogRow, logNowCol).Value = nowWeight               ' 現在重量
    wsLog.Cells(lastLogRow, logPureAlcCol).Value = Round(pureAlcohol, 1)   ' 純アル量(g)
    wsLog.Cells(lastLogRow, logDrunkCol).Value = Round(drankWeight, 1)   ' 飲んだ量(g)
    wsLog.Cells(lastLogRow, logIdCol).Value = lastLogRow
    
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
