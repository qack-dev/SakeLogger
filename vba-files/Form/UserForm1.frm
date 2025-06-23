VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "酒量を登録"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6060
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'グローバル定数
Private Const idCol As Integer = 1 'ID列
Private Const nameCol As Integer = 2 'お酒の名前列
Private Const kindsCol As Integer = 3 '種類列
Private Const alcoholCol As Integer = 4 '度数列
Private Const fullCol As Integer = 5 '未開封重量列
Private Const empCol As Integer = 6 '空重量列

Option Explicit

Private Sub btnCalc_Click()
    Dim i As Long, lastRow As Long
    Dim sakeName As String
    Dim abv As Double, fullWeight As Double, emptyWeight As Double
    Dim nowWeight As Double, drankWeight As Double, pureAlcohol As Double

    Call setObj
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

    lastRow = ws.Cells(ws.Rows.Count, nameCol).End(xlUp).Row
    For i = 2 To lastRow
        If ws.Cells(i, nameCol).Value = sakeName Then
            ' 必要な情報を取得
            abv = ws.Cells(i, alcoholCol).Value         ' 度数
            fullWeight = ws.Cells(i, fullCol).Value  ' 未開封重量

            If ws.Cells(i, empCol).Value = "" Then
                MsgBox "この酒は空ボトル重量が未登録です。" & vbCrLf & _
                       "飲み終えたら空ボトル重量を入力してください。", vbExclamation
                Exit Sub
            End If

            emptyWeight = ws.Cells(i, empCol).Value ' 空ボトル重量

            ' 入力チェック
            If nowWeight > fullWeight Or nowWeight < emptyWeight Then
                MsgBox "現在の重さが不正です。", vbExclamation
                Exit Sub
            End If

            ' 飲んだ重量計算
            drankWeight = fullWeight - nowWeight

            ' 純アルコール量計算（アルコールの比重 = 0.8）
            pureAlcohol = drankWeight * (abv / 100) * 0.8

            ' 結果を表示（小数点1桁）
            lblResult.Caption = "純アルコール量: " & Format(pureAlcohol, "0.0") & " g"
            Exit Sub
        End If
    Next i
    Call releaseObj
End Sub

Private Sub cmbSake_Change()
    Dim targetRow As Long
    Dim lastRow As Long
    Dim i As Long
    Call setObj

    lastRow = ws.Cells(ws.Rows.Count, nameCol).End(xlUp).Row

    For i = 2 To lastRow
        If ws.Cells(i, idCol).Value & "." & ws.Cells(i, nameCol).Value = cmbSake.Value Then
            ' 対象の行が見つかったら情報を取得
            lblABV.Caption = "度数：" & ws.Cells(i, alcoholCol).Value & " %"
            lblFullW.Caption = "未開封重量：" & ws.Cells(i, fullCol).Value & " g"
            
            If ws.Cells(i, empCol).Value = "" Then
                lblEmptyW.Caption = "空ボトル重量：未登録"
                lblAlert.Caption = "!!!飲み終わったら空ボトル重量を入力してください!!!"
            Else
                lblEmptyW.Caption = "空ボトル重量：" & ws.Cells(i, empCol).Value & " g"
                lblAlert.Caption = ""
            End If
            
            Exit For
        End If
    Next i
    Call releaseObj
End Sub

Private Sub UserForm_Initialize()
    '変数宣言
    Dim i As Integer
    Dim lastCell As Range
    Dim dataRange As Range
    Call setObj
    'ws.Rows.Count でシートの最大行数を取得し、そこから End(xlUp) でデータのある最終セルを探す。
    Set lastCell = ws.Cells(ws.Rows.Count, nameCol).End(xlUp)

    'データが2行目以降に存在する場合のみ処理を実行
    If lastCell.Row >= 2 Then
        '2行目から最終行までの範囲を取得
        'Set dataRange = ws.Range(Cells(2, nameCol), lastCell)
        For i = 2 To lastCell.Row
            cmbSake.AddItem ws.Cells(i, idCol).Value & "." & ws.Cells(i, nameCol).Value
        Next i
        'ComboBoxのリストに範囲の値を設定
        'Me.cmbSake.List = dataRange.Value
    End If
    Call releaseObj
End Sub
