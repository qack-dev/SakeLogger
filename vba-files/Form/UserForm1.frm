VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'グローバル定数
Private Const inputCol As Integer = 2 'お酒の名前列

Option Explicit

Private Sub UserForm_Initialize()
    '変数宣言
    Dim ws As Worksheet
    Dim lastCell As Range
    Dim dataRange As Range
    'オブジェクト格納
    Set ws = ThisWorkbook.Worksheets("お酒マスタ")
    'ws.Rows.Count でシートの最大行数を取得し、そこから End(xlUp) でデータのある最終セルを探す。
    Set lastCell = ws.Cells(ws.Rows.Count, inputCol).End(xlUp)

    'データが2行目以降に存在する場合のみ処理を実行
    If lastCell.Row >= 2 Then
        '2行目から最終行までの範囲を取得
        Set dataRange = ws.Range(Cells(2, inputCol), lastCell)
        'ComboBoxのリストに範囲の値を設定
        Me.ComboBox1.List = dataRange.Value
    End If
End Sub
