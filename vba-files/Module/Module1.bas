Attribute VB_Name = "Module1"
'グローバル変数
Public ws As Worksheet
Public lastCell As Range


'グローバル定数
Public Const idCol As Integer = 1 'ID列
Public Const nameCol As Integer = 2 'お酒の名前列
Public Const kindsCol As Integer = 3 '種類列
Public Const alcoholCol As Integer = 4 '度数列
Public Const fullCol As Integer = 5 '未開封重量列
Public Const empCol As Integer = 6 '空重量列

Option Explicit

Sub ShowUserForm()

    UserForm1.Show




End Sub

'オブジェクト変数代入
Public Sub setObj()
    Set ws = ThisWorkbook.Worksheets("お酒マスタ")
    'ws.Rows.Count でシートの最大行数を取得し、そこから End(xlUp) でデータのある最終セルを探す。
    Set lastCell = ws.Cells(ws.Rows.Count, nameCol).End(xlUp)
End Sub
'オブジェクト変数開放
Public Sub releaseObj()
    Set ws = Nothing
    Set lastCell = Nothing
End Sub
