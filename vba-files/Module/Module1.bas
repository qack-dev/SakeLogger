Attribute VB_Name = "Module1"
'グローバル変数
Public ws As Worksheet

Option Explicit

Sub ShowUserForm()

    UserForm1.Show




End Sub

'オブジェクト変数代入
Public Sub setObj()
    Set ws = ThisWorkbook.Worksheets("お酒マスタ")
End Sub
'オブジェクト変数開放
Public Sub releaseObj()
    Set ws = Nothing
End Sub
