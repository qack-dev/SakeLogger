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
    Set wsLog = Sheets("飲酒記録")
    'wsMaster.RowsMaster.Count でシートの最大行数を取得し、そこから End(xlUp) でデータのある最終セルを探す。
    Set lastCell = wsMaster.Cells(wsMaster.Rows.Count, nameCol).End(xlUp)
End Sub
'オブジェクト変数開放
Public Sub releaseObj()
    Set wsMaster = Nothing
    Set lastCell = Nothing
    Set wsLog = Nothing
End Sub

Public Sub ShowUserForm()
    Call setObj
    wsMaster.Activate
    UserForm1.Show
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
