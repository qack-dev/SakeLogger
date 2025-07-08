Attribute VB_Name = "M_SakeLogics"
Option Explicit

' 飲んだ量と純アルコール量を計算
Public Function CalculateAlcoholInfo(ByVal sakeName As String, ByVal currentWeight As Double, ByRef drankWeight As Double, ByRef pureAlcohol As Double) As Boolean
    Dim abv As Double, fullWeight As Double, emptyWeight As Double
    Dim previousWeight As Double
    Dim i As Long
    Dim isFound As Boolean
    Dim masterSheet As Worksheet
    Dim logSheet As Worksheet

    On Error GoTo ErrorHandler

    Set masterSheet = M_SakeForm.GetMasterSheet()
    Set logSheet = M_SakeForm.GetLogSheet()

    ' マスターから度数・重量情報を取得
    isFound = False
    For i = 2 To M_SakeForm.GetLastMasterDataCell().Row
        If masterSheet.Cells(i, COL_MASTER_ID).Value & "." & masterSheet.Cells(i, COL_MASTER_NAME).Value = sakeName Then
            abv = masterSheet.Cells(i, COL_MASTER_ALCOHOL).Value
            fullWeight = masterSheet.Cells(i, COL_MASTER_FULL_WEIGHT).Value
            
            If IsEmpty(masterSheet.Cells(i, COL_MASTER_EMPTY_WEIGHT).Value) Or masterSheet.Cells(i, COL_MASTER_EMPTY_WEIGHT).Value = "" Then
                MsgBox "注意: このお酒は空き容器重量が未登録です。空き容器重量を0gとして計算を続行します。", vbInformation
                emptyWeight = 0
            Else
                emptyWeight = masterSheet.Cells(i, COL_MASTER_EMPTY_WEIGHT).Value
            End If

            If currentWeight > fullWeight Or currentWeight < emptyWeight Then
                MsgBox "現在の重量の値が不正です（満タン時重量を超えているか、空き容器重量を下回っています）。", vbExclamation
                CalculateAlcoholInfo = False
                Exit Function
            End If
            
            isFound = True
            Exit For
        End If
    Next i

    If Not isFound Then
        MsgBox "お酒マスターに該当するお酒が見つかりませんでした。", vbCritical
        CalculateAlcoholInfo = False
        Exit Function
    End If

    ' 飲んだ量を計算
    If frmSakeLogger.optNewOpen.Value Then
        drankWeight = fullWeight - currentWeight
    ElseIf frmSakeLogger.optContinued.Value Then
        previousWeight = GetPreviousWeight(sakeName, logSheet)
        If previousWeight = -1 Then
            MsgBox "このお酒の直前の記録が見つかりませんでした。「新規開封」を選択してください。", vbExclamation
            CalculateAlcoholInfo = False
            Exit Function
        End If
        drankWeight = previousWeight - currentWeight
    Else
        MsgBox "「新規開封」または「継続」を選択してください。", vbExclamation
        CalculateAlcoholInfo = False
        Exit Function
    End If

    ' 純アルコール量を計算 (密度: 0.8)
    pureAlcohol = drankWeight * (abv / 100) * 0.8

    CalculateAlcoholInfo = True
    Exit Function

ErrorHandler:
    MsgBox "計算中にエラーが発生しました。" & vbCrLf & Err.Description, vbCritical
    CalculateAlcoholInfo = False
End Function

' 指定されたお酒の直前の重量を取得する
Public Function GetPreviousWeight(ByVal sakeName As String, ByVal logSheet As Worksheet) As Double
    Dim lastRow As Long
    Dim i As Long

    GetPreviousWeight = -1 ' 見つからなかった場合のデフォルト値
    lastRow = logSheet.Cells(logSheet.Rows.Count, COL_LOG_ID).End(xlUp).Row

    For i = lastRow To 2 Step -1
        If logSheet.Cells(i, COL_LOG_NAME).Value = sakeName Then
            GetPreviousWeight = logSheet.Cells(i, COL_LOG_CURRENT_WEIGHT).Value
            Exit Function
        End If
    Next i
End Function

' 日付文字列が 'yyyy/mm/dd' 形式か検証する
Public Function IsValidDateFormat(ByVal dateString As String) As Boolean
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = "^\d{4}/\d{2}/\d{2}$"
    
    If regEx.Test(dateString) And IsDate(dateString) Then
        IsValidDateFormat = True
    Else
        IsValidDateFormat = False
    End If
End Function

' フォームからの入力に基づいて現在の重量を計算する共通関数
Public Function CalculateCurrentWeightFromInput( _
    ByVal frm As Object, _
    ByVal fullWeight As Double, _
    ByVal emptyWeight As Double, _
    ByVal previousWeight As Double, _
    ByRef currentWeight As Double _
) As Boolean
    Dim inputCount As Integer
    Dim tempWeight As Double
    
    On Error GoTo ErrorHandler
    
    inputCount = 0
    If Trim(frm.txtNowWeight.Value) <> "" Then inputCount = inputCount + 1
    If Trim(frm.txtNowPercent.Value) <> "" Then inputCount = inputCount + 1
    If Trim(frm.txtNowNum.Value) <> "" Then inputCount = inputCount + 1
    
    ' 1. 入力ソースの特定とバリデーション
    If inputCount = 0 Then
        MsgBox "現在の重量、残量(%)、または杯数を入力してください。", vbExclamation
        CalculateCurrentWeightFromInput = False
        Exit Function
    ElseIf inputCount > 1 Then
        MsgBox "入力箇所は1つだけにしてください。", vbExclamation
        CalculateCurrentWeightFromInput = False
        Exit Function
    End If
    
    ' 3. 計算ロジック
    If Trim(frm.txtNowWeight.Value) <> "" Then
        If Not IsNumeric(frm.txtNowWeight.Value) Then
            MsgBox "現在の重量は数値を入力してください。", vbExclamation
            CalculateCurrentWeightFromInput = False
            Exit Function
        End If
        currentWeight = CDbl(frm.txtNowWeight.Value)
    ElseIf Trim(frm.txtNowPercent.Value) <> "" Then
        If Not IsNumeric(frm.txtNowPercent.Value) Then
            MsgBox "残量(%)は数値を入力してください。", vbExclamation
            CalculateCurrentWeightFromInput = False
            Exit Function
        End If
        tempWeight = CDbl(frm.txtNowPercent.Value)
        If tempWeight < 0 Or tempWeight > 100 Then
            MsgBox "残量(%)は0から100の範囲で入力してください。", vbExclamation
            CalculateCurrentWeightFromInput = False
            Exit Function
        End If
        currentWeight = emptyWeight + ((fullWeight - emptyWeight) * (tempWeight / 100))
    ElseIf Trim(frm.txtNowNum.Value) <> "" Then
        If Not IsNumeric(frm.txtNowNum.Value) Then
            MsgBox "杯数は数値を入力してください。", vbExclamation
            CalculateCurrentWeightFromInput = False
            Exit Function
        End If
        If previousWeight = -1 Then ' previousWeightが取得できなかった場合
            MsgBox "継続記録のデータが見つかりません。新規開封を選択するか、別の入力方法を使用してください。", vbExclamation
            CalculateCurrentWeightFromInput = False
            Exit Function
        End If
        
        Dim drankAmount As Double
        ' 開発者メモ: この「飲んだ量」の計算式は、内容量全体に杯数を掛けるという少し特殊なロジックです。
        ' 例えば、txtNowNumが「1」の場合、内容量全体を飲んだとみなされます。
        ' もし「1杯あたりの量」が定義されている場合は、その量とtxtNowNumを掛ける方が直感的かもしれません。
        ' 現在のロジックは、txtNowNumが「割合」として機能しているように見えます。
        drankAmount = (fullWeight - emptyWeight) * CDbl(frm.txtNowNum.Value)
        currentWeight = previousWeight - drankAmount
    End If
    
    CalculateCurrentWeightFromInput = True
    Exit Function

ErrorHandler:
    MsgBox "入力値の処理中にエラーが発生しました。" & vbCrLf & Err.Description, vbCritical
    CalculateCurrentWeightFromInput = False
End Function
