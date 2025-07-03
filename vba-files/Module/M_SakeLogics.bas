Attribute VB_Name = "M_SakeLogics"
Option Explicit

' 飲んだ量と純アルコール量を計算する
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

    ' マスタから度数・重量情報を取得
    isFound = False
    For i = 2 To M_SakeForm.GetLastMasterDataCell().Row
        If masterSheet.Cells(i, COL_MASTER_ID).Value & "." & masterSheet.Cells(i, COL_MASTER_NAME).Value = sakeName Then
            abv = masterSheet.Cells(i, COL_MASTER_ALCOHOL).Value
            fullWeight = masterSheet.Cells(i, COL_MASTER_FULL_WEIGHT).Value
            
            If IsEmpty(masterSheet.Cells(i, COL_MASTER_EMPTY_WEIGHT).Value) Or masterSheet.Cells(i, COL_MASTER_EMPTY_WEIGHT).Value = "" Then
                MsgBox "注意: このお酒は空ボトル重量が未登録です。空重量を0gとして計算を続行します。", vbInformation
                emptyWeight = 0
            Else
                emptyWeight = masterSheet.Cells(i, COL_MASTER_EMPTY_WEIGHT).Value
            End If

            If currentWeight > fullWeight Or currentWeight < emptyWeight Then
                MsgBox "現在の重量の値が不正です（満タン重量を超えているか、空重量を下回っています）。", vbExclamation
                CalculateAlcoholInfo = False
                Exit Function
            End If
            
            isFound = True
            Exit For
        End If
    Next i

    If Not isFound Then
        MsgBox "お酒マスタに該当するお酒が見つかりません。", vbCritical
        CalculateAlcoholInfo = False
        Exit Function
    End If

    ' 飲んだ量を計算
    If frmSakeLogger.optNewOpen.Value Then
        drankWeight = fullWeight - currentWeight
    ElseIf frmSakeLogger.optContinued.Value Then
        previousWeight = GetPreviousWeight(sakeName, logSheet)
        If previousWeight = -1 Then
            MsgBox "このお酒の過去の記録が見つかりません。「新規開封」を選択してください。", vbExclamation
            CalculateAlcoholInfo = False
            Exit Function
        End If
        drankWeight = previousWeight - currentWeight
    Else
        MsgBox "「新規開封」または「続きから」を選択してください。", vbExclamation
        CalculateAlcoholInfo = False
        Exit Function
    End If

    ' 純アルコール量を計算 (比重: 0.8)
    pureAlcohol = drankWeight * (abv / 100) * 0.8

    CalculateAlcoholInfo = True
    Exit Function

ErrorHandler:
    MsgBox "計算中にエラーが発生しました。" & vbCrLf & Err.Description, vbCritical
    CalculateAlcoholInfo = False
End Function

' 指定されたお酒の前回の重量を取得する
Private Function GetPreviousWeight(ByVal sakeName As String, ByVal logSheet As Worksheet) As Double
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
