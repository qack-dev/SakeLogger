# SakeLogger プロジェクト レビュー

**レビュー実施日**: 2025年6月26日  
**対象プロジェクト**: SakeLogger  
**プロジェクトパス**: `/mnt/share/prog/VBA/Excel_XVBA/SakeLogger`

## プロジェクト概要

SakeLoggerは、Excel VBAで開発された日本酒の飲酒量管理と純アルコール量可視化ツールです。毎日の酒量を記録し、健康管理に役立てることを目的としています。

## 良いところ

### 1. 実用的なアプリケーション設計
- **明確な目的**: 日本酒の飲酒量管理という具体的で実用的な目的がある
- **健康志向**: 純アルコール量の計算により、健康管理に配慮した設計
- **ユーザビリティ**: フォームベースのGUIで直感的な操作が可能

### 2. 適切なデータ構造
- **マスタデータ管理**: お酒マスタシートで酒の基本情報を一元管理
- **ログ記録**: 飲酒記録シートで時系列データを蓄積
- **集計機能**: 日別の純アルコール量集計で傾向分析が可能

### 3. 堅実なプログラミング手法
- **定数定義**: 列番号を定数で管理し、保守性を向上
- **エラーハンドリング**: `On Error GoTo`を使用した適切なエラー処理
- **入力検証**: 正規表現による日付形式チェック、数値検証
- **オブジェクト管理**: `setObj()`と`releaseObj()`でメモリ管理

### 4. 開発環境の整備
- **XVBA使用**: モダンなVBA開発環境の採用
- **Git管理**: バージョン管理システムの適切な使用
- **モジュール分離**: VBAコードをファイル単位で分離管理

## 悪いところ

### 1. コードの可読性・保守性の問題
- **ハードコーディング**: マジックナンバーや固定文字列の多用
- **関数の肥大化**: `CalcAlcoholInfo`関数が複数の責任を持ちすぎ
- **グローバル変数**: 多数のグローバル変数による結合度の高さ

### 2. エラーハンドリングの不備
- **不完全な例外処理**: 一部の関数でエラーハンドリングが不十分
- **ユーザーメッセージ**: エラーメッセージが技術的すぎる場合がある
- **ロールバック機能**: データ更新失敗時の復旧機能が不足

### 3. データ検証の甘さ
- **重量チェック**: 物理的に不可能な値の検証が不十分
- **日付検証**: 未来日や過去の異常な日付のチェックが不足
- **重複チェック**: 同一日時の重複記録防止機能がない

### 4. ユーザビリティの課題
- **操作フロー**: 新品開封と継続飲用の選択が分かりにくい
- **データ表示**: 過去の記録を確認する機能が不足
- **バックアップ**: データのバックアップ・復元機能がない

## 改善点

### 1. コード構造の改善
- **関数の分割**: 大きな関数を責任ごとに分割
- **定数の集約**: 設定値を設定ファイルまたは定数モジュールに集約
- **エラーハンドリング統一**: 共通のエラーハンドリング関数を作成

### 2. データ検証の強化
- **入力値検証**: より厳密な入力値チェック機能の実装
- **業務ルール検証**: 飲酒量の物理的制約チェック
- **重複防止**: 同一条件での重複記録防止

### 3. ユーザビリティの向上
- **操作ガイド**: 初回利用者向けのヘルプ機能
- **データ閲覧**: 過去記録の検索・表示機能
- **グラフ表示**: 飲酒傾向の可視化機能

### 4. 保守性の向上
- **設定外部化**: ハードコーディングされた値の外部化
- **ログ機能**: 操作履歴とエラーログの記録
- **テスト機能**: 単体テスト機能の追加

## 改善するなら置換すべきコード部分

### 1. 定数の外部化

**修正前:**
```vba
'グローバル定数
'お酒マスタシート
Public Const idCol As Integer = 1 'ID列
Public Const nameCol As Integer = 2 'お酒の名前列
Public Const kindsCol As Integer = 3 '種類列
Public Const alcoholCol As Integer = 4 '度数列
Public Const fullCol As Integer = 5 '未開封重量列
Public Const empCol As Integer = 6 '空重量列
```

**修正後:**
```vba
'設定モジュール (ConfigModule.bas)
Public Type MasterSheetConfig
    idCol As Integer
    nameCol As Integer
    kindsCol As Integer
    alcoholCol As Integer
    fullCol As Integer
    empCol As Integer
End Type

Public Type LogSheetConfig
    logDateCol As Integer
    logNameCol As Integer
    logNowCol As Integer
    logPureAlcCol As Integer
    logDrunkCol As Integer
    logComCol As Integer
    logIdCol As Integer
End Type

Public Function GetMasterConfig() As MasterSheetConfig
    Dim config As MasterSheetConfig
    config.idCol = 1
    config.nameCol = 2
    config.kindsCol = 3
    config.alcoholCol = 4
    config.fullCol = 5
    config.empCol = 6
    GetMasterConfig = config
End Function
```

### 2. エラーハンドリングの統一

**修正前:**
```vba
Public Function CalcAlcoholInfo(sakeName As String, nowWeight As Double, _
                         ByRef drankWeight As Double, ByRef pureAlcohol As Double) As Boolean
    On Error GoTo ErrHandler
    ' ... 処理 ...
ErrHandler:
    MsgBox "計算中にエラーが発生しました。" & vbCrLf & Err.Description, vbCritical
    CalcAlcoholInfo = False
End Function
```

**修正後:**
```vba
'共通エラーハンドリングモジュール (ErrorHandler.bas)
Public Sub HandleError(ByVal functionName As String, ByVal errNumber As Long, ByVal errDescription As String)
    Dim errorMsg As String
    errorMsg = "エラーが発生しました。" & vbCrLf & _
               "関数: " & functionName & vbCrLf & _
               "エラー番号: " & errNumber & vbCrLf & _
               "詳細: " & errDescription
    
    ' ログファイルに記録
    Call WriteErrorLog(functionName, errNumber, errDescription)
    
    MsgBox errorMsg, vbCritical, "SakeLogger エラー"
End Sub

Public Function CalcAlcoholInfo(sakeName As String, nowWeight As Double, _
                         ByRef drankWeight As Double, ByRef pureAlcohol As Double) As Boolean
    On Error GoTo ErrHandler
    ' ... 処理 ...
    CalcAlcoholInfo = True
    Exit Function
    
ErrHandler:
    Call HandleError("CalcAlcoholInfo", Err.Number, Err.Description)
    CalcAlcoholInfo = False
End Function
```

### 3. 入力検証の強化

**修正前:**
```vba
If nowWeight > fullWeight Or nowWeight < emptyWeight Then
    MsgBox "現在の重さが不正です。", vbExclamation
    Exit Function
End If
```

**修正後:**
```vba
'入力検証モジュール (ValidationModule.bas)
Public Function ValidateWeight(ByVal nowWeight As Double, ByVal fullWeight As Double, ByVal emptyWeight As Double) As ValidationResult
    Dim result As ValidationResult
    
    If nowWeight < 0 Then
        result.IsValid = False
        result.ErrorMessage = "重量は0以上で入力してください。"
    ElseIf nowWeight > fullWeight Then
        result.IsValid = False
        result.ErrorMessage = "現在の重量が未開封時の重量を超えています。" & vbCrLf & _
                             "未開封重量: " & fullWeight & "g" & vbCrLf & _
                             "入力値: " & nowWeight & "g"
    ElseIf nowWeight < emptyWeight Then
        result.IsValid = False
        result.ErrorMessage = "現在の重量が空ボトル重量を下回っています。" & vbCrLf & _
                             "空ボトル重量: " & emptyWeight & "g" & vbCrLf & _
                             "入力値: " & nowWeight & "g"
    Else
        result.IsValid = True
        result.ErrorMessage = ""
    End If
    
    ValidateWeight = result
End Function

Public Type ValidationResult
    IsValid As Boolean
    ErrorMessage As String
End Type
```

### 4. 関数の責任分離

**修正前:**
```vba
Public Function CalcAlcoholInfo(sakeName As String, nowWeight As Double, _
                         ByRef drankWeight As Double, ByRef pureAlcohol As Double) As Boolean
    ' マスタ検索、重量計算、純アルコール計算を一つの関数で実行
End Function
```

**修正後:**
```vba
'酒マスタ検索関数
Public Function FindSakeInfo(ByVal sakeName As String) As SakeInfo
    Dim info As SakeInfo
    Dim i As Long
    
    For i = 2 To lastCell.Row
        If wsMaster.Cells(i, idCol).Value & "." & wsMaster.Cells(i, nameCol).Value = sakeName Then
            info.Found = True
            info.ABV = wsMaster.Cells(i, alcoholCol).Value
            info.FullWeight = wsMaster.Cells(i, fullCol).Value
            info.EmptyWeight = wsMaster.Cells(i, empCol).Value
            Exit For
        End If
    Next i
    
    FindSakeInfo = info
End Function

'重量計算関数
Public Function CalcDrankWeight(ByVal sakeName As String, ByVal nowWeight As Double, ByVal isNewOpen As Boolean) As Double
    If isNewOpen Then
        Dim sakeInfo As SakeInfo
        sakeInfo = FindSakeInfo(sakeName)
        CalcDrankWeight = sakeInfo.FullWeight - nowWeight
    Else
        Dim prevWeight As Double
        prevWeight = GetPreviousWeight(sakeName)
        CalcDrankWeight = prevWeight - nowWeight
    End If
End Function

'純アルコール計算関数
Public Function CalcPureAlcohol(ByVal drankWeight As Double, ByVal abv As Double) As Double
    CalcPureAlcohol = drankWeight * (abv / 100) * 0.8
End Function

Public Type SakeInfo
    Found As Boolean
    ABV As Double
    FullWeight As Double
    EmptyWeight As Double
End Type
```

### 5. 設定ファイルの活用

**修正前:**
```vba
' ハードコーディングされたメッセージ
MsgBox "記録を保存しました！", vbInformation
```

**修正後:**
```vba
'設定ファイル (config.json) に追加
{
  "messages": {
    "save_success": "記録を保存しました！",
    "save_error": "記録の保存に失敗しました。",
    "validation_error": "入力内容に誤りがあります。",
    "calculation_error": "計算処理でエラーが発生しました。"
  },
  "validation": {
    "max_weight": 10000,
    "min_weight": 0,
    "max_abv": 100,
    "min_abv": 0
  }
}

'VBAコード
Public Function GetConfigMessage(ByVal key As String) As String
    ' config.jsonからメッセージを読み込む処理
    ' 実装は省略
End Function

' 使用例
MsgBox GetConfigMessage("save_success"), vbInformation
```

## 総合評価

SakeLoggerは実用的で価値のあるアプリケーションです。基本的な機能は適切に実装されており、VBAの特性を活かした設計となっています。しかし、コードの保守性、エラーハンドリング、ユーザビリティの面で改善の余地があります。

上記の改善点を段階的に実装することで、より堅牢で使いやすいアプリケーションに発展させることができるでしょう。特に、エラーハンドリングの統一と入力検証の強化は、ユーザー体験の向上に直結する重要な改善点です。

---
*このレビューは2025年6月26日に実施されました。*