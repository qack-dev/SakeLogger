# SakeLogger �v���W�F�N�g ���r���[

**���r���[���{��**: 2025�N6��26��
**�Ώۃv���W�F�N�g**: SakeLogger

## �v���W�F�N�g�T�v

SakeLogger�́AExcel VBA�ŊJ�����ꂽ���{���̈���ʊǗ��Ə��A���R�[���ʉ����c�[���ł��B�����̎�ʂ��L�^���A���N�Ǘ��ɖ𗧂Ă邱�Ƃ�ړI�Ƃ��Ă��܂��B

## �ǂ��Ƃ���

### 1. ���p�I�ȃA�v���P�[�V�����݌v
- **���m�ȖړI**: ���{���̈���ʊǗ��Ƃ�����̓I�Ŏ��p�I�ȖړI������
- **���N�u��**: ���A���R�[���ʂ̌v�Z�ɂ��A���N�Ǘ��ɔz�������݌v
- **���[�U�r���e�B**: �t�H�[���x�[�X��GUI�Œ����I�ȑ��삪�\

### 2. �K�؂ȃf�[�^�\��
- **�}�X�^�f�[�^�Ǘ�**: �����}�X�^�V�[�g�Ŏ��̊�{�����ꌳ�Ǘ�
- **���O�L�^**: �����L�^�V�[�g�Ŏ��n��f�[�^��~��
- **�W�v�@�\**: ���ʂ̏��A���R�[���ʏW�v�ŌX�����͂��\

### 3. �����ȃv���O���~���O��@
- **�萔��`**: ��ԍ���萔�ŊǗ����A�ێ琫������
- **�G���[�n���h�����O**: `On Error GoTo`���g�p�����K�؂ȃG���[����
- **���͌���**: ���K�\���ɂ����t�`���`�F�b�N�A���l����
- **�I�u�W�F�N�g�Ǘ�**: `setObj()`��`releaseObj()`�Ń������Ǘ�

### 4. �J�����̐���
- **XVBA�g�p**: ���_����VBA�J�����̗̍p
- **Git�Ǘ�**: �o�[�W�����Ǘ��V�X�e���̓K�؂Ȏg�p
- **���W���[������**: VBA�R�[�h���t�@�C���P�ʂŕ����Ǘ�

## �����Ƃ���

### 1. �R�[�h�̉ǐ��E�ێ琫�̖��
- **�n�[�h�R�[�f�B���O**: �}�W�b�N�i���o�[��Œ蕶����̑��p
- **�֐��̔�剻**: `CalcAlcoholInfo`�֐��������̐ӔC����������
- **�O���[�o���ϐ�**: �����̃O���[�o���ϐ��ɂ�錋���x�̍���

### 2. �G���[�n���h�����O�̕s��
- **�s���S�ȗ�O����**: �ꕔ�̊֐��ŃG���[�n���h�����O���s�\��
- **���[�U�[���b�Z�[�W**: �G���[���b�Z�[�W���Z�p�I������ꍇ������
- **���[���o�b�N�@�\**: �f�[�^�X�V���s���̕����@�\���s��

### 3. �f�[�^���؂̊Â�
- **�d�ʃ`�F�b�N**: �����I�ɕs�\�Ȓl�̌��؂��s�\��
- **���t����**: ��������ߋ��ُ̈�ȓ��t�̃`�F�b�N���s��
- **�d���`�F�b�N**: ��������̏d���L�^�h�~�@�\���Ȃ�

### 4. ���[�U�r���e�B�̉ۑ�
- **����t���[**: �V�i�J���ƌp�����p�̑I����������ɂ���
- **�f�[�^�\��**: �ߋ��̋L�^���m�F����@�\���s��
- **�o�b�N�A�b�v**: �f�[�^�̃o�b�N�A�b�v�E�����@�\���Ȃ�

## ���P�_

### 1. �R�[�h�\���̉��P
- **�֐��̕���**: �傫�Ȋ֐���ӔC���Ƃɕ���
- **�萔�̏W��**: �ݒ�l��ݒ�t�@�C���܂��͒萔���W���[���ɏW��
- **�G���[�n���h�����O����**: ���ʂ̃G���[�n���h�����O�֐����쐬

### 2. �f�[�^���؂̋���
- **���͒l����**: ��茵���ȓ��͒l�`�F�b�N�@�\�̎���
- **�Ɩ����[������**: ����ʂ̕����I����`�F�b�N
- **�d���h�~**: ��������ł̏d���L�^�h�~

### 3. ���[�U�r���e�B�̌���
- **����K�C�h**: ���񗘗p�Ҍ����̃w���v�@�\
- **�f�[�^�{��**: �ߋ��L�^�̌����E�\���@�\
- **�O���t�\��**: �����X���̉����@�\

### 4. �ێ琫�̌���
- **�ݒ�O����**: �n�[�h�R�[�f�B���O���ꂽ�l�̊O����
- **���O�@�\**: ���엚���ƃG���[���O�̋L�^
- **�e�X�g�@�\**: �P�̃e�X�g�@�\�̒ǉ�

## ���P����Ȃ�u�����ׂ��R�[�h����

### 1. �萔�̊O����

**�C���O:**
```vba
'�O���[�o���萔
'�����}�X�^�V�[�g
Public Const idCol As Integer = 1 'ID��
Public Const nameCol As Integer = 2 '�����̖��O��
Public Const kindsCol As Integer = 3 '��ޗ�
Public Const alcoholCol As Integer = 4 '�x����
Public Const fullCol As Integer = 5 '���J���d�ʗ�
Public Const empCol As Integer = 6 '��d�ʗ�
```

**�C����:**
```vba
'�ݒ胂�W���[�� (ConfigModule.bas)
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

### 2. �G���[�n���h�����O�̓���

**�C���O:**
```vba
Public Function CalcAlcoholInfo(sakeName As String, nowWeight As Double, _
                         ByRef drankWeight As Double, ByRef pureAlcohol As Double) As Boolean
    On Error GoTo ErrHandler
    ' ... ���� ...
ErrHandler:
    MsgBox "�v�Z���ɃG���[���������܂����B" & vbCrLf & Err.Description, vbCritical
    CalcAlcoholInfo = False
End Function
```

**�C����:**
```vba
'���ʃG���[�n���h�����O���W���[�� (ErrorHandler.bas)
Public Sub HandleError(ByVal functionName As String, ByVal errNumber As Long, ByVal errDescription As String)
    Dim errorMsg As String
    errorMsg = "�G���[���������܂����B" & vbCrLf & _
               "�֐�: " & functionName & vbCrLf & _
               "�G���[�ԍ�: " & errNumber & vbCrLf & _
               "�ڍ�: " & errDescription
    
    ' ���O�t�@�C���ɋL�^
    Call WriteErrorLog(functionName, errNumber, errDescription)
    
    MsgBox errorMsg, vbCritical, "SakeLogger �G���["
End Sub

Public Function CalcAlcoholInfo(sakeName As String, nowWeight As Double, _
                         ByRef drankWeight As Double, ByRef pureAlcohol As Double) As Boolean
    On Error GoTo ErrHandler
    ' ... ���� ...
    CalcAlcoholInfo = True
    Exit Function
    
ErrHandler:
    Call HandleError("CalcAlcoholInfo", Err.Number, Err.Description)
    CalcAlcoholInfo = False
End Function
```

### 3. ���͌��؂̋���

**�C���O:**
```vba
If nowWeight > fullWeight Or nowWeight < emptyWeight Then
    MsgBox "���݂̏d�����s���ł��B", vbExclamation
    Exit Function
End If
```

**�C����:**
```vba
'���͌��؃��W���[�� (ValidationModule.bas)
Public Function ValidateWeight(ByVal nowWeight As Double, ByVal fullWeight As Double, ByVal emptyWeight As Double) As ValidationResult
    Dim result As ValidationResult
    
    If nowWeight < 0 Then
        result.IsValid = False
        result.ErrorMessage = "�d�ʂ�0�ȏ�œ��͂��Ă��������B"
    ElseIf nowWeight > fullWeight Then
        result.IsValid = False
        result.ErrorMessage = "���݂̏d�ʂ����J�����̏d�ʂ𒴂��Ă��܂��B" & vbCrLf & _
                             "���J���d��: " & fullWeight & "g" & vbCrLf & _
                             "���͒l: " & nowWeight & "g"
    ElseIf nowWeight < emptyWeight Then
        result.IsValid = False
        result.ErrorMessage = "���݂̏d�ʂ���{�g���d�ʂ�������Ă��܂��B" & vbCrLf & _
                             "��{�g���d��: " & emptyWeight & "g" & vbCrLf & _
                             "���͒l: " & nowWeight & "g"
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

### 4. �֐��̐ӔC����

**�C���O:**
```vba
Public Function CalcAlcoholInfo(sakeName As String, nowWeight As Double, _
                         ByRef drankWeight As Double, ByRef pureAlcohol As Double) As Boolean
    ' �}�X�^�����A�d�ʌv�Z�A���A���R�[���v�Z����̊֐��Ŏ��s
End Function
```

**�C����:**
```vba
'���}�X�^�����֐�
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

'�d�ʌv�Z�֐�
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

'���A���R�[���v�Z�֐�
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

### 5. �ݒ�t�@�C���̊��p

**�C���O:**
```vba
' �n�[�h�R�[�f�B���O���ꂽ���b�Z�[�W
MsgBox "�L�^��ۑ����܂����I", vbInformation
```

**�C����:**
```vba
'�ݒ�t�@�C�� (config.json) �ɒǉ�
{
  "messages": {
    "save_success": "�L�^��ۑ����܂����I",
    "save_error": "�L�^�̕ۑ��Ɏ��s���܂����B",
    "validation_error": "���͓��e�Ɍ�肪����܂��B",
    "calculation_error": "�v�Z�����ŃG���[���������܂����B"
  },
  "validation": {
    "max_weight": 10000,
    "min_weight": 0,
    "max_abv": 100,
    "min_abv": 0
  }
}

'VBA�R�[�h
Public Function GetConfigMessage(ByVal key As String) As String
    ' config.json���烁�b�Z�[�W��ǂݍ��ޏ���
    ' �����͏ȗ�
End Function

' �g�p��
MsgBox GetConfigMessage("save_success"), vbInformation
```

## �����]��

SakeLogger�͎��p�I�ŉ��l�̂���A�v���P�[�V�����ł��B��{�I�ȋ@�\�͓K�؂Ɏ�������Ă���AVBA�̓��������������݌v�ƂȂ��Ă��܂��B�������A�R�[�h�̕ێ琫�A�G���[�n���h�����O�A���[�U�r���e�B�̖ʂŉ��P�̗]�n������܂��B

��L�̉��P�_��i�K�I�Ɏ������邱�ƂŁA��茘�S�Ŏg���₷���A�v���P�[�V�����ɔ��W�����邱�Ƃ��ł���ł��傤�B���ɁA�G���[�n���h�����O�̓���Ɠ��͌��؂̋����́A���[�U�[�̌��̌���ɒ�������d�v�ȉ��P�_�ł��B

---
*���̃��r���[��2025�N6��26���Ɏ��{����܂����B*
