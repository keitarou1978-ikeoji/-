Attribute VB_Name = "modManualEntry"
Option Explicit

'==================================================================
' modManualEntry
' 手入力シートの入力補助・台帳転記を担うモジュール。
' 連動ドロップダウン等の実処理はここに集約し、
' シートモジュール(shtManual)の Worksheet_Change から呼び出す。
'==================================================================

' 手入力シートの初期セットアップ（B/F列の入力規則、既定日付の下準備）。
' 名前付き範囲は既に定義済みである前提。
Public Sub SetupManualValidation()
    On Error GoTo ErrHandler
    Dim prevScreen As Boolean
    Dim prevCalc As XlCalculation
    prevScreen = Application.ScreenUpdating
    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim lastRow As Long
    lastRow = 1000  ' 入力規則を敷く行数

    ' B列: 大分類ドロップダウン（名前付き範囲 大分類リスト）
    With shtManual.Range(shtManual.Cells(MANUAL_FIRST_ROW, COL_MAJOR), _
                         shtManual.Cells(lastRow, COL_MAJOR)).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="=" & NR_MAJOR_LIST
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    ' F列: 支払方法ドロップダウン（名前付き範囲 支払方法リスト）
    With shtManual.Range(shtManual.Cells(MANUAL_FIRST_ROW, COL_PAYMENT), _
                         shtManual.Cells(lastRow, COL_PAYMENT)).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="=" & NR_PAYMENT_LIST
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

CleanExit:
    Application.ScreenUpdating = prevScreen
    Application.Calculation = prevCalc
    Exit Sub
ErrHandler:
    MsgBox "入力規則の設定でエラーが発生しました: " & Err.Description, vbExclamation, "手入力シート"
    Resume CleanExit
End Sub

' B列（大分類）変更時に、C列（中分類）の連動ドロップダウンを再構築する。
' targetCell は変更された B 列セル。
Public Sub RefreshMiddleDropdown(ByVal targetCell As Range)
    On Error GoTo ErrHandler
    Dim majorName As String
    Dim middles As Variant
    Dim listStr As String
    Dim middleCell As Range

    majorName = Trim$(CStr(targetCell.Value))
    Set middleCell = shtManual.Cells(targetCell.Row, COL_MIDDLE)

    ' いったん中分類の値と入力規則をクリア
    middleCell.Validation.Delete
    middleCell.ClearContents

    If majorName = "" Then Exit Sub

    middles = modConfig.GetMiddleCategories(majorName)
    listStr = JoinArray(middles, ",")
    If listStr = "" Then Exit Sub

    ' カンマ区切りリストで入力規則を設定（中分類は件数が少なくインライン指定で十分）
    With middleCell.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:=listStr
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    Exit Sub
ErrHandler:
    ' 連動生成の失敗はユーザー操作を止めない
End Sub

' 大分類から収支区分を判定し、G列へ書き込む
Public Sub ApplyInExpType(ByVal rowIdx As Long)
    On Error GoTo ErrHandler
    Dim majorName As String
    majorName = Trim$(CStr(shtManual.Cells(rowIdx, COL_MAJOR).Value))
    If majorName = "" Then
        shtManual.Cells(rowIdx, COL_INEXP).ClearContents
    Else
        shtManual.Cells(rowIdx, COL_INEXP).Value = modConfig.GetInExpType(majorName)
    End If
    Exit Sub
ErrHandler:
    ' 判定失敗時は何もしない
End Sub

' 新規行に既定の日付（今日）をセットする（A列が空のときのみ）
Public Sub SetDefaultDate(ByVal rowIdx As Long)
    On Error GoTo ErrHandler
    Dim dateCell As Range
    Set dateCell = shtManual.Cells(rowIdx, COL_DATE)
    If Len(Trim$(CStr(dateCell.Value))) = 0 Then
        dateCell.Value = Date
        dateCell.NumberFormatLocal = "yyyy/mm/dd"
    End If
    Exit Sub
ErrHandler:
    ' 失敗時は何もしない
End Sub

' 入力行の妥当性チェック（金額は正の整数、日付はDate型、必須項目）。
' 問題があれば False とメッセージを返す。
Public Function ValidateRow(ByVal rowIdx As Long, ByRef message As String) As Boolean
    On Error GoTo ErrHandler
    Dim dateVal As Variant
    Dim amountVal As Variant
    Dim majorVal As String

    ValidateRow = True
    message = ""

    dateVal = shtManual.Cells(rowIdx, COL_DATE).Value
    majorVal = Trim$(CStr(shtManual.Cells(rowIdx, COL_MAJOR).Value))
    amountVal = shtManual.Cells(rowIdx, COL_AMOUNT).Value

    ' 完全な空行はチェック対象外（True 扱い）
    If Len(Trim$(CStr(dateVal))) = 0 And majorVal = "" _
       And Len(Trim$(CStr(amountVal))) = 0 Then
        Exit Function
    End If

    If Not IsDate(dateVal) Then
        ValidateRow = False
        message = rowIdx & "行目: 日付が正しくありません。"
        Exit Function
    End If

    If majorVal = "" Then
        ValidateRow = False
        message = rowIdx & "行目: 大分類は必須です。"
        Exit Function
    End If

    If Not IsNumeric(amountVal) Then
        ValidateRow = False
        message = rowIdx & "行目: 金額は数値で入力してください。"
        Exit Function
    End If

    If CDbl(amountVal) <= 0 Or CDbl(amountVal) <> Int(CDbl(amountVal)) Then
        ValidateRow = False
        message = rowIdx & "行目: 金額は正の整数で入力してください。"
        Exit Function
    End If

    Exit Function
ErrHandler:
    ValidateRow = False
    message = rowIdx & "行目: 検証中にエラーが発生しました。"
End Function

'------------------------------------------------------------------
' 取引台帳へ転記（Phase 1 は枠のみ）
' Phase 2 で取引台帳シートを作成した後に本接続する。
' それまでは転記先シート存在チェックをスキップし、案内のみ表示する。
'------------------------------------------------------------------
Public Sub TransferToLedger()
    On Error GoTo ErrHandler
    Dim prevScreen As Boolean
    Dim prevCalc As XlCalculation
    prevScreen = Application.ScreenUpdating
    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' === Phase 1: 枠のみ ===
    ' Phase 2 で以下を実装する:
    '   1) 取引台帳シート(shtLedger)の存在確認
    '   2) 手入力の未転記行を検証し取引台帳へ append
    '   3) 転記済み行の H 列に「済」を記録
    MsgBox "「取引台帳へ転記」機能は Phase 2 で有効になります。" & vbCrLf & _
           "（現在は枠のみ実装されています）", vbInformation, "取引台帳へ転記"

CleanExit:
    Application.ScreenUpdating = prevScreen
    Application.Calculation = prevCalc
    Exit Sub
ErrHandler:
    MsgBox "転記処理でエラーが発生しました: " & Err.Description, vbExclamation, "取引台帳へ転記"
    Resume CleanExit
End Sub

' 配列を区切り文字で連結するヘルパー（空要素は無視）
Private Function JoinArray(ByVal arr As Variant, ByVal delimiter As String) As String
    On Error GoTo ErrHandler
    Dim i As Long
    Dim buf As String
    Dim item As String
    If Not IsArray(arr) Then
        JoinArray = ""
        Exit Function
    End If
    For i = LBound(arr) To UBound(arr)
        item = Trim$(CStr(arr(i)))
        If item <> "" Then
            If buf = "" Then
                buf = item
            Else
                buf = buf & delimiter & item
            End If
        End If
    Next i
    JoinArray = buf
    Exit Function
ErrHandler:
    JoinArray = ""
End Function
