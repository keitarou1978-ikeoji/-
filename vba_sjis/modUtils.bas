Attribute VB_Name = "modUtils"
Option Explicit

'==================================================================
' modUtils
' 共通ユーティリティ（日付処理・書式設定）とテストデータ生成。
'==================================================================

' 指定シートの指定列で、値が入っている最終行を返す
Public Function GetLastRow(ByVal ws As Worksheet, ByVal colIdx As Long) As Long
    On Error GoTo ErrHandler
    GetLastRow = ws.Cells(ws.Rows.Count, colIdx).End(xlUp).Row
    Exit Function
ErrHandler:
    GetLastRow = 1
End Function

' 日付セルへ標準書式（yyyy/mm/dd）を適用する
Public Sub FormatAsDate(ByVal targetRange As Range)
    On Error GoTo ErrHandler
    targetRange.NumberFormatLocal = "yyyy/mm/dd"
    Exit Sub
ErrHandler:
    ' 失敗時は何もしない
End Sub

' 金額セルへ標準書式（\#,##0）を適用する
Public Sub FormatAsCurrency(ByVal targetRange As Range)
    On Error GoTo ErrHandler
    targetRange.NumberFormatLocal = "\#,##0"
    Exit Sub
ErrHandler:
    ' 失敗時は何もしない
End Sub

' 月初日を返す
Public Function FirstDayOfMonth(ByVal baseDate As Date) As Date
    FirstDayOfMonth = DateSerial(Year(baseDate), Month(baseDate), 1)
End Function

' 月末日を返す
Public Function LastDayOfMonth(ByVal baseDate As Date) As Date
    LastDayOfMonth = DateSerial(Year(baseDate), Month(baseDate) + 1, 0)
End Function

'------------------------------------------------------------------
' 動作確認用: 手入力シートにサンプルデータを10件生成する
'  - 今月の日付でランダムに10件
'  - 大分類/中分類はカテゴリマスタからバランスよく選択
'  - 金額は 100～50,000 円の現実的な値
'  - 支払方法は3種をバランスよく配分
'------------------------------------------------------------------
Public Sub Test_GenerateSampleData()
    On Error GoTo ErrHandler
    Dim prevScreen As Boolean
    Dim prevCalc As XlCalculation
    prevScreen = Application.ScreenUpdating
    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Const SAMPLE_COUNT As Long = 10
    Dim majors As Variant
    Dim payments As Variant
    Dim i As Long
    Dim writeRow As Long
    Dim majorName As String
    Dim middles As Variant
    Dim middleName As String
    Dim payName As String
    Dim firstDay As Date
    Dim lastDay As Date
    Dim dayCount As Long
    Dim amountVal As Long

    majors = modConfig.GetMajorList()
    payments = modConfig.GetPaymentList()

    If Not IsArray(majors) Then GoTo NoMaster
    If UBound(majors) < LBound(majors) Then GoTo NoMaster
    If Not IsArray(payments) Then GoTo NoMaster
    If UBound(payments) < LBound(payments) Then GoTo NoMaster

    ' 既存の明細をクリアしてから書き込む（見出し行より下）
    Dim lastRow As Long
    lastRow = modUtils.GetLastRow(shtManual, COL_MAJOR)
    If lastRow >= MANUAL_FIRST_ROW Then
        shtManual.Range(shtManual.Cells(MANUAL_FIRST_ROW, COL_DATE), _
                        shtManual.Cells(lastRow, COL_POSTED)).ClearContents
    End If

    firstDay = modUtils.FirstDayOfMonth(Date)
    lastDay = modUtils.LastDayOfMonth(Date)
    dayCount = CLng(lastDay - firstDay) + 1

    Randomize

    For i = 0 To SAMPLE_COUNT - 1
        writeRow = MANUAL_FIRST_ROW + i

        ' 大分類をバランスよく巡回選択
        majorName = CStr(majors((i) Mod (UBound(majors) - LBound(majors) + 1) + LBound(majors)))

        ' 中分類は当該大分類からランダム選択
        middles = modConfig.GetMiddleCategories(majorName)
        If IsArray(middles) Then
            If UBound(middles) >= LBound(middles) Then
                middleName = CStr(middles(Int(Rnd() * (UBound(middles) - LBound(middles) + 1)) + LBound(middles)))
            Else
                middleName = ""
            End If
        Else
            middleName = ""
        End If

        ' 支払方法をバランスよく巡回配分
        payName = CStr(payments((i) Mod (UBound(payments) - LBound(payments) + 1) + LBound(payments)))

        ' 金額 100～50,000 円（100円単位に丸め）
        amountVal = (Int(Rnd() * 500) + 1) * 100

        ' 書き込み
        shtManual.Cells(writeRow, COL_DATE).Value = firstDay + Int(Rnd() * dayCount)
        shtManual.Cells(writeRow, COL_DATE).NumberFormatLocal = "yyyy/mm/dd"
        shtManual.Cells(writeRow, COL_MAJOR).Value = majorName
        shtManual.Cells(writeRow, COL_MIDDLE).Value = middleName
        shtManual.Cells(writeRow, COL_DESC).Value = "サンプル" & Format$(i + 1, "00")
        shtManual.Cells(writeRow, COL_AMOUNT).Value = amountVal
        shtManual.Cells(writeRow, COL_AMOUNT).NumberFormatLocal = "\#,##0"
        shtManual.Cells(writeRow, COL_PAYMENT).Value = payName
        shtManual.Cells(writeRow, COL_INEXP).Value = modConfig.GetInExpType(majorName)
        shtManual.Cells(writeRow, COL_POSTED).Value = ""
    Next i

    Application.ScreenUpdating = prevScreen
    Application.Calculation = prevCalc
    MsgBox SAMPLE_COUNT & " 件のサンプルデータを生成しました。", vbInformation, "テストデータ生成"
    Exit Sub

NoMaster:
    Application.ScreenUpdating = prevScreen
    Application.Calculation = prevCalc
    MsgBox "カテゴリマスタ／支払方法マスタが読み込めませんでした。設定シートを確認してください。", _
           vbExclamation, "テストデータ生成"
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = prevScreen
    Application.Calculation = prevCalc
    MsgBox "サンプルデータ生成でエラーが発生しました: " & Err.Description, vbExclamation, "テストデータ生成"
End Sub
