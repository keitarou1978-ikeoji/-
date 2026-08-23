Attribute VB_Name = "modSetup"
Option Explicit

'==================================================================
' modSetup （Phase 2）
' 既存ブックに Phase 2 の要素を「追加」するセットアップ。
'   1) 取引台帳シート・CSV取込シートを生成
'   2) 各シートのボタンを配置しマクロを割り当て
'   3) 設定シートの CSV列マッピング／自動分類ルールを初期化
'   4) 名前付き範囲を追加/更新
' 何度実行しても安全（冪等）になるよう設計。
'==================================================================

' 一度だけ実行すればよいセットアップ本体
Public Sub Setup_Phase2()
    On Error GoTo ErrHandler
    Dim prevScreen As Boolean
    Dim prevCalc As XlCalculation
    prevScreen = Application.ScreenUpdating
    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    BuildLedgerSheet
    BuildImportSheet
    ConfigureSettings
    WireManualTransferButton

    Application.ScreenUpdating = prevScreen
    Application.Calculation = prevCalc
    MsgBox "Phase 2 のセットアップが完了しました。" & vbCrLf & _
           "・取引台帳／CSV取込シートを作成" & vbCrLf & _
           "・各ボタンとマクロを接続" & vbCrLf & _
           "・CSV列マッピング／自動分類ルールを初期化", vbInformation, "Setup_Phase2"
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = prevScreen
    Application.Calculation = prevCalc
    MsgBox "セットアップでエラーが発生しました: " & Err.Description, vbExclamation, "Setup_Phase2"
End Sub

'------------------------------------------------------------------
' 取引台帳シートの作成
'------------------------------------------------------------------
Private Sub BuildLedgerSheet()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(SHEET_LEDGER, RGB(31, 59, 95))  ' 濃紺
    ws.Cells.Clear

    ws.Range("A1").Value = "取引台帳 ― 全取引の正本（手入力・CSVの集約先）"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    Dim headers As Variant
    headers = Array("取引ID", "日付", "大分類", "中分類", "摘要", "金額", _
                    "支払方法", "収支区分", "入力元", "取込日時", "重複キー")
    ApplyHeaderRow ws, LEDGER_HEADER_ROW, headers

    ' 列幅
    Dim widths As Variant
    widths = Array(9, 12, 14, 14, 34, 12, 20, 10, 12, 18, 28)
    Dim c As Long
    For c = 0 To UBound(widths)
        ws.Columns(c + 1).ColumnWidth = widths(c)
    Next c
    ws.Range("K:K").Font.Color = RGB(150, 150, 150)  ' 重複キーは補助情報
    ws.Rows(LEDGER_HEADER_ROW).AutoFilter

    ' 見出し行の下で固定（アクティブ化して安全に設定）
    ws.Activate
    ws.Range("A" & LEDGER_FIRST_ROW).Select
    ActiveWindow.FreezePanes = True

    ' 動的名前付き範囲（Phase 3 の集計で使用）
    AddOrReplaceName "取引台帳", _
        "=OFFSET('" & SHEET_LEDGER & "'!$A$1,0,0,COUNTA('" & SHEET_LEDGER & "'!$A:$A)," & LEDGER_LAST_COL & ")"
End Sub

'------------------------------------------------------------------
' CSV取込シートの作成
'------------------------------------------------------------------
Private Sub BuildImportSheet()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(SHEET_IMPORT, RGB(0, 128, 0))  ' 緑
    ws.Cells.Clear
    On Error Resume Next
    ws.Buttons.Delete
    On Error GoTo 0

    ws.Range("A1").Value = "CSV取込 ― 口座／カードCSVを読み込み、確認して台帳へ取込む"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    ws.Range("A3").Value = "手順: ①「口座CSVを取込」or「カードCSVを取込」→ ②内容(黄色=未分類)を確認/修正 → ③「取引台帳へ取込」"
    ws.Range("A3").Font.Color = RGB(120, 120, 120)
    ws.Range("A3").Font.Size = 9

    Dim headers As Variant
    headers = Array("取込○", "日付", "摘要", "金額", "収支区分", "大分類", _
                    "中分類", "支払方法", "種別", "元行", "重複?")
    ApplyHeaderRow ws, IMPORT_HEADER_ROW, headers

    Dim widths As Variant
    widths = Array(7, 12, 34, 12, 10, 14, 14, 20, 8, 8, 8)
    Dim c As Long
    For c = 0 To UBound(widths)
        ws.Columns(c + 1).ColumnWidth = widths(c)
    Next c

    ' 大分類・支払方法の修正用ドロップダウン（未分類の手直し用）
    AddListValidation ws, IM_MAJOR, "=" & NR_MAJOR_LIST
    AddListValidation ws, IM_PAY, "=" & NR_PAYMENT_LIST

    ' ボタン配置（見出し行の上、2行目）
    AddButton ws, ws.Range("B2"), 120, 24, "口座CSVを取込", "modImport.ImportAccountCSV"
    AddButton ws, ws.Range("D2"), 120, 24, "カードCSVを取込", "modImport.ImportCardCSV"
    AddButton ws, ws.Range("F2"), 120, 24, "取引台帳へ取込", "modImport.CommitPreviewToLedger"
    AddButton ws, ws.Range("H2"), 110, 24, "プレビュー消去", "modImport.ClearPreview"
End Sub

'------------------------------------------------------------------
' 設定シート：CSV列マッピングと自動分類ルールを初期化
'------------------------------------------------------------------
Private Sub ConfigureSettings()
    ' --- CSV列マッピング（I:L を作り直す） ---
    shtConfig.Range("I2:M100").Clear
    Dim mapHdr As Variant
    mapHdr = Array("種別", "項目", "値", "備考")
    Dim c As Long
    For c = 0 To UBound(mapHdr)
        With shtConfig.Cells(2, 9 + c)
            .Value = mapHdr(c)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(68, 114, 196)
            .HorizontalAlignment = xlCenter
        End With
    Next c

    Dim rows As Variant
    rows = Array( _
        Array("口座", "見出し行数", 1, "先頭の見出し行数"), _
        Array("口座", "日付列", 1, "1始まりの列番号"), _
        Array("口座", "摘要列", 2, ""), _
        Array("口座", "出金列", 4, "支出になる列"), _
        Array("口座", "入金列", 5, "収入になる列"), _
        Array("口座", "文字コード", "Shift_JIS", "UTF-8も可"), _
        Array("口座", "既定支払方法", "メガバンク口座引落", ""), _
        Array("カード", "見出し行数", 1, ""), _
        Array("カード", "日付列", 1, ""), _
        Array("カード", "摘要列", 2, ""), _
        Array("カード", "金額列", 5, "支出金額の列"), _
        Array("カード", "文字コード", "Shift_JIS", "UTF-8も可"), _
        Array("カード", "既定支払方法", "クレジットカード", ""))

    Dim r As Long
    For r = 0 To UBound(rows)
        shtConfig.Cells(3 + r, 9).Value = rows(r)(0)
        shtConfig.Cells(3 + r, 10).Value = rows(r)(1)
        shtConfig.Cells(3 + r, 11).Value = rows(r)(2)
        shtConfig.Cells(3 + r, 12).Value = rows(r)(3)
    Next r
    AddOrReplaceName NR_CSV_MAPPING, "='設定'!$I$3:$L$" & (2 + UBound(rows) + 1)

    ' --- 自動分類ルール（O:Q を拡充） ---
    shtConfig.Range("O3:Q200").ClearContents
    Dim rules As Variant
    rules = Array( _
        Array("セブンイレブン", "食費", "食料品"), _
        Array("ローソン", "食費", "食料品"), _
        Array("ファミリーマート", "食費", "食料品"), _
        Array("イオン", "食費", "食料品"), _
        Array("スターバックス", "食費", "カフェ・嗜好品"), _
        Array("マクドナルド", "食費", "外食"), _
        Array("東京電力", "水道光熱費", "電気"), _
        Array("東京ガス", "水道光熱費", "ガス"), _
        Array("水道", "水道光熱費", "水道"), _
        Array("ドコモ", "通信費", "携帯電話"), _
        Array("ソフトバンク", "通信費", "携帯電話"), _
        Array("JR", "交通費", "電車・バス"), _
        Array("Suica", "交通費", "電車・バス"), _
        Array("ENEOS", "交通費", "ガソリン"), _
        Array("Amazon", "日用品", "消耗品"), _
        Array("ドラッグ", "日用品", "衛生用品"), _
        Array("薬局", "医療・健康", "病院・薬"), _
        Array("Netflix", "娯楽・交際費", "趣味・レジャー"), _
        Array("ユニクロ", "被服・美容", "衣服"), _
        Array("給与", "給与収入", "給与"))
    For r = 0 To UBound(rules)
        shtConfig.Cells(3 + r, 15).Value = rules(r)(0)
        shtConfig.Cells(3 + r, 16).Value = rules(r)(1)
        shtConfig.Cells(3 + r, 17).Value = rules(r)(2)
    Next r
    AddOrReplaceName NR_CLASSIFY_RULE, "='設定'!$O$3:$Q$" & (2 + UBound(rules) + 1)
End Sub

'------------------------------------------------------------------
' 手入力シートの「取引台帳へ転記」ボタンを本接続する
'------------------------------------------------------------------
Private Sub WireManualTransferButton()
    On Error Resume Next
    shtManual.Buttons.Delete
    On Error GoTo 0
    AddButton shtManual, shtManual.Range("A2"), 140, 22, "取引台帳へ転記", "modLedger.TransferManualToLedger"
    ' Phase 1 の説明セルは残置（害はない）
End Sub

'==================================================================
' 共通ヘルパー
'==================================================================

' 指定名のシートを取得（無ければ末尾に作成）。タブ色を設定。
Private Function GetOrCreateSheet(ByVal sheetName As String, ByVal tabColor As Long) As Worksheet
    Dim ws As Worksheet
    Set ws = modLedger.GetSheetByName(sheetName)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If
    ws.Tab.Color = tabColor
    Set GetOrCreateSheet = ws
End Function

' 見出し行を整形して書き込む
Private Sub ApplyHeaderRow(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal headers As Variant)
    Dim c As Long
    For c = 0 To UBound(headers)
        With ws.Cells(headerRow, c + 1)
            .Value = headers(c)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(68, 114, 196)
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(180, 180, 180)
        End With
    Next c
End Sub

' 指定列（明細行）にリスト入力規則を設定
Private Sub AddListValidation(ByVal ws As Worksheet, ByVal colIdx As Long, ByVal formula1 As String)
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(IMPORT_FIRST_ROW, colIdx), ws.Cells(1000, colIdx))
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=formula1
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub

' フォームボタンを配置しマクロを割り当てる
Private Sub AddButton(ByVal ws As Worksheet, ByVal anchor As Range, _
                      ByVal w As Double, ByVal h As Double, _
                      ByVal caption As String, ByVal macroName As String)
    Dim btn As Button
    Set btn = ws.Buttons.Add(anchor.Left, anchor.Top, w, h)
    btn.caption = caption
    btn.OnAction = macroName
    btn.Font.Size = 10
End Sub

' 名前付き範囲を追加/置換（ブックレベル）
Private Sub AddOrReplaceName(ByVal nm As String, ByVal refersTo As String)
    On Error Resume Next
    ThisWorkbook.Names(nm).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=nm, RefersTo:=refersTo
End Sub
