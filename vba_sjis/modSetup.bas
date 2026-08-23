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

    ws.Range("A3").Value = "手順: ①プロファイルを選ぶ → ②「CSVを取込」→ ③内容(黄色=未分類/灰色=振替)を確認/修正 → ④「取引台帳へ取込」"
    ws.Range("A3").Font.Color = RGB(120, 120, 120)
    ws.Range("A3").Font.Size = 9

    ' プロファイル選択セル（B2）＋ラベル
    ws.Range("A2").Value = "取込プロファイル:"
    ws.Range("A2").Font.Bold = True
    With ws.Range(IMPORT_PROFILE_CELL).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=プロファイルリスト"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    ws.Range(IMPORT_PROFILE_CELL).Value = "口座A"
    ws.Range(IMPORT_PROFILE_CELL).Interior.Color = RGB(255, 255, 0)

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

    ' ボタン配置：B2ドロップダウンの右へ、座標を順送りで等間隔に並べる（重なり防止）
    ws.Rows(2).RowHeight = 28
    Dim topPos As Double, leftPos As Double
    topPos = ws.Range("C2").Top + 2
    leftPos = ws.Range("C2").Left
    leftPos = AddButtonSeq(ws, leftPos, topPos, 110, 24, "CSVを取込", "modImport.ImportSelectedProfile")
    leftPos = AddButtonSeq(ws, leftPos, topPos, 120, 24, "取引台帳へ取込", "modImport.CommitPreviewToLedger")
    leftPos = AddButtonSeq(ws, leftPos, topPos, 110, 24, "プレビュー消去", "modImport.ClearPreview")
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

    ' プロファイル定義（実CSVに合わせた初期値。列番号は1始まり）
    '  口座A: 日付1 / 摘要=摘要内容3 / 出金4 / 入金5 / 見出し1行
    '  口座B: 日付1 / 摘要=取扱内容4 / 出金=お引出し2 / 入金=お預入れ3 / 見出し1行
    '  カード: 日付=利用日4 / 摘要=利用店名3 / 金額7 / 見出し2行(名義行を除外)
    Dim rows As Variant
    rows = Array( _
        Array("口座A", "種類", "口座", "口座/カード"), _
        Array("口座A", "見出し行数", 1, "先頭の見出し行数"), _
        Array("口座A", "日付列", 1, "1始まりの列番号"), _
        Array("口座A", "摘要列", 3, "摘要内容(相手先)"), _
        Array("口座A", "出金列", 4, "支出になる列"), _
        Array("口座A", "入金列", 5, "収入になる列"), _
        Array("口座A", "文字コード", "Shift_JIS", "UTF-8も可"), _
        Array("口座A", "既定支払方法", "メガバンク口座引落", ""), _
        Array("口座B", "種類", "口座", ""), _
        Array("口座B", "見出し行数", 1, ""), _
        Array("口座B", "日付列", 1, ""), _
        Array("口座B", "摘要列", 4, "お取り扱い内容"), _
        Array("口座B", "出金列", 2, "お引出し"), _
        Array("口座B", "入金列", 3, "お預入れ"), _
        Array("口座B", "文字コード", "Shift_JIS", ""), _
        Array("口座B", "既定支払方法", "メガバンク口座引落", ""), _
        Array("カード", "種類", "カード", ""), _
        Array("カード", "見出し行数", 2, "名義行を除外"), _
        Array("カード", "日付列", 4, "ご利用日"), _
        Array("カード", "摘要列", 3, "ご利用店名"), _
        Array("カード", "金額列", 7, "ご利用金額(円)"), _
        Array("カード", "文字コード", "Shift_JIS", ""), _
        Array("カード", "既定支払方法", "クレジットカード", ""))

    Dim r As Long
    For r = 0 To UBound(rows)
        shtConfig.Cells(3 + r, 9).Value = rows(r)(0)
        shtConfig.Cells(3 + r, 10).Value = rows(r)(1)
        shtConfig.Cells(3 + r, 11).Value = rows(r)(2)
        shtConfig.Cells(3 + r, 12).Value = rows(r)(3)
    Next r
    AddOrReplaceName NR_CSV_MAPPING, "='設定'!$I$3:$L$" & (2 + UBound(rows) + 1)

    ' --- プロファイル一覧（N列。CSV取込シートのドロップダウン用） ---
    shtConfig.Range("N2:N50").ClearContents
    With shtConfig.Cells(2, 14)
        .Value = "プロファイル"
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(68, 114, 196)
        .HorizontalAlignment = xlCenter
    End With
    Dim profs As Variant
    profs = Array("口座A", "口座B", "カード")
    Dim p As Long
    For p = 0 To UBound(profs)
        shtConfig.Cells(3 + p, 14).Value = profs(p)
    Next p
    AddOrReplaceName "プロファイルリスト", "='設定'!$N$3:$N$" & (2 + UBound(profs) + 1)

    ' --- 自動分類ルール（O:Q を拡充） ---
    shtConfig.Range("O3:Q200").ClearContents
    ' 上から順に部分一致で評価。振替(内部移動)を先に置き、集計対象外にする。
    Dim rules As Variant
    rules = Array( _
        Array("ニコス", "振替", "カード引落"), _
        Array("ﾆｺｽ", "振替", "カード引落"), _
        Array("ATM", "振替", "ATM出金"), _
        Array("セブンギンコウ", "振替", "ATM出金"), _
        Array("ｾﾌﾞﾝｷﾞﾝｺｳ", "振替", "ATM出金"), _
        Array("カード手数料", "振替", "手数料"), _
        Array("カード", "振替", "ATM出金"), _
        Array("ｶｰﾄﾞ", "振替", "ATM出金"), _
        Array("パソコン振替", "振替", "口座間移動"), _
        Array("ﾊﾟｿｺﾝ振替", "振替", "口座間移動"), _
        Array("振替", "振替", "口座間移動"), _
        Array("まいばすけっと", "食費", "食料品"), _
        Array("スキヤ", "食費", "外食"), _
        Array("すき家", "食費", "外食"), _
        Array("マクドナルド", "食費", "外食"), _
        Array("スターバックス", "食費", "カフェ・嗜好品"), _
        Array("セブンイレブン", "食費", "食料品"), _
        Array("ローソン", "食費", "食料品"), _
        Array("ﾛ-ｿﾝ", "食費", "食料品"), _
        Array("ファミリーマート", "食費", "食料品"), _
        Array("イオン", "食費", "食料品"), _
        Array("AMAZON", "日用品", "消耗品"), _
        Array("ENEOS", "交通費", "ガソリン"), _
        Array("JR", "交通費", "電車・バス"), _
        Array("Suica", "交通費", "電車・バス"), _
        Array("NETFLIX", "娯楽・交際費", "趣味・レジャー"), _
        Array("DRAMABOX", "娯楽・交際費", "趣味・レジャー"), _
        Array("OCULUS", "娯楽・交際費", "趣味・レジャー"), _
        Array("ANTHROPIC", "教育・教養", "学習・講座"), _
        Array("ドコモ", "通信費", "携帯電話"), _
        Array("ソフトバンク", "通信費", "携帯電話"), _
        Array("NTT", "通信費", "インターネット"), _
        Array("ユニクロ", "被服・美容", "衣服"), _
        Array("薬局", "医療・健康", "病院・薬"), _
        Array("東京電力", "水道光熱費", "電気"), _
        Array("東京ガス", "水道光熱費", "ガス"))
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
    ' Phase 1 の赤い飾りセル（A2:C2 結合＋D2 注記）を解除・消去してからボタンを置く
    shtManual.Range("A2:D2").UnMerge
    With shtManual.Range("A2:D2")
        .ClearContents
        .Interior.ColorIndex = xlNone
        .Borders.LineStyle = xlNone
    End With
    shtManual.Rows(2).RowHeight = 26
    On Error GoTo 0
    Dim topPos As Double, leftPos As Double
    topPos = shtManual.Range("A2").Top + 2
    leftPos = shtManual.Range("A2").Left + 2
    AddButtonSeq shtManual, leftPos, topPos, 140, 22, "取引台帳へ転記", "modLedger.TransferManualToLedger"
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

' フォームボタンを座標指定で配置し、次に置くボタンの左位置(現在+幅+余白)を返す
Private Function AddButtonSeq(ByVal ws As Worksheet, ByVal leftPos As Double, ByVal topPos As Double, _
                             ByVal w As Double, ByVal h As Double, _
                             ByVal caption As String, ByVal macroName As String) As Double
    Dim btn As Button
    Set btn = ws.Buttons.Add(leftPos, topPos, w, h)
    btn.caption = caption
    btn.OnAction = macroName
    btn.Font.Size = 10
    AddButtonSeq = leftPos + w + 8   ' 8pt の間隔をあけて次のボタン位置を返す
End Function

' 名前付き範囲を追加/置換（ブックレベル）
Private Sub AddOrReplaceName(ByVal nm As String, ByVal refersTo As String)
    On Error Resume Next
    ThisWorkbook.Names(nm).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=nm, RefersTo:=refersTo
End Sub
