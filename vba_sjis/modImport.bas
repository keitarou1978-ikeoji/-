Attribute VB_Name = "modImport"
Option Explicit

'==================================================================
' modImport （Phase 2 / 複数プロファイル対応版）
' 銀行口座（複数レイアウト）・クレジットカードの CSV を取り込み、
' CSV取込シートへプレビュー（自動分類つき）→ 取引台帳へ追記する。
' 列マッピングは設定シートの「CSV列マッピング」をプロファイル単位で参照。
'==================================================================

' --- CSV取込シートのプレビュー列 ---
Public Const IM_CHECK As Long = 1     ' A: 取込○
Public Const IM_DATE As Long = 2      ' B: 日付
Public Const IM_DESC As Long = 3      ' C: 摘要
Public Const IM_AMOUNT As Long = 4    ' D: 金額
Public Const IM_INEXP As Long = 5     ' E: 収支区分
Public Const IM_MAJOR As Long = 6     ' F: 大分類
Public Const IM_MIDDLE As Long = 7    ' G: 中分類
Public Const IM_PAY As Long = 8       ' H: 支払方法
Public Const IM_TYPE As Long = 9      ' I: プロファイル
Public Const IM_SRCROW As Long = 10   ' J: 元行
Public Const IM_DUP As Long = 11      ' K: 重複?
Public Const IMPORT_LAST_COL As Long = 11
Public Const IMPORT_HEADER_ROW As Long = 4
Public Const IMPORT_FIRST_ROW As Long = 5
Public Const IMPORT_PROFILE_CELL As String = "B2"   ' プロファイル選択セル

Public Const UNCLASSIFIED As String = "未分類"
Public Const TRANSFER_MAJOR As String = "振替"       ' 集計対象外（内部移動）
Public Const INEXP_EXCLUDE As String = "対象外"

' ボタン割当用：選択したプロファイルで CSV を取込
Public Sub ImportSelectedProfile()
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = modLedger.GetSheetByName(SHEET_IMPORT)
    If ws Is Nothing Then
        MsgBox "CSV取込シートがありません。先に modSetup.Setup_Phase2 を実行してください。", _
               vbExclamation, "CSV取込"
        Exit Sub
    End If
    Dim profileName As String
    profileName = Trim$(CStr(ws.Range(IMPORT_PROFILE_CELL).Value))
    If profileName = "" Then
        MsgBox "取込プロファイル（" & IMPORT_PROFILE_CELL & " セル）を選択してください。", _
               vbExclamation, "CSV取込"
        Exit Sub
    End If
    ImportCSV profileName
    Exit Sub
ErrHandler:
    MsgBox "CSV取込でエラーが発生しました: " & Err.Description, vbExclamation, "CSV取込"
End Sub

'------------------------------------------------------------------
' CSV取込本体（プロファイル単位）
'------------------------------------------------------------------
Private Sub ImportCSV(ByVal profileName As String)
    On Error GoTo ErrHandler
    Dim prevScreen As Boolean
    Dim prevCalc As XlCalculation
    prevScreen = Application.ScreenUpdating
    prevCalc = Application.Calculation

    Dim ws As Worksheet
    Set ws = modLedger.GetSheetByName(SHEET_IMPORT)
    If ws Is Nothing Then Exit Sub

    ' マッピング取得
    Dim mapping As Object
    Set mapping = GetMapping()
    Dim kind As String, charset As String
    Dim headerRows As Long
    Dim colDate As Long, colDesc As Long, colOut As Long, colIn As Long, colAmt As Long
    Dim defPay As String
    kind = MapVal(mapping, profileName, "種類", "口座")
    charset = MapVal(mapping, profileName, "文字コード", "Shift_JIS")
    headerRows = CLng(Val(MapVal(mapping, profileName, "見出し行数", "1")))
    colDate = CLng(Val(MapVal(mapping, profileName, "日付列", "1")))
    colDesc = CLng(Val(MapVal(mapping, profileName, "摘要列", "2")))
    colOut = CLng(Val(MapVal(mapping, profileName, "出金列", "0")))
    colIn = CLng(Val(MapVal(mapping, profileName, "入金列", "0")))
    colAmt = CLng(Val(MapVal(mapping, profileName, "金額列", "0")))
    defPay = MapVal(mapping, profileName, "既定支払方法", _
                    IIf(kind = "カード", "クレジットカード", "メガバンク口座引落"))

    ' ファイル選択
    Dim filePath As Variant
    filePath = Application.GetOpenFilename( _
        "CSVファイル (*.csv),*.csv,すべてのファイル (*.*),*.*", , profileName & " のCSVを選択")
    If VarType(filePath) = vbBoolean Then Exit Sub

    ' 読込
    Dim lines As Variant
    lines = ReadTextLines(CStr(filePath), charset)
    If Not IsArray(lines) Then
        MsgBox "ファイルを読み込めませんでした。", vbExclamation, "CSV取込"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim writeRow As Long
    writeRow = ws.Cells(ws.Rows.Count, IM_DATE).End(xlUp).Row
    If writeRow < IMPORT_FIRST_ROW Then writeRow = IMPORT_FIRST_ROW - 1

    Dim i As Long
    Dim fields As Variant
    Dim dVal As Variant
    Dim descVal As String
    Dim amtVal As Double
    Dim inexp As String
    Dim cls As Variant
    Dim importedCount As Long, transferCount As Long

    For i = LBound(lines) To UBound(lines)
        If i < headerRows Then GoTo NextLine
        If Len(Trim$(CStr(lines(i)))) = 0 Then GoTo NextLine

        fields = ParseCsvLine(CStr(lines(i)))
        dVal = ParseDateFlexible(GetField(fields, colDate))
        If Not IsDate(dVal) Then GoTo NextLine

        descVal = GetField(fields, colDesc)

        If kind = "カード" Then
            Dim signed As Double
            signed = ParseAmountSigned(GetField(fields, colAmt))
            If signed = 0 Then GoTo NextLine
            If signed < 0 Then
                amtVal = Abs(signed): inexp = "収入"   ' 返金・キャンセル
            Else
                amtVal = signed: inexp = "支出"
            End If
        Else
            Dim outAmt As Double, inAmt As Double
            outAmt = ParseAmount(GetField(fields, colOut))
            inAmt = ParseAmount(GetField(fields, colIn))
            If outAmt > 0 Then
                amtVal = outAmt: inexp = "支出"
            ElseIf inAmt > 0 Then
                amtVal = inAmt: inexp = "収入"
            Else
                GoTo NextLine
            End If
        End If

        ' 自動分類（振替判定を含む）
        cls = Classify(descVal)
        If cls(0) = TRANSFER_MAJOR Then inexp = INEXP_EXCLUDE   ' 内部移動は集計対象外

        writeRow = writeRow + 1
        ws.Cells(writeRow, IM_CHECK).Value = "○"
        ws.Cells(writeRow, IM_DATE).Value = dVal
        ws.Cells(writeRow, IM_DATE).NumberFormatLocal = "yyyy/mm/dd"
        ws.Cells(writeRow, IM_DESC).Value = descVal
        ws.Cells(writeRow, IM_AMOUNT).Value = amtVal
        ws.Cells(writeRow, IM_AMOUNT).NumberFormatLocal = "\#,##0"
        ws.Cells(writeRow, IM_INEXP).Value = inexp
        ws.Cells(writeRow, IM_MAJOR).Value = cls(0)
        ws.Cells(writeRow, IM_MIDDLE).Value = cls(1)
        ws.Cells(writeRow, IM_PAY).Value = defPay
        ws.Cells(writeRow, IM_TYPE).Value = profileName
        ws.Cells(writeRow, IM_SRCROW).Value = i + 1
        If cls(0) = UNCLASSIFIED Then
            ws.Range(ws.Cells(writeRow, IM_MAJOR), ws.Cells(writeRow, IM_MIDDLE)).Interior.Color = RGB(255, 242, 204)
            importedCount = importedCount + 1
        ElseIf cls(0) = TRANSFER_MAJOR Then
            ws.Range(ws.Cells(writeRow, IM_MAJOR), ws.Cells(writeRow, IM_INEXP)).Interior.Color = RGB(217, 217, 217)
            transferCount = transferCount + 1
        Else
            importedCount = importedCount + 1
        End If
    NextLine:
    Next i

    Application.ScreenUpdating = prevScreen
    Application.Calculation = prevCalc
    MsgBox profileName & " を読み込みました。" & vbCrLf & _
           "通常明細: " & importedCount & " 行 / 振替(灰色=集計対象外): " & transferCount & " 行" & vbCrLf & _
           "黄色=未分類 を確認/修正し、「取引台帳へ取込」を押してください。", _
           vbInformation, "CSV取込"
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = prevScreen
    Application.Calculation = prevCalc
    MsgBox "CSV取込でエラーが発生しました: " & Err.Description, vbExclamation, "CSV取込"
End Sub

'------------------------------------------------------------------
' プレビュー（○行）を取引台帳へ取込む。ボタン割当用。
'------------------------------------------------------------------
Public Sub CommitPreviewToLedger()
    On Error GoTo ErrHandler
    Dim prevScreen As Boolean
    Dim prevCalc As XlCalculation
    prevScreen = Application.ScreenUpdating
    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim ws As Worksheet
    Set ws = modLedger.GetSheetByName(SHEET_IMPORT)
    If ws Is Nothing Then GoTo CleanExit

    Dim lastRow As Long, r As Long
    Dim records As Collection
    Dim unclassified As Long
    Set records = New Collection
    lastRow = ws.Cells(ws.Rows.Count, IM_DATE).End(xlUp).Row

    For r = IMPORT_FIRST_ROW To lastRow
        If Len(Trim$(CStr(ws.Cells(r, IM_CHECK).Value))) = 0 Then GoTo ContinueRow
        If Len(Trim$(CStr(ws.Cells(r, IM_DATE).Value))) = 0 Then GoTo ContinueRow
        If Trim$(CStr(ws.Cells(r, IM_MAJOR).Value)) = UNCLASSIFIED Then unclassified = unclassified + 1

        records.Add Array( _
            ws.Cells(r, IM_DATE).Value, _
            Trim$(CStr(ws.Cells(r, IM_MAJOR).Value)), _
            Trim$(CStr(ws.Cells(r, IM_MIDDLE).Value)), _
            Trim$(CStr(ws.Cells(r, IM_DESC).Value)), _
            CLng(ws.Cells(r, IM_AMOUNT).Value), _
            Trim$(CStr(ws.Cells(r, IM_PAY).Value)), _
            Trim$(CStr(ws.Cells(r, IM_INEXP).Value)), _
            "CSV(" & Trim$(CStr(ws.Cells(r, IM_TYPE).Value)) & ")")
    ContinueRow:
    Next r

    If records.Count = 0 Then
        Application.ScreenUpdating = prevScreen
        Application.Calculation = prevCalc
        MsgBox "取込対象（○のある行）がありません。", vbInformation, "取引台帳へ取込"
        Exit Sub
    End If

    Dim addedCount As Long, dupCount As Long
    modLedger.AppendRecords records, addedCount, dupCount
    ClearPreviewRows ws

CleanExit:
    Application.ScreenUpdating = prevScreen
    Application.Calculation = prevCalc
    MsgBox "取引台帳へ取込みました。" & vbCrLf & _
           "追加: " & addedCount & " 件 / 重複スキップ: " & dupCount & " 件" & _
           IIf(unclassified > 0, vbCrLf & "（未分類のまま取込: " & unclassified & " 件）", ""), _
           vbInformation, "取引台帳へ取込"
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = prevScreen
    Application.Calculation = prevCalc
    MsgBox "取込処理でエラーが発生しました: " & Err.Description, vbExclamation, "取引台帳へ取込"
End Sub

' プレビュー消去（ボタン割当用）
Public Sub ClearPreview()
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = modLedger.GetSheetByName(SHEET_IMPORT)
    If ws Is Nothing Then Exit Sub
    If MsgBox("プレビューの内容をすべて消去しますか？", vbQuestion + vbYesNo, "プレビュー消去") = vbYes Then
        ClearPreviewRows ws
    End If
    Exit Sub
ErrHandler:
End Sub

Private Sub ClearPreviewRows(ByVal ws As Worksheet)
    On Error Resume Next
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, IM_DATE).End(xlUp).Row
    If lastRow >= IMPORT_FIRST_ROW Then
        With ws.Range(ws.Cells(IMPORT_FIRST_ROW, IM_CHECK), ws.Cells(lastRow, IMPORT_LAST_COL))
            .ClearContents
            .Interior.ColorIndex = xlNone
        End With
    End If
End Sub

'==================================================================
' ヘルパー
'==================================================================

' 摘要から自動分類ルールを適用し Array(大分類, 中分類) を返す（全半角を吸収して部分一致）
Public Function Classify(ByVal desc As String) As Variant
    On Error GoTo ErrHandler
    Dim ruleRange As Range
    Dim r As Long
    Dim kw As String
    Dim nDesc As String
    Set ruleRange = Range(NR_CLASSIFY_RULE)
    nDesc = Narrow(desc)
    For r = 1 To ruleRange.Rows.Count
        kw = Trim$(CStr(ruleRange.Cells(r, 1).Value))
        If kw <> "" Then
            If InStr(1, nDesc, Narrow(kw), vbTextCompare) > 0 Then
                Classify = Array(Trim$(CStr(ruleRange.Cells(r, 2).Value)), _
                                 Trim$(CStr(ruleRange.Cells(r, 3).Value)))
                Exit Function
            End If
        End If
    Next r
    Classify = Array(UNCLASSIFIED, UNCLASSIFIED)
    Exit Function
ErrHandler:
    Classify = Array(UNCLASSIFIED, UNCLASSIFIED)
End Function

' CSV列マッピングを辞書（"プロファイル|項目" -> 値）として読み込む
Public Function GetMapping() As Object
    On Error GoTo ErrHandler
    Dim dict As Object
    Dim rng As Range
    Dim r As Long
    Dim prof As String, item As String, val As String
    Set dict = CreateObject("Scripting.Dictionary")
    Set rng = Range(NR_CSV_MAPPING)
    For r = 1 To rng.Rows.Count
        prof = Trim$(CStr(rng.Cells(r, 1).Value))
        item = Trim$(CStr(rng.Cells(r, 2).Value))
        val = Trim$(CStr(rng.Cells(r, 3).Value))
        If prof <> "" And item <> "" Then dict(prof & "|" & item) = val
    Next r
    Set GetMapping = dict
    Exit Function
ErrHandler:
    Set GetMapping = CreateObject("Scripting.Dictionary")
End Function

Private Function MapVal(ByVal dict As Object, ByVal prof As String, _
                        ByVal item As String, ByVal defaultVal As String) As String
    Dim k As String
    k = prof & "|" & item
    If dict.Exists(k) Then
        If Trim$(CStr(dict(k))) <> "" Then
            MapVal = CStr(dict(k))
            Exit Function
        End If
    End If
    MapVal = defaultVal
End Function

' 全角英数字・記号を半角へ（分類の全半角ゆれ吸収用）
Public Function Narrow(ByVal s As String) As String
    On Error GoTo ErrHandler
    Narrow = StrConv(s, vbNarrow)
    Exit Function
ErrHandler:
    Narrow = s
End Function

' テキストファイルを行配列で読み込む（文字コード指定・ADODB.Stream）
Public Function ReadTextLines(ByVal path As String, ByVal charset As String) As Variant
    On Error GoTo ErrHandler
    Dim st As Object
    Dim content As String
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.charset = charset
    st.Open
    st.LoadFromFile path
    content = st.ReadText(-1)
    st.Close
    content = Replace(content, vbCrLf, vbLf)
    content = Replace(content, vbCr, vbLf)
    ReadTextLines = Split(content, vbLf)
    Exit Function
ErrHandler:
    ReadTextLines = False
End Function

' 1行を CSV としてフィールド配列（0基点）に分解（"..."内カンマ・""対応）
Public Function ParseCsvLine(ByVal line As String) As Variant
    Dim result As Collection
    Dim i As Long, ch As String, cur As String
    Dim inQuotes As Boolean
    Dim arr() As String, n As Long
    Set result = New Collection
    inQuotes = False
    cur = ""
    For i = 1 To Len(line)
        ch = Mid$(line, i, 1)
        If inQuotes Then
            If ch = """" Then
                If i < Len(line) And Mid$(line, i + 1, 1) = """" Then
                    cur = cur & """": i = i + 1
                Else
                    inQuotes = False
                End If
            Else
                cur = cur & ch
            End If
        Else
            If ch = """" Then
                inQuotes = True
            ElseIf ch = "," Then
                result.Add cur: cur = ""
            Else
                cur = cur & ch
            End If
        End If
    Next i
    result.Add cur
    ReDim arr(0 To result.Count - 1)
    For n = 1 To result.Count
        arr(n - 1) = result.Item(n)
    Next n
    ParseCsvLine = arr
End Function

' フィールド配列から1基点の列番号で安全に取得（前後空白除去）
Public Function GetField(ByVal fields As Variant, ByVal colNum As Long) As String
    On Error GoTo ErrHandler
    If Not IsArray(fields) Then GetField = "": Exit Function
    If colNum < 1 Then GetField = "": Exit Function
    If colNum - 1 > UBound(fields) Then GetField = "": Exit Function
    GetField = Trim$(CStr(fields(colNum - 1)))
    Exit Function
ErrHandler:
    GetField = ""
End Function

' 金額文字列を数値化（絶対値。カンマ・通貨記号・全角対応）
Public Function ParseAmount(ByVal s As String) As Double
    ParseAmount = Abs(ParseAmountSigned(s))
End Function

' 金額文字列を数値化（符号つき。括弧・先頭マイナスを負とみなす）
Public Function ParseAmountSigned(ByVal s As String) As Double
    On Error GoTo ErrHandler
    Dim t As String
    Dim neg As Boolean
    t = Narrow(Trim$(s))
    If t = "" Then ParseAmountSigned = 0: Exit Function
    t = Replace(t, ",", "")
    t = Replace(t, "\", "")
    t = Replace(t, "￥", "")
    t = Replace(t, "円", "")
    t = Replace(t, " ", "")
    t = Replace(t, "　", "")
    neg = (Left$(t, 1) = "-") Or (Left$(t, 1) = "(") Or (Left$(t, 1) = "（")
    t = Replace(t, "(", ""): t = Replace(t, ")", "")
    t = Replace(t, "（", ""): t = Replace(t, "）", "")
    t = Replace(t, "-", "")
    If IsNumeric(t) Then
        ParseAmountSigned = IIf(neg, -CDbl(t), CDbl(t))
    Else
        ParseAmountSigned = 0
    End If
    Exit Function
ErrHandler:
    ParseAmountSigned = 0
End Function

' 各種書式の日付文字列を Date へ（yyyy年m月d日 / 区切り各種 / YYYYMMDD 対応）
Public Function ParseDateFlexible(ByVal s As String) As Variant
    On Error GoTo ErrHandler
    Dim t As String
    t = Narrow(Trim$(s))
    If t = "" Then ParseDateFlexible = "": Exit Function
    t = Replace(t, "年", "/")
    t = Replace(t, "月", "/")
    t = Replace(t, "日", "")
    t = Replace(t, ".", "/")
    t = Replace(t, "-", "/")
    Do While InStr(t, "//") > 0
        t = Replace(t, "//", "/")
    Loop
    If Right$(t, 1) = "/" Then t = Left$(t, Len(t) - 1)
    If Len(t) = 8 And IsNumeric(t) Then
        t = Left$(t, 4) & "/" & Mid$(t, 5, 2) & "/" & Mid$(t, 7, 2)
    End If
    If IsDate(t) Then
        ParseDateFlexible = CDate(t)
    Else
        ParseDateFlexible = ""
    End If
    Exit Function
ErrHandler:
    ParseDateFlexible = ""
End Function
