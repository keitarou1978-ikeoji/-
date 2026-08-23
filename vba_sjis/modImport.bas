Attribute VB_Name = "modImport"
Option Explicit

'==================================================================
' modImport （Phase 2）
' 銀行口座／クレジットカードの CSV を取り込み、CSV取込シートへ
' プレビュー表示（自動分類つき）→ 確認後に取引台帳へ追記する。
' 列マッピングは設定シートの「CSV列マッピング」を参照する。
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
Public Const IM_TYPE As Long = 9      ' I: 種別
Public Const IM_SRCROW As Long = 10   ' J: 元行
Public Const IM_DUP As Long = 11      ' K: 重複?
Public Const IMPORT_LAST_COL As Long = 11
Public Const IMPORT_HEADER_ROW As Long = 4
Public Const IMPORT_FIRST_ROW As Long = 5

Public Const KIND_ACCOUNT As String = "口座"
Public Const KIND_CARD As String = "カード"
Public Const UNCLASSIFIED As String = "未分類"

' ボタン割当用：口座CSVを取込
Public Sub ImportAccountCSV()
    ImportCSV KIND_ACCOUNT
End Sub

' ボタン割当用：カードCSVを取込
Public Sub ImportCardCSV()
    ImportCSV KIND_CARD
End Sub

'------------------------------------------------------------------
' CSV取込本体（種別ごと）
'------------------------------------------------------------------
Private Sub ImportCSV(ByVal kind As String)
    On Error GoTo ErrHandler
    Dim prevScreen As Boolean
    Dim prevCalc As XlCalculation
    prevScreen = Application.ScreenUpdating
    prevCalc = Application.Calculation

    Dim ws As Worksheet
    Set ws = modLedger.GetSheetByName(SHEET_IMPORT)
    If ws Is Nothing Then
        MsgBox "CSV取込シートがありません。先に modSetup.Setup_Phase2 を実行してください。", _
               vbExclamation, "CSV取込"
        Exit Sub
    End If

    ' ファイル選択
    Dim filePath As Variant
    filePath = Application.GetOpenFilename( _
        "CSVファイル (*.csv),*.csv,すべてのファイル (*.*),*.*", , kind & "CSVを選択")
    If VarType(filePath) = vbBoolean Then Exit Sub  ' キャンセル

    Dim mapping As Object
    Set mapping = GetMapping()

    Dim charset As String
    charset = MapVal(mapping, kind, "文字コード", "Shift_JIS")
    Dim headerRows As Long
    headerRows = CLng(Val(MapVal(mapping, kind, "見出し行数", "1")))
    Dim colDate As Long, colDesc As Long, colOut As Long, colIn As Long, colAmt As Long
    colDate = CLng(Val(MapVal(mapping, kind, "日付列", "1")))
    colDesc = CLng(Val(MapVal(mapping, kind, "摘要列", "2")))
    colOut = CLng(Val(MapVal(mapping, kind, "出金列", "0")))
    colIn = CLng(Val(MapVal(mapping, kind, "入金列", "0")))
    colAmt = CLng(Val(MapVal(mapping, kind, "金額列", "0")))
    Dim defPay As String
    If kind = KIND_ACCOUNT Then
        defPay = MapVal(mapping, kind, "既定支払方法", "メガバンク口座引落")
    Else
        defPay = MapVal(mapping, kind, "既定支払方法", "クレジットカード")
    End If

    ' ファイル読込
    Dim lines As Variant
    lines = ReadTextLines(CStr(filePath), charset)
    If Not IsArray(lines) Then
        MsgBox "ファイルを読み込めませんでした。", vbExclamation, "CSV取込"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' 既存プレビューの末尾に追記（複数回取込に対応）
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
    Dim importedCount As Long

    For i = LBound(lines) To UBound(lines)
        If i < headerRows Then GoTo NextLine        ' 見出し行スキップ
        If Len(Trim$(CStr(lines(i)))) = 0 Then GoTo NextLine

        fields = ParseCsvLine(CStr(lines(i)))
        dVal = ParseDateFlexible(GetField(fields, colDate))
        If Not IsDate(dVal) Then GoTo NextLine       ' 日付が取れない行は除外

        descVal = GetField(fields, colDesc)

        ' 金額と収支区分の決定
        If kind = KIND_ACCOUNT Then
            Dim outAmt As Double, inAmt As Double
            outAmt = ParseAmount(GetField(fields, colOut))
            inAmt = ParseAmount(GetField(fields, colIn))
            If outAmt > 0 Then
                amtVal = outAmt: inexp = "支出"
            ElseIf inAmt > 0 Then
                amtVal = inAmt: inexp = "収入"
            Else
                GoTo NextLine                        ' 金額なし行は除外
            End If
        Else
            amtVal = ParseAmount(GetField(fields, colAmt))
            If amtVal <= 0 Then GoTo NextLine
            inexp = "支出"                           ' カードは基本支出
        End If

        ' 自動分類
        cls = Classify(descVal)

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
        ws.Cells(writeRow, IM_TYPE).Value = kind
        ws.Cells(writeRow, IM_SRCROW).Value = i + 1
        ' 未分類は目立たせる
        If cls(0) = UNCLASSIFIED Then
            ws.Range(ws.Cells(writeRow, IM_MAJOR), ws.Cells(writeRow, IM_MIDDLE)).Interior.Color = RGB(255, 242, 204)
        End If
        importedCount = importedCount + 1
NextLine:
    Next i

    Application.ScreenUpdating = prevScreen
    Application.Calculation = prevCalc
    MsgBox kind & "CSV を読み込みました：" & importedCount & " 行。" & vbCrLf & _
           "内容（特に未分類の黄色セル）を確認し、「取引台帳へ取込」を押してください。", _
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

    Dim lastRow As Long
    Dim r As Long
    Dim records As Collection
    Dim unclassified As Long
    Set records = New Collection
    lastRow = ws.Cells(ws.Rows.Count, IM_DATE).End(xlUp).Row

    For r = IMPORT_FIRST_ROW To lastRow
        If Len(Trim$(CStr(ws.Cells(r, IM_CHECK).Value))) = 0 Then GoTo ContinueRow  ' ○なしは除外
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

    ' 取込に成功した内容はプレビューから消す（重複も処理済みなので消す）
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

' プレビュー明細をクリアする（内部処理）
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

' 摘要から自動分類ルールを適用し Array(大分類, 中分類) を返す
Public Function Classify(ByVal desc As String) As Variant
    On Error GoTo ErrHandler
    Dim ruleRange As Range
    Dim r As Long
    Dim kw As String
    Set ruleRange = Range(NR_CLASSIFY_RULE)
    For r = 1 To ruleRange.Rows.Count
        kw = Trim$(CStr(ruleRange.Cells(r, 1).Value))
        If kw <> "" Then
            If InStr(1, desc, kw, vbTextCompare) > 0 Then
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

' CSV列マッピングを辞書（"種別|項目" -> 値）として読み込む
Public Function GetMapping() As Object
    On Error GoTo ErrHandler
    Dim dict As Object
    Dim rng As Range
    Dim r As Long
    Dim kind As String, item As String, val As String
    Set dict = CreateObject("Scripting.Dictionary")
    Set rng = Range(NR_CSV_MAPPING)
    For r = 1 To rng.Rows.Count
        kind = Trim$(CStr(rng.Cells(r, 1).Value))
        item = Trim$(CStr(rng.Cells(r, 2).Value))
        val = Trim$(CStr(rng.Cells(r, 3).Value))
        If kind <> "" And item <> "" Then dict(kind & "|" & item) = val
    Next r
    Set GetMapping = dict
    Exit Function
ErrHandler:
    Set GetMapping = CreateObject("Scripting.Dictionary")
End Function

' マッピング値の取得（無ければ既定値）
Private Function MapVal(ByVal dict As Object, ByVal kind As String, _
                        ByVal item As String, ByVal defaultVal As String) As String
    Dim k As String
    k = kind & "|" & item
    If dict.Exists(k) Then
        If Trim$(CStr(dict(k))) <> "" Then
            MapVal = CStr(dict(k))
            Exit Function
        End If
    End If
    MapVal = defaultVal
End Function

' テキストファイルを行配列で読み込む（文字コード指定・ADODB.Stream）
Public Function ReadTextLines(ByVal path As String, ByVal charset As String) As Variant
    On Error GoTo ErrHandler
    Dim st As Object
    Dim content As String
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                 ' adTypeText
    st.charset = charset
    st.Open
    st.LoadFromFile path
    content = st.ReadText(-1)    ' adReadAll
    st.Close
    content = Replace(content, vbCrLf, vbLf)
    content = Replace(content, vbCr, vbLf)
    ReadTextLines = Split(content, vbLf)
    Exit Function
ErrHandler:
    ReadTextLines = False
End Function

' 1行を CSV としてフィールド配列（0基点）に分解する（"..."内のカンマ・""対応）
Public Function ParseCsvLine(ByVal line As String) As Variant
    Dim result As Collection
    Dim i As Long
    Dim ch As String
    Dim cur As String
    Dim inQuotes As Boolean
    Dim arr() As String
    Dim n As Long

    Set result = New Collection
    inQuotes = False
    cur = ""
    For i = 1 To Len(line)
        ch = Mid$(line, i, 1)
        If inQuotes Then
            If ch = """" Then
                If i < Len(line) And Mid$(line, i + 1, 1) = """" Then
                    cur = cur & """"
                    i = i + 1
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
                result.Add cur
                cur = ""
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

' 金額文字列を数値化（カンマ・通貨記号・全角空白・括弧マイナス対応、絶対値）
Public Function ParseAmount(ByVal s As String) As Double
    On Error GoTo ErrHandler
    Dim t As String
    t = Trim$(s)
    If t = "" Then ParseAmount = 0: Exit Function
    t = Replace(t, ",", "")
    t = Replace(t, "，", "")
    t = Replace(t, "\", "")
    t = Replace(t, "￥", "")
    t = Replace(t, "円", "")
    t = Replace(t, " ", "")
    t = Replace(t, "　", "")
    t = Replace(t, "(", "-")
    t = Replace(t, ")", "")
    t = Replace(t, "（", "-")
    t = Replace(t, "）", "")
    If IsNumeric(t) Then
        ParseAmount = Abs(CDbl(t))
    Else
        ParseAmount = 0
    End If
    Exit Function
ErrHandler:
    ParseAmount = 0
End Function

' 各種書式の日付文字列を Date へ（失敗時は空文字を返す）
Public Function ParseDateFlexible(ByVal s As String) As Variant
    On Error GoTo ErrHandler
    Dim t As String
    t = Trim$(s)
    If t = "" Then ParseDateFlexible = "": Exit Function
    t = Replace(t, ".", "/")
    t = Replace(t, "-", "/")
    ' YYYYMMDD（8桁数字）
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
