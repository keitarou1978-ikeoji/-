Attribute VB_Name = "modLedger"
Option Explicit

'==================================================================
' modLedger （Phase 2）
' 取引台帳（正本）への追記・取引ID採番・重複排除、
' および手入力シートからの台帳転記（本接続）を担うモジュール。
' 取引台帳/CSV取込シートは実行時生成のため、CodeName ではなく
' シート名で参照する（modSetup.Setup_Phase2 で生成）。
'==================================================================

' --- Phase 2 シート名 ---
Public Const SHEET_LEDGER As String = "取引台帳"
Public Const SHEET_IMPORT As String = "CSV取込"

' --- 取引台帳の列位置 ---
Public Const LG_ID As Long = 1        ' A: 取引ID
Public Const LG_DATE As Long = 2      ' B: 日付
Public Const LG_MAJOR As Long = 3     ' C: 大分類
Public Const LG_MIDDLE As Long = 4    ' D: 中分類
Public Const LG_DESC As Long = 5      ' E: 摘要
Public Const LG_AMOUNT As Long = 6    ' F: 金額
Public Const LG_PAY As Long = 7       ' G: 支払方法
Public Const LG_INEXP As Long = 8     ' H: 収支区分
Public Const LG_SRC As Long = 9       ' I: 入力元
Public Const LG_TS As Long = 10       ' J: 取込日時
Public Const LG_KEY As Long = 11      ' K: 重複キー
Public Const LEDGER_LAST_COL As Long = 11
Public Const LEDGER_HEADER_ROW As Long = 1
Public Const LEDGER_FIRST_ROW As Long = 2

' 指定名のシートを返す（無ければ Nothing）
Public Function GetSheetByName(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheetByName = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

' 取引台帳シートを返す。準備できていなければ Nothing を返し、案内を表示する。
Public Function LedgerSheetOrWarn() As Worksheet
    Dim ws As Worksheet
    Set ws = GetSheetByName(SHEET_LEDGER)
    If ws Is Nothing Then
        MsgBox "取引台帳シートがありません。先に modSetup.Setup_Phase2 を実行してください。", _
               vbExclamation, "取引台帳"
        Set LedgerSheetOrWarn = Nothing
    Else
        Set LedgerSheetOrWarn = ws
    End If
End Function

' 重複判定キーを生成する（日付＋金額＋摘要先頭20字＋支払方法）
Public Function BuildDupKey(ByVal d As Variant, ByVal amount As Variant, _
                            ByVal desc As String, ByVal pay As String) As String
    Dim dateStr As String
    If IsDate(d) Then
        dateStr = Format$(CDate(d), "yyyymmdd")
    Else
        dateStr = Trim$(CStr(d))
    End If
    BuildDupKey = dateStr & "|" & Trim$(CStr(amount)) & "|" & _
                  Left$(Trim$(desc), 20) & "|" & Trim$(pay)
End Function

' 既存の重複キーを辞書に読み込む
Public Function LoadExistingKeys(ByVal ws As Worksheet) As Object
    On Error GoTo ErrHandler
    Dim dict As Object
    Dim lastRow As Long
    Dim r As Long
    Dim k As String
    Set dict = CreateObject("Scripting.Dictionary")
    lastRow = ws.Cells(ws.Rows.Count, LG_ID).End(xlUp).Row
    For r = LEDGER_FIRST_ROW To lastRow
        k = Trim$(CStr(ws.Cells(r, LG_KEY).Value))
        If k <> "" And Not dict.Exists(k) Then dict.Add k, r
    Next r
    Set LoadExistingKeys = dict
    Exit Function
ErrHandler:
    Set LoadExistingKeys = CreateObject("Scripting.Dictionary")
End Function

' 次の取引ID（T00001 形式）を返す
Public Function NextTransactionId(ByVal ws As Worksheet) As String
    On Error GoTo ErrHandler
    Dim lastRow As Long
    Dim r As Long
    Dim maxNum As Long
    Dim idStr As String
    Dim numPart As String
    lastRow = ws.Cells(ws.Rows.Count, LG_ID).End(xlUp).Row
    maxNum = 0
    For r = LEDGER_FIRST_ROW To lastRow
        idStr = Trim$(CStr(ws.Cells(r, LG_ID).Value))
        If Left$(idStr, 1) = "T" And Len(idStr) > 1 Then
            numPart = Mid$(idStr, 2)
            If IsNumeric(numPart) Then
                If CLng(numPart) > maxNum Then maxNum = CLng(numPart)
            End If
        End If
    Next r
    NextTransactionId = "T" & Format$(maxNum + 1, "00000")
    Exit Function
ErrHandler:
    NextTransactionId = "T" & Format$(1, "00000")
End Function

'------------------------------------------------------------------
' レコード配列を台帳へ追記する共通処理。
' records: 各要素が Array(日付,大分類,中分類,摘要,金額,支払方法,収支区分,入力元) の配列。
' addedCount / dupCount を ByRef で返す。
'------------------------------------------------------------------
Public Sub AppendRecords(ByVal records As Collection, _
                         ByRef addedCount As Long, ByRef dupCount As Long)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Dim keys As Object
    Dim rec As Variant
    Dim writeRow As Long
    Dim dupKey As String
    Dim seq As Long
    Dim nowTs As Date

    addedCount = 0
    dupCount = 0
    Set ws = LedgerSheetOrWarn()
    If ws Is Nothing Then Exit Sub

    Set keys = LoadExistingKeys(ws)
    writeRow = ws.Cells(ws.Rows.Count, LG_ID).End(xlUp).Row
    If writeRow < LEDGER_FIRST_ROW Then writeRow = LEDGER_FIRST_ROW - 1
    seq = 0
    ' 既存の最大番号を取得（IDの連番付与用）
    Dim baseNum As Long
    baseNum = CLng(Mid$(NextTransactionId(ws), 2)) - 1
    nowTs = Now

    For Each rec In records
        dupKey = BuildDupKey(rec(0), rec(4), CStr(rec(3)), CStr(rec(5)))
        If keys.Exists(dupKey) Then
            dupCount = dupCount + 1
        Else
            writeRow = writeRow + 1
            seq = seq + 1
            ws.Cells(writeRow, LG_ID).Value = "T" & Format$(baseNum + seq, "00000")
            ws.Cells(writeRow, LG_DATE).Value = rec(0)
            ws.Cells(writeRow, LG_DATE).NumberFormatLocal = "yyyy/mm/dd"
            ws.Cells(writeRow, LG_MAJOR).Value = rec(1)
            ws.Cells(writeRow, LG_MIDDLE).Value = rec(2)
            ws.Cells(writeRow, LG_DESC).Value = rec(3)
            ws.Cells(writeRow, LG_AMOUNT).Value = rec(4)
            ws.Cells(writeRow, LG_AMOUNT).NumberFormatLocal = "\#,##0"
            ws.Cells(writeRow, LG_PAY).Value = rec(5)
            ws.Cells(writeRow, LG_INEXP).Value = rec(6)
            ws.Cells(writeRow, LG_SRC).Value = rec(7)
            ws.Cells(writeRow, LG_TS).Value = nowTs
            ws.Cells(writeRow, LG_TS).NumberFormatLocal = "yyyy/mm/dd hh:mm"
            ws.Cells(writeRow, LG_KEY).Value = dupKey
            keys.Add dupKey, writeRow
            addedCount = addedCount + 1
        End If
    Next rec
    Exit Sub
ErrHandler:
    MsgBox "台帳への追記でエラーが発生しました: " & Err.Description, vbExclamation, "取引台帳"
End Sub

'------------------------------------------------------------------
' 手入力シート → 取引台帳 転記（本接続）。
' Phase 1 の枠 modManualEntry.TransferToLedger に代わり、
' Setup_Phase2 でボタンにこのプロシージャを割り当てる。
'------------------------------------------------------------------
Public Sub TransferManualToLedger()
    On Error GoTo ErrHandler
    Dim prevScreen As Boolean
    Dim prevCalc As XlCalculation
    prevScreen = Application.ScreenUpdating
    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim ws As Worksheet
    Set ws = LedgerSheetOrWarn()
    If ws Is Nothing Then GoTo CleanExit

    Dim lastRow As Long
    Dim r As Long
    Dim records As Collection
    Dim rowNumbers As Collection
    Dim msg As String
    Dim posted As String

    Set records = New Collection
    Set rowNumbers = New Collection
    lastRow = shtManual.Cells(shtManual.Rows.Count, COL_MAJOR).End(xlUp).Row

    For r = MANUAL_FIRST_ROW To lastRow
        posted = Trim$(CStr(shtManual.Cells(r, COL_POSTED).Value))
        If posted = "済" Then GoTo ContinueRow   ' 既に転記済みはスキップ

        ' 空行はスキップ
        If Len(Trim$(CStr(shtManual.Cells(r, COL_MAJOR).Value))) = 0 _
           And Len(Trim$(CStr(shtManual.Cells(r, COL_AMOUNT).Value))) = 0 Then
            GoTo ContinueRow
        End If

        ' 妥当性チェック（Phase 1 の modManualEntry.ValidateRow を再利用）
        If Not modManualEntry.ValidateRow(r, msg) Then
            Application.ScreenUpdating = prevScreen
            Application.Calculation = prevCalc
            MsgBox "転記を中止しました。" & vbCrLf & msg, vbExclamation, "取引台帳へ転記"
            Exit Sub
        End If

        records.Add Array( _
            shtManual.Cells(r, COL_DATE).Value, _
            Trim$(CStr(shtManual.Cells(r, COL_MAJOR).Value)), _
            Trim$(CStr(shtManual.Cells(r, COL_MIDDLE).Value)), _
            Trim$(CStr(shtManual.Cells(r, COL_DESC).Value)), _
            CLng(shtManual.Cells(r, COL_AMOUNT).Value), _
            Trim$(CStr(shtManual.Cells(r, COL_PAYMENT).Value)), _
            Trim$(CStr(shtManual.Cells(r, COL_INEXP).Value)), _
            "手入力")
        rowNumbers.Add r
ContinueRow:
    Next r

    If records.Count = 0 Then
        Application.ScreenUpdating = prevScreen
        Application.Calculation = prevCalc
        MsgBox "転記対象（未転記行）がありません。", vbInformation, "取引台帳へ転記"
        Exit Sub
    End If

    Dim addedCount As Long, dupCount As Long
    modLedger.AppendRecords records, addedCount, dupCount

    ' 転記対象行に転記済みマークを付ける（追加・重複いずれも処理済み）
    Dim i As Long
    For i = 1 To rowNumbers.Count
        r = rowNumbers.Item(i)
        shtManual.Cells(r, COL_POSTED).Value = "済"
    Next i

CleanExit:
    Application.ScreenUpdating = prevScreen
    Application.Calculation = prevCalc
    If Not (ws Is Nothing) Then
        MsgBox "台帳へ転記しました。" & vbCrLf & _
               "追加: " & addedCount & " 件 / 重複スキップ: " & dupCount & " 件", _
               vbInformation, "取引台帳へ転記"
    End If
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = prevScreen
    Application.Calculation = prevCalc
    MsgBox "転記処理でエラーが発生しました: " & Err.Description, vbExclamation, "取引台帳へ転記"
End Sub
