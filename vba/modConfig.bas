Attribute VB_Name = "modConfig"
Option Explicit

'==================================================================
' modConfig
' 設定シートのマスタ・システム設定値を読み込む共通モジュール。
' シート参照は CodeName（shtConfig / shtManual）を使用する。
'==================================================================

' --- シート CodeName（VBEのプロパティで各シートに設定すること） ---
' 設定シート   : shtConfig
' 手入力シート : shtManual

' --- 名前付き範囲の名称（定数） ---
Public Const NR_CATEGORY_MASTER As String = "カテゴリマスタ"
Public Const NR_PAYMENT_MASTER As String = "支払方法マスタ"
Public Const NR_PAYMENT_LIST As String = "支払方法リスト"
Public Const NR_MAJOR_LIST As String = "大分類リスト"
Public Const NR_CSV_MAPPING As String = "CSV列マッピング"
Public Const NR_CLASSIFY_RULE As String = "自動分類ルール"
Public Const NR_SYSTEM_SETTING As String = "システム設定"
Public Const NR_TARGET_YEAR As String = "対象年度"
Public Const NR_WASTE_THRESHOLD As String = "浪費アラート閾値"
Public Const NR_TARGET_NETWORTH As String = "目標純資産額"
Public Const NR_TARGET_DEADLINE As String = "達成期限"

' --- 手入力シートの列位置（定数） ---
Public Const COL_DATE As Long = 1        ' A: 日付
Public Const COL_MAJOR As Long = 2       ' B: 大分類
Public Const COL_MIDDLE As Long = 3      ' C: 中分類
Public Const COL_DESC As Long = 4        ' D: 摘要
Public Const COL_AMOUNT As Long = 5      ' E: 金額
Public Const COL_PAYMENT As Long = 6     ' F: 支払方法
Public Const COL_INEXP As Long = 7       ' G: 収支区分
Public Const COL_POSTED As Long = 8      ' H: 台帳転記済
Public Const MANUAL_HEADER_ROW As Long = 4   ' 手入力シートの見出し行
Public Const MANUAL_FIRST_ROW As Long = 5    ' 手入力シートの明細開始行

' 対象年度を取得する
Public Function GetTargetYear() As Long
    On Error GoTo ErrHandler
    GetTargetYear = CLng(Range(NR_TARGET_YEAR).Value)
    Exit Function
ErrHandler:
    GetTargetYear = Year(Date)   ' 取得失敗時はシステム年で代替
End Function

' 浪費アラート閾値（割合。0.15 = 15%）を取得する
Public Function GetWasteThreshold() As Double
    On Error GoTo ErrHandler
    GetWasteThreshold = CDbl(Range(NR_WASTE_THRESHOLD).Value)
    Exit Function
ErrHandler:
    GetWasteThreshold = 0.15
End Function

' 目標純資産額（円）を取得する
Public Function GetTargetNetWorth() As Double
    On Error GoTo ErrHandler
    GetTargetNetWorth = CDbl(Range(NR_TARGET_NETWORTH).Value)
    Exit Function
ErrHandler:
    GetTargetNetWorth = 0
End Function

' 目標達成期限を取得する
Public Function GetTargetDeadline() As Date
    On Error GoTo ErrHandler
    GetTargetDeadline = CDate(Range(NR_TARGET_DEADLINE).Value)
    Exit Function
ErrHandler:
    GetTargetDeadline = 0
End Function

' 指定した大分類に属する中分類の配列を返す（連動ドロップダウン用）
Public Function GetMiddleCategories(ByVal majorName As String) As Variant
    On Error GoTo ErrHandler
    Dim masterRange As Range
    Dim resultList As Collection
    Dim rowIdx As Long
    Dim currentMajor As String
    Dim currentMiddle As String
    Dim outArr() As String
    Dim i As Long

    Set masterRange = Range(NR_CATEGORY_MASTER)
    Set resultList = New Collection

    For rowIdx = 1 To masterRange.Rows.Count
        currentMajor = Trim$(CStr(masterRange.Cells(rowIdx, 1).Value))
        currentMiddle = Trim$(CStr(masterRange.Cells(rowIdx, 2).Value))
        If currentMajor = majorName And currentMiddle <> "" Then
            resultList.Add currentMiddle
        End If
    Next rowIdx

    If resultList.Count = 0 Then
        GetMiddleCategories = Array()
        Exit Function
    End If

    ReDim outArr(0 To resultList.Count - 1)
    For i = 1 To resultList.Count
        outArr(i - 1) = resultList.Item(i)
    Next i
    GetMiddleCategories = outArr
    Exit Function
ErrHandler:
    GetMiddleCategories = Array()
End Function

' 指定した大分類の収支区分（収入/支出）を返す。見つからなければ空文字。
Public Function GetInExpType(ByVal majorName As String) As String
    On Error GoTo ErrHandler
    Dim masterRange As Range
    Dim rowIdx As Long

    Set masterRange = Range(NR_CATEGORY_MASTER)
    For rowIdx = 1 To masterRange.Rows.Count
        If Trim$(CStr(masterRange.Cells(rowIdx, 1).Value)) = majorName Then
            GetInExpType = Trim$(CStr(masterRange.Cells(rowIdx, 3).Value))
            Exit Function
        End If
    Next rowIdx
    GetInExpType = ""
    Exit Function
ErrHandler:
    GetInExpType = ""
End Function

' 大分類の一覧（重複なし）を配列で返す
Public Function GetMajorList() As Variant
    On Error GoTo ErrHandler
    Dim majorRange As Range
    Dim cell As Range
    Dim seen As Object
    Dim resultList As Collection
    Dim outArr() As String
    Dim i As Long
    Dim v As String

    Set seen = CreateObject("Scripting.Dictionary")
    Set resultList = New Collection
    Set majorRange = Range(NR_CATEGORY_MASTER).Columns(1)

    For Each cell In majorRange.Cells
        v = Trim$(CStr(cell.Value))
        If v <> "" Then
            If Not seen.Exists(v) Then
                seen.Add v, True
                resultList.Add v
            End If
        End If
    Next cell

    If resultList.Count = 0 Then
        GetMajorList = Array()
        Exit Function
    End If

    ReDim outArr(0 To resultList.Count - 1)
    For i = 1 To resultList.Count
        outArr(i - 1) = resultList.Item(i)
    Next i
    GetMajorList = outArr
    Exit Function
ErrHandler:
    GetMajorList = Array()
End Function

' 支払方法の一覧を配列で返す
Public Function GetPaymentList() As Variant
    On Error GoTo ErrHandler
    Dim payRange As Range
    Dim cell As Range
    Dim resultList As Collection
    Dim outArr() As String
    Dim i As Long
    Dim v As String

    Set resultList = New Collection
    Set payRange = Range(NR_PAYMENT_LIST)

    For Each cell In payRange.Cells
        v = Trim$(CStr(cell.Value))
        If v <> "" Then resultList.Add v
    Next cell

    If resultList.Count = 0 Then
        GetPaymentList = Array()
        Exit Function
    End If

    ReDim outArr(0 To resultList.Count - 1)
    For i = 1 To resultList.Count
        outArr(i - 1) = resultList.Item(i)
    Next i
    GetPaymentList = outArr
    Exit Function
ErrHandler:
    GetPaymentList = Array()
End Function
