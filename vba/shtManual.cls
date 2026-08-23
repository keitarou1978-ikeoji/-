VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'==================================================================
' 手入力シート（CodeName: shtManual）シートモジュール
'  - B列変更 : C列の連動ドロップダウン再構築 ＋ G列の収支区分自動判定
'  - B〜F列入力: A列が空なら今日の日付を既定セット
'==================================================================

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrHandler
    Dim cell As Range
    Dim isInDataArea As Boolean

    ' 明細エリア（A〜F列 かつ 明細開始行以降）以外の変更は無視
    If Intersect(Target, Me.Range(Me.Cells(MANUAL_FIRST_ROW, COL_DATE), _
                                  Me.Cells(Me.Rows.Count, COL_PAYMENT))) Is Nothing Then
        Exit Sub
    End If

    Application.EnableEvents = False

    For Each cell In Target.Cells
        isInDataArea = (cell.Row >= MANUAL_FIRST_ROW)
        If isInDataArea Then
            ' 大分類(B列)が変更されたら連動DDと収支区分を更新
            If cell.Column = COL_MAJOR Then
                modManualEntry.RefreshMiddleDropdown cell
                modManualEntry.ApplyInExpType cell.Row
            End If

            ' B〜F列いずれかへの入力で、A列が空なら今日の日付を既定セット
            If cell.Column >= COL_MAJOR And cell.Column <= COL_PAYMENT Then
                If Len(Trim$(CStr(cell.Value))) > 0 Then
                    modManualEntry.SetDefaultDate cell.Row
                End If
            End If
        End If
    Next cell

CleanExit:
    Application.EnableEvents = True
    Exit Sub
ErrHandler:
    Application.EnableEvents = True
    ' 入力中のエラーは操作を止めない（必要なら下行を有効化）
    ' MsgBox "入力処理でエラー: " & Err.Description, vbExclamation
End Sub
