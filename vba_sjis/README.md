# VBA モジュール（Shift-JIS / インポート用）

日本語版 Excel の VBE は `ファイル → ファイルのインポート` で読み込む `.bas` / `.cls` を
**Shift-JIS(CP932)** として解釈します。`vba/` 配下は UTF-8 のため、そのまま取り込むと
日本語コメント・文字列が文字化けし、構文エラーの原因になります。

**Excel へ取り込むときは、この `vba_sjis/`（Shift-JIS・CRLF）のファイルを使用してください。**

| ファイル | 取り込み方 |
|---|---|
| modConfig.bas / modManualEntry.bas / modUtils.bas | `ファイル → ファイルのインポート` |
| shtManual.cls | 手入力シートのコードウィンドウに、`Option Explicit` 以降を貼り付け |

`vba/`（UTF-8）はコードレビュー・差分確認用の正本です。内容は同一で、
`build/encode_sjis.py` で本フォルダを再生成できます。
