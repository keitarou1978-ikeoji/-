# -*- coding: utf-8 -*-
"""
vba/ (UTF-8) の VBA モジュールを Shift-JIS(CP932)・CRLF に変換して
vba_sjis/ へ出力する。日本語版 Excel の VBE はインポート時に .bas/.cls を
Shift-JIS として読むため、UTF-8 のままだと文字化け→構文エラーになる。
"""
import os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_DIR = os.path.join(ROOT, "vba")
DST_DIR = os.path.join(ROOT, "vba_sjis")
FILES = ["modConfig.bas", "modManualEntry.bas", "modUtils.bas", "shtManual.cls"]


def main():
    os.makedirs(DST_DIR, exist_ok=True)
    for name in FILES:
        with open(os.path.join(SRC_DIR, name), "r", encoding="utf-8") as fh:
            text = fh.read()
        # cp932 で表現できない文字を検出（あれば警告）
        bad = sorted({ch for ch in text if not _cp932_ok(ch)})
        text_crlf = text.replace("\r\n", "\n").replace("\n", "\r\n")
        with open(os.path.join(DST_DIR, name), "w", encoding="cp932", newline="") as fh:
            fh.write(text_crlf)
        note = f"  !! cp932非対応文字: {bad}" if bad else ""
        print(f"{name} -> vba_sjis/{name}{note}")


def _cp932_ok(ch):
    try:
        ch.encode("cp932")
        return True
    except UnicodeEncodeError:
        return False


if __name__ == "__main__":
    main()
