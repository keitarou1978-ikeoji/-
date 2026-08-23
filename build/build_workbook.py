# -*- coding: utf-8 -*-
"""
人生奪還PJ_資金管理ツール_v1.xlsm を生成するビルドスクリプト（Phase 1）。

生成対象:
  - 設定シート   (CodeName: shtConfig, タブ色: グレー)
  - 手入力シート (CodeName: shtManual, タブ色: 青)
  - 名前付き範囲・入力規則(ドロップダウン)・書式・転記ボタンのプレースホルダ

VBA本体は vba/*.bas, *.cls として別管理（Excelで手動インポート）。
本スクリプトはシート構造・データ・入力補助のみを構築する。
"""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter

# ---- 出力先 ----
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_PATH = os.path.join(ROOT, "人生奪還PJ_資金管理ツール_v1.xlsm")

FONT_NAME = "Meiryo UI"  # 日本語向け。未導入環境では代替フォントで表示される

# ---- カテゴリマスタ（大分類, [中分類...], 収支区分, メモ） ----
CATEGORIES = [
    ("食費", ["食料品", "外食", "カフェ・嗜好品"], "支出", ""),
    ("住居費", ["家賃・ローン", "管理費・修繕積立", "火災保険"], "支出", ""),
    ("水道光熱費", ["電気", "ガス", "水道"], "支出", ""),
    ("通信費", ["携帯電話", "インターネット", "放送・その他"], "支出", ""),
    ("交通費", ["電車・バス", "ガソリン", "タクシー"], "支出", ""),
    ("日用品", ["消耗品", "衛生用品"], "支出", ""),
    ("医療・健康", ["病院・薬", "医療保険", "フィットネス"], "支出", ""),
    ("教育・教養", ["書籍", "学習・講座", "新聞・雑誌"], "支出", ""),
    ("娯楽・交際費", ["趣味・レジャー", "交際費", "旅行"], "支出", ""),
    ("被服・美容", ["衣服", "美容・理容", "化粧品"], "支出", ""),
    ("保険", ["生命保険", "損害保険"], "支出", ""),
    ("税金・社会保険", ["所得税・住民税", "社会保険料"], "支出", ""),
    ("特別費", ["冠婚葬祭", "家電・家具", "その他特別費"], "支出", ""),
    ("給与収入", ["給与", "賞与"], "収入", ""),
    ("事業・副業収入", ["事業所得", "副業"], "収入", ""),
    ("資産収入", ["配当・利息", "不動産収入"], "収入", ""),
    ("その他収入", ["臨時収入", "還付金"], "収入", ""),
]

PAYMENTS = [
    ("現金", "現金"),
    ("メガバンク口座引落", "口座"),
    ("クレジットカード", "カード"),
]

CLASSIFY_RULES = [
    ("セブンイレブン", "食費", "食料品"),
    ("東京電力", "水道光熱費", "電気"),
    ("JR東日本", "交通費", "電車・バス"),
    ("Amazon", "日用品", "消耗品"),
    ("スターバックス", "食費", "カフェ・嗜好品"),
]

# ---- スタイル ----
HDR_FONT = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=11)
HDR_FILL = PatternFill("solid", fgColor="4472C4")
SUBHDR_FILL = PatternFill("solid", fgColor="8EAADB")
TITLE_FONT = Font(name=FONT_NAME, bold=True, size=14)
BASE_FONT = Font(name=FONT_NAME, size=11)
CENTER = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center")
THIN = Side(style="thin", color="B0B0B0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def style_header(cell):
    cell.font = HDR_FONT
    cell.fill = HDR_FILL
    cell.alignment = CENTER
    cell.border = BORDER


def qsheet(name):
    """シート名を安全に引用符付きで返す（名前付き範囲・数式参照用）。"""
    return "'" + name.replace("'", "''") + "'"


def _patch_macro_enabled(path):
    """openpyxl は vbaProject の無いブックを通常 xlsx の content-type で保存するため、
    拡張子(.xlsm)と不一致になる。workbook.xml の content-type を
    macroEnabled に書き換え、Excel が警告なくマクロ有効ブックとして開けるようにする。
    （VBA本体は Excel 上で vba/ 配下のモジュールをインポートして付与する）"""
    import zipfile
    import shutil
    import tempfile

    PLAIN = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
    MACRO = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"

    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".xlsm")
    os.close(tmp_fd)
    with zipfile.ZipFile(path, "r") as zin, \
         zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.replace(PLAIN.encode("utf-8"), MACRO.encode("utf-8"))
            zout.writestr(item, data)
    shutil.move(tmp_path, path)


def build():
    wb = Workbook()

    # ============================================================
    # 設定シート
    # ============================================================
    ws = wb.active
    ws.title = "設定"
    ws.sheet_properties.codeName = "shtConfig"
    ws.sheet_properties.tabColor = "808080"  # グレー
    ws.sheet_view.showGridLines = False

    ws["A1"] = "設定シート — マスタ／システム設定"
    ws["A1"].font = TITLE_FONT

    # --- カテゴリマスタ A〜D（見出し行=2, データ=3〜） ---
    cat_hdr_row = 2
    for col, title in enumerate(["大分類", "中分類", "収支区分", "メモ"], start=1):
        c = ws.cell(row=cat_hdr_row, column=col, value=title)
        style_header(c)

    r = cat_hdr_row + 1
    major_order = []  # 大分類の出現順（重複なし）
    for major, middles, inexp, memo in CATEGORIES:
        if major not in major_order:
            major_order.append(major)
        for mid in middles:
            ws.cell(row=r, column=1, value=major).font = BASE_FONT
            ws.cell(row=r, column=2, value=mid).font = BASE_FONT
            ws.cell(row=r, column=3, value=inexp).font = BASE_FONT
            ws.cell(row=r, column=4, value=memo).font = BASE_FONT
            for col in range(1, 5):
                ws.cell(row=r, column=col).border = BORDER
            r += 1
    cat_last_row = r - 1

    # --- 支払方法マスタ F〜G ---
    pay_hdr_row = 2
    ws.cell(row=pay_hdr_row, column=6, value="支払方法").font = HDR_FONT
    ws.cell(row=pay_hdr_row, column=7, value="種別").font = HDR_FONT
    for col in (6, 7):
        style_header(ws.cell(row=pay_hdr_row, column=col))
    pr = pay_hdr_row + 1
    for pay, kind in PAYMENTS:
        ws.cell(row=pr, column=6, value=pay).font = BASE_FONT
        ws.cell(row=pr, column=7, value=kind).font = BASE_FONT
        ws.cell(row=pr, column=6).border = BORDER
        ws.cell(row=pr, column=7).border = BORDER
        pr += 1
    pay_last_row = pr - 1  # = 4

    # --- CSV列マッピング I〜M（Phase2 用・ヘッダー＋空欄） ---
    csv_hdr_row = 2
    for idx, title in enumerate(["論理名", "CSV列見出し", "列番号", "データ型", "備考"], start=9):
        style_header(ws.cell(row=csv_hdr_row, column=idx, value=title))
    csv_first = csv_hdr_row + 1
    csv_last = csv_first + 19  # 20 行の空欄枠
    for rr in range(csv_first, csv_last + 1):
        for cc in range(9, 14):
            ws.cell(row=rr, column=cc).border = BORDER

    # --- 自動分類ルール O〜Q（Phase2 用・サンプル5件） ---
    rule_hdr_row = 2
    for idx, title in enumerate(["キーワード", "大分類", "中分類"], start=15):
        style_header(ws.cell(row=rule_hdr_row, column=idx, value=title))
    rr = rule_hdr_row + 1
    for kw, mj, md in CLASSIFY_RULES:
        ws.cell(row=rr, column=15, value=kw).font = BASE_FONT
        ws.cell(row=rr, column=16, value=mj).font = BASE_FONT
        ws.cell(row=rr, column=17, value=md).font = BASE_FONT
        for cc in (15, 16, 17):
            ws.cell(row=rr, column=cc).border = BORDER
        rr += 1
    rule_last = rr - 1  # = 7

    # --- システム設定 S〜T ---
    sys_hdr_row = 2
    style_header(ws.cell(row=sys_hdr_row, column=19, value="設定項目"))
    style_header(ws.cell(row=sys_hdr_row, column=20, value="値"))
    sys_items = [
        ("対象年度", 2026, "0"),
        ("浪費アラート閾値", 0.15, "0.0%"),
        ("目標純資産額", 100000000, "\\#,##0"),
        ("達成期限", None, "yyyy/mm/dd"),  # 日付は下で設定
    ]
    import datetime
    sys_first = sys_hdr_row + 1
    sr = sys_first
    for label, val, numfmt in sys_items:
        ws.cell(row=sr, column=19, value=label).font = BASE_FONT
        cell = ws.cell(row=sr, column=20)
        if label == "達成期限":
            cell.value = datetime.date(2033, 8, 31)
        else:
            cell.value = val
        cell.font = BASE_FONT
        cell.number_format = numfmt
        ws.cell(row=sr, column=19).border = BORDER
        cell.border = BORDER
        sr += 1
    sys_last = sr - 1  # 対象年度=T3, 閾値=T4, 目標=T5, 期限=T6

    # --- 大分類リスト（入力補助ヘルパー）V列 ---
    ws.cell(row=2, column=22, value="大分類リスト").font = HDR_FONT
    style_header(ws.cell(row=2, column=22))
    vr = 3
    for major in major_order:
        ws.cell(row=vr, column=22, value=major).font = BASE_FONT
        ws.cell(row=vr, column=22).border = BORDER
        vr += 1
    major_last = vr - 1

    # 列幅
    for col, width in {
        "A": 14, "B": 18, "C": 10, "D": 14, "E": 2,
        "F": 20, "G": 10, "H": 2,
        "I": 16, "J": 16, "K": 8, "L": 10, "M": 16, "N": 2,
        "O": 16, "P": 14, "Q": 14, "R": 2,
        "S": 18, "T": 16, "U": 2, "V": 16,
    }.items():
        ws.column_dimensions[col].width = width

    # ============================================================
    # 手入力シート
    # ============================================================
    ms = wb.create_sheet("手入力")
    ms.sheet_properties.codeName = "shtManual"
    ms.sheet_properties.tabColor = "2E75B6"  # 青
    ms.sheet_view.showGridLines = False

    ms["A1"] = "手入力シート — 日次の入出金を入力"
    ms["A1"].font = TITLE_FONT

    # 転記ボタンのプレースホルダ（Excelでフォームボタンを配置し
    #  modManualEntry.TransferToLedger を割り当てる。詳細は docs/SETUP_VBA.md）
    ms["A2"] = "▶ 取引台帳へ転記"
    ms["A2"].font = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=11)
    ms["A2"].fill = PatternFill("solid", fgColor="C00000")
    ms["A2"].alignment = CENTER
    ms["A2"].border = BORDER
    ms.merge_cells("A2:C2")
    ms["D2"] = "← このセル位置にフォームボタンを配置し modManualEntry.TransferToLedger を割当（Phase 1 は枠）"
    ms["D2"].font = Font(name=FONT_NAME, size=9, color="808080")

    # 見出し行（4行目）
    hdr_row = 4
    first_row = 5
    headers = ["日付", "大分類", "中分類", "摘要", "金額", "支払方法", "収支区分", "台帳転記済"]
    for col, title in enumerate(headers, start=1):
        style_header(ms.cell(row=hdr_row, column=col, value=title))

    # 明細行の書式（1000行分）
    last_row = 1000
    for rr in range(first_row, last_row + 1):
        ms.cell(row=rr, column=1).number_format = "yyyy/mm/dd"  # A:日付
        ms.cell(row=rr, column=5).number_format = "\\#,##0"       # E:金額
        for cc in range(1, 9):
            ms.cell(row=rr, column=cc).font = BASE_FONT

    # 入力例（1行・薄い文字で説明）
    ms.cell(row=first_row, column=4, value="（例）B列で大分類を選ぶとC列が連動、A列は自動で今日の日付")
    ms.cell(row=first_row, column=4).font = Font(name=FONT_NAME, italic=True, color="A6A6A6", size=9)

    # 列幅
    for col, width in {
        "A": 13, "B": 16, "C": 16, "D": 40, "E": 12,
        "F": 20, "G": 10, "H": 12,
    }.items():
        ms.column_dimensions[col].width = width
    ms.freeze_panes = "A5"

    # ============================================================
    # 入力規則（ドロップダウン）
    # ============================================================
    # B列: 大分類（名前付き範囲）
    dv_major = DataValidation(type="list", formula1="=大分類リスト", allow_blank=True, showDropDown=False)
    ms.add_data_validation(dv_major)
    dv_major.add(f"B{first_row}:B{last_row}")

    # F列: 支払方法（名前付き範囲）
    dv_pay = DataValidation(type="list", formula1="=支払方法リスト", allow_blank=True, showDropDown=False)
    ms.add_data_validation(dv_pay)
    dv_pay.add(f"F{first_row}:F{last_row}")

    # C列（中分類）は Worksheet_Change で動的生成のため、ここでは規則を敷かない。

    # ============================================================
    # 名前付き範囲
    # ============================================================
    cfg = qsheet("設定")
    defined = {
        "カテゴリマスタ": f"{cfg}!$A$3:$D${cat_last_row}",
        "支払方法マスタ": f"{cfg}!$F$3:$G${pay_last_row}",
        "支払方法リスト": f"{cfg}!$F$3:$F${pay_last_row}",
        "CSV列マッピング": f"{cfg}!$I${csv_first}:$M${csv_last}",
        "自動分類ルール": f"{cfg}!$O$3:$Q${rule_last}",
        "システム設定": f"{cfg}!$S${sys_first}:$T${sys_last}",
        "大分類リスト": f"{cfg}!$V$3:$V${major_last}",
        "対象年度": f"{cfg}!$T${sys_first}",
        "浪費アラート閾値": f"{cfg}!$T${sys_first + 1}",
        "目標純資産額": f"{cfg}!$T${sys_first + 2}",
        "達成期限": f"{cfg}!$T${sys_first + 3}",
    }
    for name, ref in defined.items():
        wb.defined_names[name] = DefinedName(name, attr_text=ref)

    wb.save(OUT_PATH)
    _patch_macro_enabled(OUT_PATH)
    print("Saved:", OUT_PATH)
    print("カテゴリマスタ:", f"$A$3:$D${cat_last_row}", "(件数=", cat_last_row - 2, ")")
    print("大分類リスト:", f"$V$3:$V${major_last}", "(件数=", major_last - 2, ")")
    print("システム設定:", f"$S${sys_first}:$T${sys_last}")


if __name__ == "__main__":
    build()
