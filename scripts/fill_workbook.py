"""第2回 SPReAD 様式1 への書き込み

テンプレ xlsx を ZIP として開き、対象シートの空セルだけを XML 文字列置換で
書き換える。テンプレートの構造（リッチテキスト・条件付き書式・データ検証・
名前空間・customXml・printerSettings 等）はバイト単位でそのまま保持される。

書き込み対象:
  - 空セル `<c r="C8" s="318"/>` を
    `<c r="C8" s="318" t="inlineStr"><is><t xml:space="preserve">VALUE</t></is></c>`
    に置換（文字列値）
  - 数値値は `<c r="H6" s="318"><v>2026</v></c>` 形式
  - 既値セル・数式セルは触らない

シートと座標:
  - sheet3.xml = 1枚目（基本情報）
  - sheet4.xml = 2枚目（研究目的等）
  - sheet5.xml = 3枚目（経費）
  - sheet6.xml = 4枚目（API・計算資源）

CLI:
    python3 fill_workbook.py --data payload.json --output 様式1.xlsx
    python3 fill_workbook.py --language en --data payload.json --output Form1.xlsx
"""

from __future__ import annotations
import argparse
import json
import re
import zipfile
from pathlib import Path

# スキルルート（このファイルの親の親）
SKILL_ROOT = Path(__file__).resolve().parent.parent
TEMPLATE_DIR = SKILL_ROOT / "call_materials" / "application_forms"
TEMPLATE_JA = TEMPLATE_DIR / "2nd_form1_proposal.xlsx"
TEMPLATE_EN = TEMPLATE_DIR / "2nd_form1_proposal_en.xlsx"

# シート→XMLパス対応
# 日本語版・英語版とも同じ（テンプレートが共通設計のため）
SHEET_PATHS = {
    "sheet1": "xl/worksheets/sheet3.xml",  # 1枚目
    "sheet2": "xl/worksheets/sheet4.xml",  # 2枚目
    "sheet3": "xl/worksheets/sheet5.xml",  # 3枚目
    "sheet4": "xl/worksheets/sheet6.xml",  # 4枚目
}


def _xml_escape_text(s: str) -> str:
    """XMLテキストノード用エスケープ（属性値ではない）"""
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def _formula_inject_guard(value: str) -> str:
    """CSV インジェクション類似攻撃対策。
    先頭が = + - @ TAB CR LF の文字列はシングルクォートを前置して
    Excel が数式評価しないようにする。Excel 上では ' は隠れて元の文字列に見える。
    """
    if value and value[0] in ("=", "+", "-", "@", "\t", "\r", "\n"):
        return "'" + value
    return value


def _make_inline_str_cell(coord: str, style: str, text: str) -> str:
    """空セルを inlineStr セルに置換するXMLフラグメントを生成"""
    text = _formula_inject_guard(text)
    escaped = _xml_escape_text(text)
    style_attr = f' s="{style}"' if style else ""
    return (
        f'<c r="{coord}"{style_attr} t="inlineStr">'
        f'<is><t xml:space="preserve">{escaped}</t></is>'
        f'</c>'
    )


def _make_number_cell(coord: str, style: str, value) -> str:
    """空セルを数値セルに置換するXMLフラグメントを生成"""
    style_attr = f' s="{style}"' if style else ""
    return f'<c r="{coord}"{style_attr}><v>{value}</v></c>'


def _set_cell(xml: str, coord: str, value, *, as_number: bool = False) -> str:
    """sheet*.xml の指定セル（座標 coord, 例 "C8"）を value で書き換える。

    対象は **空セル** `<c r="..." s="..."/>` のみ。
    既値セル（<v> や <is> や <f> を含む）は触らない（混乱を避けるため）。
    """
    if value is None or value == "":
        return xml

    # 空セル（self-closing）パターン: <c r="C8" s="318"/> または <c r="C8"/>
    # 属性順は s が後ろ、または s 無し。s 値は数字のみ。
    pat_with_s = re.compile(
        r'<c\s+r="' + re.escape(coord) + r'"\s+s="(\d+)"\s*/>'
    )
    pat_no_s = re.compile(
        r'<c\s+r="' + re.escape(coord) + r'"\s*/>'
    )

    m = pat_with_s.search(xml)
    if m:
        style = m.group(1)
        if as_number:
            replacement = _make_number_cell(coord, style, value)
        else:
            replacement = _make_inline_str_cell(coord, style, str(value))
        return pat_with_s.sub(replacement, xml, count=1)

    m = pat_no_s.search(xml)
    if m:
        if as_number:
            replacement = _make_number_cell(coord, "", value)
        else:
            replacement = _make_inline_str_cell(coord, "", str(value))
        return pat_no_s.sub(replacement, xml, count=1)

    # 空セルが見つからない（既に値が入っている等）。スキップ。
    return xml


# ======== 1枚目セル割当 ========
# キー → (座標, 数値か否か)
SHEET1_CELLS_TEXT = {
    "erad_researcher_number": ("C8", True),
    "email": ("C10", False),
    "name_kana": ("D12", False),
    "name_kanji": ("D13", False),
    "birth_year": ("C15", True),
    "birth_month": ("F15", True),
    "birth_day": ("H15", True),
    "erad_institution_code": ("C16", True),
    "institution": ("C18", False),
    "department": ("C19", False),
    "position": ("C20", False),
    "institution_category": ("C21", False),
    "applicant_category": ("C23", False),
    "research_field": ("C27", False),
    "main_usecase": ("C29", False),
    "main_usecase_other": ("C31", False),
    "title": ("C35", False),
    "ai_usage_text": ("C42", False),
}

# 提出日（H6/J6/L6）— 数値
SHEET1_DATE = {
    "submit_year": ("H6", True),
    "submit_month": ("J6", True),
    "submit_day": ("L6", True),
}

# サブユースケース（"Y" を入れる8セル → ラベル）
SUBCASE_CELLS = [
    ("C32", "1.学習用データセット構築"),
    ("E32", "2.既存モデルの適応"),
    ("G32", "3.AIモデル開発"),
    ("I32", "4.既存モデル評価"),
    ("C33", "5.実験自動化・自律化"),
    ("E33", "6.シミュレーション・デジタルツイン"),
    ("G33", "7.発見・設計支援"),
    ("I33", "8.高度データ解析・モデリング"),
]

# AI活用度（"Y" を入れる10セル → ラベル）
AI_USAGE_CELLS = [
    ("C39", "研究でAIをまったく使っていない"),
    ("E39", "研究そのもの以外の業務でAIを使っている"),
    ("G39", "文献探索や要約にAIを使っている"),
    ("I39", "論文執筆や発表資料にAIを使っている"),
    ("K39", "仮説検討やアイデア出しにAIを使っている"),
    ("C40", "自作コード含めAIでデータ分析"),
    ("E40", "AI分析結果を論文・発表で発表"),
    ("G40", "APIで既存AIを研究プロセスに組み込み"),
    ("I40", "AIモデルの開発経験あり"),
    ("K40", "新しい基盤モデル開発の経験あり"),
]


def _label_match(label: str, selections: list) -> bool:
    """label と selections のいずれかが一致するか判定"""
    head = label.split("（")[0].split("(")[0].split("・")[0].strip()
    head_short = head[:6] if len(head) >= 6 else head
    for s in selections:
        s_norm = str(s).strip()
        if not s_norm:
            continue
        if s_norm in label or label in s_norm:
            return True
        if head and head in s_norm:
            return True
        if head_short and head_short in s_norm:
            return True
    return False


# ======== 2枚目セル割当 ========
SHEET2_TEXT = {
    "purpose": "D8",
    "method": "D9",
    "ai_rationale": "D10",
    "goals": "D11",
    "knowhow": "D12",
    "publication_policy": "D13",
}
SHEET2_ACHIEVEMENT_CELLS = ["D14", "D15", "D16", "D17", "D18"]


# ======== 3枚目セル割当 ========
S3_EQUIPMENT_ROWS = range(11, 31)
S3_CONSUMABLES_ROWS = range(11, 31)
S3_HONORARIUM_ROWS = range(40, 60)
S3_DOMESTIC_TRAVEL_ROWS = range(40, 60)
S3_FOREIGN_TRAVEL_ROWS = range(40, 60)
S3_OTHER_ROWS = range(69, 89)

S3_NECESSITY_CELLS = {
    "equipment_consumables_necessity": "C34",
    "honorarium_travel_necessity": "C63",
    "other_necessity": "C92",
}


# ======== 4枚目セル割当 ========
S4_API_ROWS = range(9, 19)
S4_COMPUTE_ROWS = range(22, 32)


def _write_sheet1(xml: str, data: dict) -> str:
    # 提出日
    for k, (coord, is_num) in SHEET1_DATE.items():
        v = data.get(k)
        if v not in (None, ""):
            xml = _set_cell(xml, coord, v, as_number=is_num)
    # 共通テキスト/数値
    for k, (coord, is_num) in SHEET1_CELLS_TEXT.items():
        v = data.get(k)
        if v not in (None, ""):
            xml = _set_cell(xml, coord, v, as_number=is_num)
    # 学生フラグ
    if data.get("is_student"):
        xml = _set_cell(xml, "L20", "Y")
    # サブユースケース
    selected_subs = data.get("subcases") or []
    for coord, label in SUBCASE_CELLS:
        if _label_match(label, selected_subs):
            xml = _set_cell(xml, coord, "Y")
    # AI活用度
    selected_levels = data.get("ai_usage_levels") or []
    for coord, label in AI_USAGE_CELLS:
        if _label_match(label, selected_levels):
            xml = _set_cell(xml, coord, "Y")
    return xml


def _write_sheet2(xml: str, data: dict) -> str:
    # テキスト
    for k, coord in SHEET2_TEXT.items():
        v = data.get(k)
        if v not in (None, ""):
            xml = _set_cell(xml, coord, v)
    # 業績（最大5件）
    achievements = data.get("achievements") or []
    if isinstance(achievements, str):
        achievements = [a for a in achievements.split("\n") if a.strip()]
    for i, item in enumerate(achievements[:5]):
        if item:
            xml = _set_cell(xml, SHEET2_ACHIEVEMENT_CELLS[i], item)
    return xml


def _write_sheet3(xml: str, data: dict) -> str:
    # 設備備品費 (rows 11..30): D=item, E=org, G=unit_price, I=qty (J=数式は触らず)
    rows = data.get("equipment_rows") or []
    for i, row in enumerate(rows[: len(S3_EQUIPMENT_ROWS)]):
        r = list(S3_EQUIPMENT_ROWS)[i]
        if row.get("item"):
            xml = _set_cell(xml, f"D{r}", row["item"])
        if row.get("org"):
            xml = _set_cell(xml, f"E{r}", row["org"])
        if row.get("unit_price") not in (None, ""):
            xml = _set_cell(xml, f"G{r}", row["unit_price"], as_number=True)
        if row.get("qty") not in (None, ""):
            xml = _set_cell(xml, f"I{r}", row["qty"], as_number=True)

    # 消耗品費 (rows 11..30): M=item, N=amount
    rows = data.get("consumables_rows") or []
    for i, row in enumerate(rows[: len(S3_CONSUMABLES_ROWS)]):
        r = list(S3_CONSUMABLES_ROWS)[i]
        if row.get("item"):
            xml = _set_cell(xml, f"M{r}", row["item"])
        if row.get("amount") not in (None, ""):
            xml = _set_cell(xml, f"N{r}", row["amount"], as_number=True)

    # 謝金 (rows 40..59): D=item, E=amount
    rows = data.get("honorarium_rows") or []
    for i, row in enumerate(rows[: len(S3_HONORARIUM_ROWS)]):
        r = list(S3_HONORARIUM_ROWS)[i]
        if row.get("item"):
            xml = _set_cell(xml, f"D{r}", row["item"])
        if row.get("amount") not in (None, ""):
            xml = _set_cell(xml, f"E{r}", row["amount"], as_number=True)

    # 国内旅費 (rows 40..59): H=item, J=amount
    rows = data.get("domestic_travel_rows") or []
    for i, row in enumerate(rows[: len(S3_DOMESTIC_TRAVEL_ROWS)]):
        r = list(S3_DOMESTIC_TRAVEL_ROWS)[i]
        if row.get("item"):
            xml = _set_cell(xml, f"H{r}", row["item"])
        if row.get("amount") not in (None, ""):
            xml = _set_cell(xml, f"J{r}", row["amount"], as_number=True)

    # 外国旅費 (rows 40..59): M=item, N=amount
    rows = data.get("foreign_travel_rows") or []
    for i, row in enumerate(rows[: len(S3_FOREIGN_TRAVEL_ROWS)]):
        r = list(S3_FOREIGN_TRAVEL_ROWS)[i]
        if row.get("item"):
            xml = _set_cell(xml, f"M{r}", row["item"])
        if row.get("amount") not in (None, ""):
            xml = _set_cell(xml, f"N{r}", row["amount"], as_number=True)

    # その他 (rows 69..88): D=item, E=amount
    rows = data.get("other_rows") or []
    for i, row in enumerate(rows[: len(S3_OTHER_ROWS)]):
        r = list(S3_OTHER_ROWS)[i]
        if row.get("item"):
            xml = _set_cell(xml, f"D{r}", row["item"])
        if row.get("amount") not in (None, ""):
            xml = _set_cell(xml, f"E{r}", row["amount"], as_number=True)

    # 必要性
    for k, coord in S3_NECESSITY_CELLS.items():
        v = data.get(k)
        if v not in (None, ""):
            xml = _set_cell(xml, coord, v)

    return xml


def _write_sheet4(xml: str, data: dict) -> str:
    # API費用 (rows 9..18): D=target, E=amount, F=basis
    rows = data.get("api_rows") or []
    for i, row in enumerate(rows[: len(S4_API_ROWS)]):
        r = list(S4_API_ROWS)[i]
        if row.get("target"):
            xml = _set_cell(xml, f"D{r}", row["target"])
        if row.get("amount") not in (None, ""):
            xml = _set_cell(xml, f"E{r}", row["amount"], as_number=True)
        if row.get("basis"):
            xml = _set_cell(xml, f"F{r}", row["basis"])
    # 計算資源費用 (rows 22..31): D=gpu, E=rationale, F=amount, G=basis
    rows = data.get("compute_rows") or []
    for i, row in enumerate(rows[: len(S4_COMPUTE_ROWS)]):
        r = list(S4_COMPUTE_ROWS)[i]
        if row.get("gpu"):
            xml = _set_cell(xml, f"D{r}", row["gpu"])
        if row.get("rationale"):
            xml = _set_cell(xml, f"E{r}", row["rationale"])
        if row.get("amount") not in (None, ""):
            xml = _set_cell(xml, f"F{r}", row["amount"], as_number=True)
        if row.get("basis"):
            xml = _set_cell(xml, f"G{r}", row["basis"])
    return xml


def _enable_full_calc_on_load(workbook_xml: str) -> str:
    """xl/workbook.xml の <calcPr> に fullCalcOnLoad="1" を付与する。

    テンプレ由来の <calcPr calcId="191028"/> のままだと、Excel で開いた瞬間に
    LEN(C35) などの数式セルが再計算されず、キャッシュ済みの <v>0</v> が
    そのまま表示されてしまう。fullCalcOnLoad="1" を付けると、Excel は開いた
    瞬間に全数式を強制再計算し、文字数カウントや経費合計が即座に正しい値に
    なる。
    """
    # 既に fullCalcOnLoad が付いていればそのまま
    if 'fullCalcOnLoad' in workbook_xml:
        return workbook_xml
    # <calcPr ... /> または <calcPr ...></calcPr>
    pat_self = re.compile(r'<calcPr([^/]*?)/>')
    m = pat_self.search(workbook_xml)
    if m:
        attrs = m.group(1)
        # 末尾空白を除いて fullCalcOnLoad を追加
        new_attrs = attrs.rstrip() + ' fullCalcOnLoad="1"'
        return workbook_xml[:m.start()] + f'<calcPr{new_attrs}/>' + workbook_xml[m.end():]
    pat_open = re.compile(r'<calcPr([^>]*?)>')
    m = pat_open.search(workbook_xml)
    if m:
        attrs = m.group(1)
        new_attrs = attrs.rstrip() + ' fullCalcOnLoad="1"'
        return workbook_xml[:m.start()] + f'<calcPr{new_attrs}>' + workbook_xml[m.end():]
    # calcPr 自体が無い場合は </workbook> 直前に挿入
    if '</workbook>' in workbook_xml:
        return workbook_xml.replace(
            '</workbook>',
            '<calcPr calcId="191028" fullCalcOnLoad="1"/></workbook>'
        )
    return workbook_xml


def fill(template_path: str, data: dict, output_path: str) -> list:
    """テンプレートを ZIP コピーし、対象シートだけ XML 文字列置換で書き換える。

    openpyxl は使わない。テンプレ構造を一切変更しないため、Excel の修復ダイアログ
    が出ない（出るならテンプレ自体が原因）。
    """
    warnings: list[str] = []
    src = Path(template_path)
    dst = Path(output_path)

    # テンプレを ZIP として読み、各 sheet*.xml を必要に応じて書き換える
    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(
        dst, "w", zipfile.ZIP_DEFLATED
    ) as zout:
        for item in zin.infolist():
            name = item.filename
            payload = zin.read(name)
            if name == SHEET_PATHS["sheet1"]:
                xml = payload.decode("utf-8")
                xml = _write_sheet1(xml, data)
                payload = xml.encode("utf-8")
            elif name == SHEET_PATHS["sheet2"]:
                xml = payload.decode("utf-8")
                xml = _write_sheet2(xml, data)
                payload = xml.encode("utf-8")
            elif name == SHEET_PATHS["sheet3"]:
                xml = payload.decode("utf-8")
                xml = _write_sheet3(xml, data)
                payload = xml.encode("utf-8")
            elif name == SHEET_PATHS["sheet4"]:
                xml = payload.decode("utf-8")
                xml = _write_sheet4(xml, data)
                payload = xml.encode("utf-8")
            elif name == "xl/workbook.xml":
                # 文字数カウント・経費合計などの数式を Excel 起動時に
                # 強制再計算させる
                xml = payload.decode("utf-8")
                xml = _enable_full_calc_on_load(xml)
                payload = xml.encode("utf-8")
            zout.writestr(item, payload)
    return warnings


def resolve_template(language: str, override: str | None) -> Path:
    if override:
        return Path(override)
    if language == "en":
        return TEMPLATE_EN
    return TEMPLATE_JA


def main():
    p = argparse.ArgumentParser(
        description="第2回 SPReAD 様式1 書き込み（openpyxl 非依存版）"
    )
    p.add_argument("--language", choices=["ja", "en"], default="ja")
    p.add_argument("--template", default=None)
    p.add_argument("--data", required=True)
    p.add_argument("--output", required=True)
    args = p.parse_args()

    with open(args.data, "r", encoding="utf-8") as f:
        data = json.load(f)
    if not data.get("language"):
        data["language"] = args.language

    template = resolve_template(data["language"], args.template)
    if not template.exists():
        raise FileNotFoundError(f"テンプレートが見つかりません: {template}")

    warnings = fill(str(template), data, args.output)
    print(f"Wrote: {args.output}")
    print(f"Template used: {template}")
    if warnings:
        print("\n警告:")
        for w in warnings:
            print(f"  - {w}")


if __name__ == "__main__":
    main()
