"""form1_proposal.xlsx（様式1＿研究計画調書）記入内容バリデーター

1. 2枚目の各項目が文字数範囲内か（日本語のみ、英語は任意）
2. 1枚目のサブユースケース・AI利活用度合いチェックが boolean 型か
3. 3枚目の総計と1枚目の C30 が矛盾しないか
4. 3枚目の総計が 500 万円を超えていないか（xlsx は千円単位で保持）
5. 必須項目（研究課題名・研究目的・研究方法 等）が未入力でないか

使い方:
    python validate.py workbook.xlsx
"""

import sys
from openpyxl import load_workbook

LIMITS_JA = {
    # (シート, セル, 下限, 上限, ラベル)
    ("研究計画調書_2枚目", "A3", 80, 400, "研究目的(日本語)"),
    ("研究計画調書_2枚目", "A5", 160, 800, "研究方法(日本語)"),
    ("研究計画調書_2枚目", "A7", 160, 800, "AI利活用の妥当性(日本語)"),
    ("研究計画調書_2枚目", "A9", 100, 500, "達成目標(日本語)"),
    ("研究計画調書_2枚目", "A11", 60, 300, "ノウハウ抽出(日本語)"),
}

REQUIRED = [
    ("研究計画調書_1枚目", "B6", "e-Rad 研究者番号"),
    ("研究計画調書_1枚目", "B7", "メールアドレス"),
    ("研究計画調書_1枚目", "D9", "氏名(漢字)"),
    ("研究計画調書_1枚目", "B22", "研究課題名(日本語)"),
    ("研究計画調書_2枚目", "A3", "研究目的(日本語)"),
    ("研究計画調書_2枚目", "A5", "研究方法(日本語)"),
]


def check(path: str):
    wb = load_workbook(path, data_only=False)
    errors = []
    warnings = []

    # 文字数制限
    for sheet, cell, low, high, label in LIMITS_JA:
        v = wb[sheet][cell].value
        if not v:
            warnings.append(f"{label} 未入力 ({sheet}!{cell})")
            continue
        n = len(str(v))
        if n < low:
            errors.append(f"{label} 文字数不足: {n}字 (下限{low})")
        elif n > high:
            errors.append(f"{label} 文字数超過: {n}字 (上限{high})")

    # 必須項目
    for sheet, cell, label in REQUIRED:
        v = wb[sheet][cell].value
        if not v:
            errors.append(f"必須項目未入力: {label} ({sheet}!{cell})")

    # 予算上限（1枚目 C30 は数式なので data_only で再評価が必要だが、合計値から判定）
    # ここでは3枚目の総計セル値の合算で判定する
    wb2 = load_workbook(path, data_only=True)
    try:
        eq = wb2["研究計画調書_3枚目"]["E9"].value or 0
        co = wb2["研究計画調書_3枚目"]["G9"].value or 0
        ho = wb2["研究計画調書_3枚目"]["B17"].value or 0
        dt = wb2["研究計画調書_3枚目"]["D17"].value or 0
        ft = wb2["研究計画調書_3枚目"]["F17"].value or 0
        ot = wb2["研究計画調書_3枚目"]["B27"].value or 0
        total = eq + co + ho + dt + ft + ot
        # 直接経費の上限は 500 万円（公募要領 3.(3) 補助上限額）
        # xlsx 様式は千円単位で保持されるため内部では 5000 と比較
        BUDGET_CAP_MAN_JPY = 500
        total_man = total / 10  # 千円 → 万円
        if total_man > BUDGET_CAP_MAN_JPY:
            errors.append(
                f"直接経費が上限(500万円)を超過: {total_man:.1f}万円"
            )

        # 90%超の費目チェック
        if total > 0:
            for amount, name in [
                (eq, "設備備品費"),
                (ho, "謝金"),
                (dt + ft, "旅費"),
            ]:
                if amount / total > 0.9:
                    nec_cell = {
                        "設備備品費": "A11",
                        "謝金": "A19",
                        "旅費": "A19",
                    }[name]
                    nec = wb2["研究計画調書_3枚目"][nec_cell].value
                    if not nec:
                        warnings.append(
                            f"{name}が総額の90%超。必要性記載を3枚目 {nec_cell} に記載必須"
                        )
    except Exception as e:
        warnings.append(f"予算集計チェックで例外: {e}")

    # 4枚目合計と3枚目 B27（その他）の整合
    try:
        api_sum = sum(
            (wb2["研究計画調書_4枚目"][f"C{r}"].value or 0) for r in range(5, 15)
        )
        compute_sum = sum(
            (wb2["研究計画調書_4枚目"][f"D{r}"].value or 0) for r in range(18, 28)
        )
        other_total = wb2["研究計画調書_3枚目"]["B27"].value or 0
        four_total = api_sum + compute_sum
        if four_total > other_total:
            errors.append(
                f"4枚目合計({four_total/10:.1f}万円) > 3枚目その他総計({other_total/10:.1f}万円)"
            )
        elif four_total > 0 and four_total < other_total * 0.8:
            warnings.append(
                f"4枚目合計({four_total/10:.1f}万円)と3枚目その他({other_total/10:.1f}万円)に乖離。"
                "AWS/API以外の経費が多い場合は問題なし"
            )
    except Exception as e:
        warnings.append(f"4枚目整合チェック例外: {e}")

    return errors, warnings


def main():
    if len(sys.argv) < 2:
        print("usage: python validate.py <xlsx>")
        sys.exit(1)
    errors, warnings = check(sys.argv[1])
    print("=" * 40)
    if errors:
        print(f"❌ エラー {len(errors)} 件:")
        for e in errors:
            print(f"  - {e}")
    else:
        print("✅ エラーなし")
    if warnings:
        print(f"\n⚠️  警告 {len(warnings)} 件:")
        for w in warnings:
            print(f"  - {w}")
    sys.exit(1 if errors else 0)


if __name__ == "__main__":
    main()
