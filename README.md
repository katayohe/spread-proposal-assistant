# spread-proposal-assistant

**SPReAD（AI for Science 萌芽的挑戦研究創出事業）第2回公募** に応募する研究者向けに、研究計画調書（様式1）の作成を **対話形式で伴走支援** する生成 AI スキルです。

> [!NOTE]
> **第2回公募（令和8年6月2日〜7月3日）の様式・要領変更にあわせて、スキル全体を新規に作り直しています。**
> 第2回からの主な変更点：AIインタビュー（旧5枚目）の廃止、所属機関の区分／応募者属性の区分の分離、業績欄の構造変更、3枚目の行追加禁止、成果の公開方針（任意）の新設、英語版様式の追加 など。これらすべてに対応した記入項目・選択肢マスタ・xlsx 書き込みロジックを採用しています。

> [!IMPORTANT]
> 本スキルは筆者が個人開発した OSS ツールです。
> 生成された内容は参考情報の位置づけです。提出前に必ずご自身で検証してください。
> e-Rad 研究者番号や未発表の研究アイデアを入力する前に、所属機関の AI 利用ポリシーをご確認ください。

---

## 🔗 関連リンク

- 📄 **[SPReAD 公式ページ（文部科学省）](https://www.mext.go.jp/aifors_spread/)** — 公募要領・様式のダウンロード、最新情報
- 📝 **[Zenn 解説記事](https://zenn.dev/aws_japan/articles/35e76df16e02b4)** — 本スキルの背景・使い方・セットアップ手順の詳細

---

## ✨ できること

| # | 機能                              | 説明                                                                                                 |
| :-: | --------------------------------- | ---------------------------------------------------------------------------------------------------- |
| 1 | 📚 公募要領・様式の読み込み済み   | 公募要領 PDF や様式テンプレートを同梱。AI が参照しながらヒアリングを進めます                         |
| 2 | 💬 対話形式で研究計画調書を作成   | 4枚のシート（基本情報・研究内容・経費・AWS費用算定）を順にヒアリングし xlsx に書き込み               |
| 3 | 💰 AWS コスト試算と算定根拠の整形 | AWS Price List API から単価を取得し、算定根拠テンプレートに整形                                      |
| 4 | ✅ 文字数・予算上限の自動検証     | 文字数・必須項目・予算上下限・桁数（e-Rad番号/機関コード）・選択肢一致・ファイル名規則を自動チェック |

---

## 📁 ディレクトリ構成

```
spread-proposal-assistant/
├── SKILL.md                    # メインの指示書
├── README.md                   # 本ファイル
├── references/                 # シート別の記入ガイド・選択肢マスタ
│   ├── overview.md
│   ├── lists.md
│   ├── sheet1_basic.md
│   ├── sheet2_research.md
│   ├── sheet3_budget.md
│   ├── sheet4_aws_cost.md
│   └── aws_gpu_instances.md
├── scripts/                    # AWS 料金試算・xlsx 書き込み・検証スクリプト
│   ├── fetch_aws_price.py      # AWS Price List Bulk API から単価取得
│   ├── compute_aws_cost.py     # 費用算定ヘルパー
│   ├── fill_workbook.py        # 様式1への書き込み（XML文字列置換、--language ja|en）
│   └── validate.py             # 文字数・必須項目・合計一致の検証
└── call_materials/             # SPReAD 第2回公募要領・研究計画調書テンプレート等（※）
    ├── application_forms/      # 第2回公募の様式 (2nd_ prefix)
    └── e_rad_guide/            # e-Rad 操作・入力ガイド (第2回別紙)
```

**※ `call_materials/` 内の資料は、[SPReAD 公式サイト](https://www.mext.go.jp/aifors_spread/)から第2回公募開始時点（2026年6月8日）に取得した内容です。公募側で改訂があった場合は最新版をご確認ください。**

---

## 🚀 セットアップ

### 前提条件

- Python 3.10+
- 依存パッケージ（[`requirements.txt`](./requirements.txt) 参照）:

```bash
pip install -r requirements.txt
```

> 一部の OS（Debian/Ubuntu の system Python など）では PEP 668 により pip install が拒否されます。その場合は `--break-system-packages` を付けるか、仮想環境を使ってください：
>
> ```bash
> python3 -m venv .venv && source .venv/bin/activate && pip install -r requirements.txt
> ```

Skills が利用可能な環境（Amazon Quick Desktop, Claude Code, Codex, Kiroなど）で動作します

### Amazon Quick Desktop でセットアップする例

- 📝 **[Zenn 記事 - Amazon Quick Desktop をインストールして SPReAD 研究計画調書作成支援 Skill を起動する](https://zenn.dev/aws_japan/articles/f035e40e01444e)** — Quick Desktop のインストールからスキルの利用開始の設定までを行う

---

## 💬 使い方

スキルを導入した環境で、以下のように話しかけると起動します。

> **SPReAD の研究計画調書を作りたい**

> **AI for Science 萌芽の様式1を埋めたい**

あとは AI エージェントのヒアリングに答えていくだけで、研究計画調書の作成が進みます。

---

## ✅ 提出前の形式チェック

本スキルにも独自の検証スクリプト（[`scripts/validate.py`](./scripts/validate.py)）が同梱されており、文字数制限・必須項目・直接経費の上下限・e-Rad 番号と機関コードの桁数・研究領域／メインユースケース分類の選択肢一致・ファイル名規則などを自動でチェックします。

ただし、**最終的な提出前には文部科学省が配布している公式の形式チェックツールも併用することを強く推奨します**。公式ツールは様式1（Excel）と様式0/2/3/4（PDF）に対応しており、文部科学省側の判定基準に最も近い検証が行えます。

- 📥 **公式形式チェックツールの配布元（Box フォルダ）**：[https://mext.ent.box.com/s/qf8vbuj3pso1hj9mwp1vs6rx2hashpuc/folder/385379226260](https://mext.ent.box.com/s/qf8vbuj3pso1hj9mwp1vs6rx2hashpuc/folder/385379226260)
- [!IMPORTANT]

> **本スキルは公式形式チェックツールを利用していません。** また、ライセンス上の理由から公式ツールの **再配布も行っていません**。
> 公式ツールをご利用の際は、上記 Box フォルダから直接ダウンロードしてください。参考となる、ダウンロード方法・実行手順は [SKILL.md の Step 7](./SKILL.md) に詳細を記載しています

> [!NOTE]
> 本スキルが生成した直後の xlsx ファイル は、文字数カウントや経費合計の数式が**まだ計算されていない状態**で保存されます（Excel が `fullCalcOnLoad` で開いた瞬間に再計算する設計）。
> このまま公式の形式チェックツールにかけると数式キャッシュが空のため誤判定される可能性があります。**提出前および公式ツール実行前には、必ず一度 Excel で開いて「名前を付けて保存／コピーを保存」で保存し直してください。**

---

## ⚠️ 免責事項

- 本スキルは申請書作成の **補助ツール** です。生成された内容・金額・文章の適切性の確認、および提出前の全項目の精読・検証は、すべて応募者ご自身の責任において実施してください
- 本スキルの利用により生じたいかなる損害（申請書の不受理、採択機会の逸失、経費計算の誤りに起因する会計処理上の不利益等）についても、配布者および関係者は一切の責任を負いません
- AWS 料金は公開価格ページを参照しますが、精度は担保できません。正確な見積もりには **[AWS Pricing Calculator](https://calculator.aws/#/estimate)** をご利用ください
- 公募要領の最新情報は **[文部科学省 SPReAD ページ](https://www.mext.go.jp/aifors_spread/)** をご確認ください
