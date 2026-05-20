# spread-proposal-assistant

**SPReAD（AI for Science 萌芽的挑戦研究創出事業）** に応募する研究者向けに、  
研究計画調書（様式1）の作成を **対話形式で伴走支援** する生成 AI スキルです。

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

| # | 機能 | 説明 |
|:-:|------|------|
| 1 | 📚 公募要領・様式の読み込み済み | 公募要領 PDF や様式テンプレートを同梱。AI が参照しながらヒアリングを進めます |
| 2 | 💬 対話形式で研究計画調書を作成 | 4枚のシート（基本情報・研究内容・経費・AWS費用算定）を順にヒアリングし xlsx に書き込み |
| 3 | 💰 AWS コスト試算と算定根拠の整形 | AWS Price List API から単価を取得し、算定根拠テンプレートに整形 |
| 4 | ✅ 文字数・予算上限の自動検証 | 提出前に文字数制限・必須項目・申請予算上限をチェック |

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
│   ├── fill_workbook.py        # openpyxl で様式1に書き込み
│   └── validate.py             # 文字数・必須項目・合計一致の検証
└── call_materials/             # SPReAD 公募要領・研究計画調書テンプレート等（※）
    ├── application_forms/
    ├── ai_interview/
    └── grant_terms/
```

**※ `call_materials/` 内の資料は、[SPReAD 公式サイト](https://www.mext.go.jp/aifors_spread/)から 2026年5月1日時点に取得した内容です。公募側で改訂があった場合は最新版をご確認ください。**


---

## 🚀 セットアップ

### 前提条件

- Python 3.10+
- 依存パッケージ:

```bash
pip install ijson openpyxl
```

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

## ⚠️ 免責事項

- 本スキルは申請書作成の **補助ツール** です。生成された内容・金額・文章の適切性の確認、および提出前の全項目の精読・検証は、すべて応募者ご自身の責任において実施してください
- 本スキルの利用により生じたいかなる損害（申請書の不受理、採択機会の逸失、経費計算の誤りに起因する会計処理上の不利益等）についても、配布者および関係者は一切の責任を負いません
- AWS 料金は公開価格ページを参照しますが、精度は担保できません。正確な見積もりには **[AWS Pricing Calculator](https://calculator.aws/#/estimate)** をご利用ください
- 公募要領の最新情報は **[文部科学省 SPReAD ページ](https://www.mext.go.jp/aifors_spread/)** をご確認ください

