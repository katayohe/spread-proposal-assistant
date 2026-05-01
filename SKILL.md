---
name: spread-proposal-assistant
description: AI for Science萌芽的挑戦研究創出事業（SPReAD）への応募者向けに、研究計画のヒアリングから「様式1＿研究計画調書.xlsx」の作成までを伴走支援するスキル。研究者はAWSを計算資源として利用する前提。公募要領・様式の内容を踏まえ、ヒアリングで情報を引き出し、文字数制限・費目区分・AWS利用料の算定根拠までを整えた上でExcelに埋め込む。ユーザーが「SPReAD」「AI for Science萌芽」「研究計画調書を作りたい」「様式1を埋めたい」と言及した場合に起動する。
license: Proprietary
---

# SPReAD 研究計画調書 作成支援スキル

## 使用タイミング

以下のいずれかに該当する場合に本スキルを使う。

- ユーザーが SPReAD（AI for Science 萌芽的挑戦研究創出事業）に応募したい
- ユーザーが「様式1＿研究計画調書.xlsx」を埋めたい／下書きしたい
- AWS を計算資源として使う前提で費用見積もりが必要

## スキルの全体方針

本スキルは**対話ヒアリング→ドラフト生成→xlsx 書き込み→検証**の4段構えで進める。ユーザーの時間を尊重し、1回に聞くのは1テーマに絞る。不明点は仮置きし、最後に一括で確認を取る。

### 前提・ファイル配置

本スキル本体は以下の構成で配置されており、`call_materials/` フォルダーはスキルディレクトリ直下（`SKILL.md` と同じ階層）にある。ファイル名は zip 配布互換のため ASCII に統一（元の日本語名は括弧内に併記）。

```
spread-proposal-assistant/
├── SKILL.md                               ← 本ファイル
├── references/                            ← ガイド・マスタ情報
├── scripts/                               ← xlsx 書込・費用計算・検証スクリプト
└── call_materials/                        (元: 公募要領等)
    ├── application_forms/                 (元: 公募要領・申請様式等)
    │   ├── spread_call_for_proposals.pdf  (元: AI for Science萌芽的挑戦研究創出事業(SPReAD)　公募要領.pdf)
    │   ├── form0_checklist.docx           (元: 様式0＿申請様式チェックリスト.docx)
    │   ├── form1_proposal.xlsx            (元: 様式1＿研究計画調書.xlsx) ← テンプレート本体
    │   ├── form2_review_consent.docx      (元: 様式2＿審査手法等同意確認書.docx)
    │   ├── form3_student_consent.docx     (元: 様式3＿学生応募の同意確認書.docx)
    │   └── form4_advisor_consent.docx     (元: 様式4＿指導教員等の同意確認書.docx)
    ├── ai_interview/                      (元: AIインタビュー)
    │   ├── ai_interview_overview.pdf
    │   └── ai_interview_guide_ja.pdf
    ├── grant_terms/                       (元: 交付要綱)
    │   ├── grant_terms.pdf
    │   └── grant_terms_forms.pdf
    └── grant_application_documents/       (元: 交付申請手続の際に必要な書類について)
```

**重要**：xlsx の**シート名**（`研究計画調書_1枚目` 等）は AWS Price List JSON と同じく Excel ファイル**内部**の名称なので、zip ファイル名の互換性とは無関係。そのまま日本語で保持している。

パスの扱い：

- `SKILL.md`・`references/`・`scripts/`・`call_materials/` はすべて**同じ親ディレクトリ**配下（以下これを `{SKILL_ROOT}` と呼ぶ）
- スクリプトは `__file__` から `{SKILL_ROOT}` を自動解決するので、絶対パスを渡さなくても動く（`fill_workbook.py` の `--template` は省略可）
- Claude 側で明示的に `{SKILL_ROOT}` を知る必要がある場合は、`scripts/` や `references/` のパスから逆引きする（例：本 `SKILL.md` を Read したときの絶対パスの親ディレクトリが `{SKILL_ROOT}`）
- 最終成果物は `{SKILL_ROOT}/../`（ユーザー作業フォルダ直下）に保存

その他の前提：

- xlsx の編集は既存の `anthropic-skills:xlsx` スキルのガイドラインに従う（フォーマット保持、既存テンプレートの規約優先、0フォーミュラエラー）。
- **`anthropic-skills:xlsx` の `scripts/recalc.py`（LibreOffice 経由）は呼ばない**。セル内の部分赤字（リッチテキスト）・条件付き書式が劣化するため。代わりに `fill_workbook.py` が `openpyxl.load_workbook(..., rich_text=True)` でリッチテキストを保持、費目小計を数値で直接書き込む。
- AWS 費用は `scripts/fetch_aws_price.py`（AWS 公開料金 JSON を都度取得）を利用。MCP は使わない。

### 出力物

1. 埋め込み済みの `form1_proposal_{e-Rad所属機関コード}_{ローマ字氏名}.xlsx`（最終提出時は文科省指定のファイル名規則 `様式１＿研究計画調書＿e-Rad所属機関コード＿ローマ字氏名.xlsx` にリネームしてからユーザーに渡す）
2. ヒアリング内容を要約した `hearing_notes.md`（内部メモ）
3. AWS 費用算定の内訳 `aws_cost_breakdown.md`（研究計画調書_4枚目の算定根拠ソース）

全て `{WORK_DIR}` 配下に保存し、最終ファイルは `computer://` リンクで提示する。

## 推奨ワークフロー

### Step 0: 準備

1. `{SKILL_ROOT}/references/overview.md` を Read し、調書の全シート構成・記入項目・文字数制約を把握する。
2. 公募要領PDFの重要箇所（補助上限・対象経費・計算資源・AI インタビュー要件）を `pdftotext -layout "{SKILL_ROOT}/call_materials/application_forms/spread_call_for_proposals.pdf" -` で抽出してコンテキストに入れる。

### Step 1: キックオフヒアリング（AskUserQuestion）

最初に以下4点をまとめて AskUserQuestion で聞く。テキスト回答で来るものは一旦受け取り、後で整形する。

1. 研究課題名（仮で可）
2. 研究領域（`references/lists.md` の11領域から選択）
3. メインユースケース分類（8+「その他」から選択）
4. 身分区分（大学教員／博士課程学生等、`references/lists.md` から）

### Step 2: 基本情報の確認

`references/sheet1_basic.md` を Read し、記入項目を順にヒアリングする。e-Rad 研究者番号、所属機関、メールアドレス、生年月日等は**必須だが推測禁止**なのでユーザーに直接聞く。不明な項目は空欄のまま進めてよい（ユーザーが後日手入力）。

### Step 3: 研究内容ヒアリング（2枚目）

`references/sheet2_research.md` のプロンプト集に従い、以下を順番にヒアリングする。各項目は**文字数下限・上限が厳密**なので、ヒアリング後にドラフトを生成し、LEN() で確認してから書き込む。

- 研究目的（日本語80〜400字）
- 研究方法（日本語160〜800字）
- AI 利活用の妥当性・実現可能性（日本語160〜800字）
- 達成目標（日本語100〜500字、3ヶ月／6ヶ月の中間・最終目標を区別）
- ノウハウ抽出・共有計画（日本語60〜300字）
- 研究業績等（最大5件、改行区切り）

ヒアリング時は open question→深掘り→要約確認の順で行う。1項目につき2〜3往復で収束させる。

### Step 4: 研究経費の設計（3枚目）

`references/sheet3_budget.md` を Read。直接経費 500 万円以内の制約下で、費目（設備備品費／消耗品費／謝金／国内旅費／外国旅費／その他）の配分をヒアリングする。**人件費は対象外**。本事業では計算資源（AWS）は主に「その他」に計上し、詳細を4枚目に記載する。

**単位ルール**：ユーザーとの会話・ドラフトは **万円** で統一（「500 万円」「300 万円」）。xlsx 様式の金額欄は「千円」単位で固定されているので、書き込む時のみ `千円 = 万円 × 10` に換算し、千円未満は切り捨て。

3枚目は行追加が許可されている。品名・数量・単価・金額（xlsx 上は千円単位）を埋める。

### Step 5: API・計算資源の算定（4枚目）★ AWS MCP 活用

`references/sheet4_aws_cost.md` を Read。このステップが本スキルの中核。

#### 5-1. AWS 利用想定ヒアリング

以下を AskUserQuestion で確認：

- 主な用途（LLM ファインチューニング／推論／データ前処理／ベクター検索 等）
- 想定 GPU インスタンス（A100／H100／L40S／推論用 等、未定なら「要選定」）
- 見込み稼働時間（時間/月 × 月数）
- データ容量（S3 保管量・転送量）
- 使用予定 AWS サービス（EC2, SageMaker, Bedrock, S3, EFS, ECR 等）

#### 5-2. AWS 公開料金 API から**必要なサービスだけ**取得

**MCP サーバーは使わない**（本スキル環境で利用不可のため）。代わりに AWS Price List Bulk API（認証不要、公開エンドポイント）を利用する。

アクセスフロー：

1. `https://pricing.us-east-1.amazonaws.com/offers/v1.0/aws/index.json` で**サービス一覧のメタデータ**（~100KB）を1回取得し、対象サービス（AmazonEC2, AmazonBedrock, AmazonS3 等）の `currentRegionIndexUrl` を確認
2. そのサービスの region_index から、対象リージョン（例：`ap-northeast-1`）の `currentVersionUrl` を解決
3. regional 本体 JSON（数百MB）は**ダウンロードせず HTTP ストリームで読む**。`ijson` で products を順次パースし、**対象 SKU を見つけた時点でストリームを閉じる**（全ダウンロードしない）
4. 取得した単価は `/tmp/aws_price_cache/results/` に SKU 単位で1週間キャッシュ。2回目以降は本体に触らない

実装は `scripts/fetch_aws_price.py`。Python API として以下を提供：

```python
from scripts.fetch_aws_price import (
    get_ec2_ondemand_usd_hour,        # 例: p4de.24xlarge, ap-northeast-1
    get_bedrock_token_usd_per_mtok,   # 入出力トークン $/1M tokens
    get_s3_standard_usd_gb_month,
)
```

CLI でも使える：

```bash
python3 spread-proposal-assistant/scripts/fetch_aws_price.py ec2 p4de.24xlarge ap-northeast-1
python3 spread-proposal-assistant/scripts/fetch_aws_price.py bedrock anthropic.claude-sonnet-4-20250514 us-east-1 --io output
python3 spread-proposal-assistant/scripts/fetch_aws_price.py s3 standard ap-northeast-1
```

ヒアリングで使うサービスが EC2 だけなら EC2 の metadata だけ、Bedrock も使うなら Bedrock もという風に、**使うサービスに絞って都度取得**する方針。取得できない場合（新しい SKU 等）は AWS 公式料金ページ（https://aws.amazon.com/ec2/pricing/ 等）をフォールバックとし、**参照日と情報源を算定根拠欄に明記**する。

ベストプラクティス（Reserved/Savings Plans/Spot の適用可否等）は AWS 公式ドキュメント URL を直接参照（例：https://docs.aws.amazon.com/whitepapers/latest/cost-optimization-laying-the-foundation/aws-savings-plans-reserved-instances-and-spot-instances.html）。

#### 5-3. 算定根拠を表に整形

`scripts/compute_aws_cost.py` を呼び出し、ヒアリング値＋`fetch_aws_price.py` から取得した単価を掛け合わせて月額／総額（千円単位）を算出。同スクリプトが `aws_cost_breakdown.md` を生成し、以下の欄に書き込むテキストを準備する：

- API 費用テーブル（処理対象／金額／算定根拠）- 該当時のみ（Bedrock 等）
- 計算資源費用テーブル（GPU種類／選定理由／金額／算定根拠）

選定理由では「学習データ量・モデル規模から A100 80GB 以上が必要」等、研究要件と結びつけて書く。算定根拠では「$X/時 × Y時間/月 × Zヶ月 × 為替150円/USD = ◯千円」のように式を明示する。**為替レートはユーザーに確認（所属機関の会計基準に従う）。**

### Step 6: xlsx への書き込み

`scripts/fill_workbook.py` を使ってテンプレをコピー→openpyxl で書き込み。

**重要：LibreOffice による再計算（`anthropic-skills:xlsx` の `scripts/recalc.py`）は呼ばない**。テンプレの条件付き書式や赤字の注意書き（「←応募者はここは入力しないこと」等）の**フォント色が失われる**事故が起きるため。代わりに `fill_workbook.py` 側で 3 枚目の各費目小計を**数値で直接書き込む**ようにしている（1 枚目の C30 等は元の数式を維持し、Excel で開いたときに再計算される）。

呼び出し例：

```bash
# --template は省略可（スクリプトが同梱テンプレートを自動参照）
python3 "{SKILL_ROOT}/scripts/fill_workbook.py" \
  --data /tmp/payload.json \
  --output "{SKILL_ROOT}/../様式１＿研究計画調書＿{機関コード}＿{ローマ字氏名}.xlsx"
```

重要：
- 既存セルの数式・書式は絶対に変更しない（`references/overview.md` の数式一覧を参照）
- `研究計画調書_1枚目!C30` など合計は数式のまま残す
- 未入力セル（薄オレンジ）は入力必須ではない項目もあるので、ユーザーが空欄を選んだ項目はそのまま
- 文字数カウント列（L列/M列）は既存数式を保持

### Step 7: 検証

最後に以下を自動チェック：

1. `scripts/validate.py` 実行 → 各文字数制限の上下限チェック、必須項目（課題名・研究目的 等）の未入力検出、3枚目合計と1枚目`C30`の一致確認
2. フォーミュラエラー 0 件
3. ユーザーに最終レビュー依頼（AI インタビュー完了スクショの貼付、様式0チェックリストの記入は手動で必要であることも明示）

## 重要な制約・注意

- **調書は PDF 化せず xlsx のまま提出**（公募要領に明記）
- **ファイル名は `様式１＿研究計画調書＿e-Rad所属機関コード＿ローマ字氏名.xlsx`**
- **AI インタビューのメールアドレスと調書のメールアドレスは完全一致必須**（不一致は不受理）
- **直接経費 500 万円以下**、間接経費は直接経費の 30%（機関配分）
- 金額はユーザー会話では **万円**、xlsx 書込は **千円**（`千円 = 万円 × 10`、千円未満切り捨て）
- **人件費は対象外**
- **設備備品費・旅費・謝金のいずれかが90%超の場合は必要性を3枚目に記載**
- 色を付した図や文字もそのまま審査に付される
- 2枚目の図は最大1枚
- 様式0〜4の他書類も必要だが、本スキルは様式1のみ扱う（他は手動案内）

## 参照ファイル

- `references/overview.md` — 調書全シートの構造・セル一覧・数式
- `references/lists.md` — 研究領域・ユースケース・区分のマスタ
- `references/sheet1_basic.md` — 1枚目の記入項目とヒアリング順
- `references/sheet2_research.md` — 2枚目のプロンプト集・文字数制限
- `references/sheet3_budget.md` — 3枚目の費目設計ガイド・府省共通経費取扱区分
- `references/sheet4_aws_cost.md` — 4枚目 AWS 費用算定プロトコル
- `references/aws_gpu_instances.md` — SPReAD向け代表的GPUインスタンスと選定指針
- `scripts/compute_aws_cost.py` — ヒアリング値から費用表を生成
- `scripts/fill_workbook.py` — openpyxl で様式1に書き込み
- `scripts/validate.py` — 文字数・必須項目・合計一致の検証
