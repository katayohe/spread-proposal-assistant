# 4枚目（API・計算資源費用 算定根拠）ガイド — AWS 前提版

本スキルは研究者が **AWS** を計算資源として使う前提。価格は **AWS Price List Bulk API（公開・認証不要）** から都度取得する。ヘルパーは `scripts/fetch_aws_price.py`。

**単位ルール**：
- ヒアリング・ドラフト・メモは **万円** で統一
- xlsx 様式（4枚目の「金額(千円)」欄）への書き込み時のみ **千円** に換算（`千円 = 万円 × 10`）
- 千円未満切り捨て（様式の注記に従う）

## 4枚目の構造

### API費用テーブル（行5〜14、最大10行）

| 列 | 内容 | 本スキルでの必須記載 |
|---|---|---|
| A | 通し番号（1〜10、既に入っている） | — |
| B | 処理対象 | **AWSサービス名**（例：`Amazon Bedrock (Claude Sonnet 4)`、`Amazon EC2 推論エンドポイント`） |
| C | 金額（千円） | 数値 |
| D | 算定根拠 | **[API種類]（日本語）／[サービス]／[リージョン]／[単価]／[想定量]／[稼働時間]／[計算式]／[備考]** をテンプレ形式で記述。B列がサービス名でも、D列先頭にはユーザーが一目で分かる**日本語のAPI種類**を置く |

### 計算資源費用テーブル（行18〜27、最大10行）

| 列 | 内容 | 本スキルでの必須記載 |
|---|---|---|
| A | 通し番号（1〜10） | — |
| B | GPU種類 | **AWSサービス名＋インスタンス名＋GPU型番** を全て書く（例：`Amazon EC2 p4de.24xlarge (NVIDIA A100 80GB × 8)`） |
| C | 当該GPUを選定した理由 | 研究要件（モデル規模・VRAM・代替案比較）と紐づけて記載 |
| D | 金額（千円） | 数値 |
| E | 算定根拠 | **計算式 ＋ 想定量 ＋ 単価情報源** を書く |

## AWS 費用算定プロトコル

### Step 1. 用途の整理

ユーザーから以下を引き出す：

1. **用途分類**：学習（training）／推論（inference）／データ前処理／埋め込み生成／ベクター検索 等
2. **モデル規模**：〜1B, 1〜10B, 10〜70B, 70B+ パラメータ
3. **データ量**：学習データ全体サイズ（GB/TB）、想定エポック数
4. **稼働想定**：連続稼働（例：1週間×8h）か、断続稼働か、研究期間全体での総時間
5. **マネージドサービス利用**：SageMaker、Bedrock（Claude/Titan等）、EC2 生インスタンス

### Step 2. AWS 公開料金 API で価格取得

**MCP は使わない**。アクセスフローは以下：

1. `https://pricing.us-east-1.amazonaws.com/offers/v1.0/aws/index.json`（~100KB）を取得
2. 対象サービスの `currentRegionIndexUrl` → 対象リージョンの `currentVersionUrl` を解決
3. regional 本体 JSON は**ストリーム読み**（`ijson`）、対象 SKU で早期終了

`scripts/fetch_aws_price.py` がこれを行う。典型呼び出し：

```bash
python3 scripts/fetch_aws_price.py ec2 p4de.24xlarge ap-northeast-1
python3 scripts/fetch_aws_price.py bedrock anthropic.claude-sonnet-4-20250514 us-east-1 --io input
python3 scripts/fetch_aws_price.py bedrock anthropic.claude-sonnet-4-20250514 us-east-1 --io output
python3 scripts/fetch_aws_price.py s3 standard ap-northeast-1
```

EC2 GPU インスタンスの典型例：`p4d.24xlarge`（A100 40GB×8）、`p4de.24xlarge`（A100 80GB×8）、`p5.48xlarge`（H100×8）、`p5e.48xlarge`（H200×8）、`g6e.12xlarge`（L40S×4）、`g5.xlarge`（A10G×1）

リージョンは原則 `ap-northeast-1`（東京）、必要に応じて `us-east-1`/`us-west-2` も比較。オンデマンドが基本。Spot 割引（最大70〜90%OFF）や Savings Plans（1年〜、3年〜）は研究期間6ヶ月には**不向き**（Savings Plans のコミット期間が合わない）。

付帯コスト（S3 Standard 容量・データ転送・EBS gp3・CloudWatch 等）は `fetch_aws_price.py` の該当サービス呼び出しで取得するか、AWS 公式料金ページを参照。

### Step 2.5. ベストプラクティス参照

AWS 公式ドキュメントを直接参照：

- Cost Optimization (Well-Architected)：https://docs.aws.amazon.com/wellarchitected/latest/cost-optimization-pillar/
- Spot を学習で使う際の注意：https://docs.aws.amazon.com/AWSEC2/latest/UserGuide/using-spot-instances.html
- SageMaker vs EC2：https://docs.aws.amazon.com/sagemaker/latest/dg/whatis.html

### Step 3. 算定式を組む

```
月額(円) = インスタンス時間単価(USD) × 稼働時間(h/月) × 為替(円/USD) × インスタンス数
総額(円) = 月額 × 月数
```

ヒアリング・メモでは **万円** 表示（`総額 ÷ 10000`）、xlsx書き込み時のみ**千円** 表示（`総額 ÷ 1000`、千円未満切り捨て）。

為替レートは**所属機関の会計基準**に従う（ユーザーに確認）。一般的には1USD=150円や、会計年度ごとの予算用レートを使う。

### Step 3.5. 算定根拠テンプレ

**算定根拠の書き方テンプレ**（xlsx の E列に書き込む文章、**万円単位で表示**）：

```
[用途] LLMファインチューニング学習（学習データ10GB × 3エポック）
[サービス] Amazon EC2 p4de.24xlarge（NVIDIA A100 80GB × 8）
[リージョン] ap-northeast-1（東京）、オンデマンド
[単価] $40.96/時（2026年4月時点、AWS Price List Bulk API より `scripts/fetch_aws_price.py ec2 p4de.24xlarge ap-northeast-1` で取得）
[稼働時間] 80時間/月 × 6ヶ月 = 480時間
[再実行率] 20% バッファ
[計算] $40.96/時 × 480時間 × 1.20 × 150円/USD = 3,538,944円 ≒ 353.9 万円
[備考] 複数回のハイパーパラメータ探索を含む
```

（xlsx 4枚目「金額(千円)」欄には `3,538` と入れる）

単価の情報源は `scripts/fetch_aws_price.py`（AWS Price List Bulk API）で都度取得し、取得日を `[単価]` 行に明記する。取得できない SKU は AWS 公式料金ページ（https://aws.amazon.com/ec2/pricing/ 等）の参照日を明記。**Pricing Calculator の共有URLは不要**。

**選定理由の書き方例（C列）**：

```
70Bパラメータ級LLMのフルファインチューニング（学習データ10GB）には、
80GB VRAM×8 GPUが必要。A100 80GB搭載の p4de.24xlarge を選定。
H100 (p5) も検討したが、本研究では FP16 学習で充分であり、
価格優位な A100 で要件を満たすため。
```

### Step 4. API費用の特殊ケース（Bedrock 等）

Bedrock を使う場合は「その他（API費用）」として4枚目上段に記載：

- 処理対象（B列）：**AWSサービス名**（例：`Amazon Bedrock (Claude Sonnet 4)`）。具体的な用途・処理対象は算定根拠（D列）の `[備考]` に書く
- 金額（C列、千円）：月あたりトークン数 × 1Mトークン単価 × 月数 × 為替（×10で千円換算）
- 算定根拠（D列）：下記テンプレに沿って記述。審査員が一目で何のAPIか分かるよう **API種類を日本語**で、さらに**リージョン・単価・稼働時間（またはトークン量）**を必ず含める

```
[API種類] 大規模言語モデルAPI（文献要約・情報抽出用）
[サービス] Amazon Bedrock (Claude Sonnet 4)
[リージョン] us-east-1
[単価] 入力 $3.00/1M tokens、出力 $15.00/1M tokens（2026年4月時点、AWS Price List Bulk API取得）
[想定量] 入力 500,000 tok/月、出力 100,000 tok/月 × 6ヶ月
[計算] (500K×$3 + 100K×$15)/1M × 6ヶ月 × 150円/USD × 1.2(再実行) = 3,240円 ≒ 0.3 万円
[備考] PubMed論文のメタデータ抽出とサマリ生成
```

時間ベースの API（推論エンドポイント等）の場合は `[稼働時間] X時間/月 × Yヶ月` の行も追加する。

**データ取得・照会・連携のための API**（例：PubMed API, GenBank API の有償プラン）も同じ書式で記載。AWSサービス以外の場合はサービス提供元名を先頭に書く（例：`[サービス] NCBI E-utilities API (有償プラン)`）。

**データ取得・照会・連携のための API**（例：PubMed API, GenBank API の有償プラン）も同じ書式で記載。

### Step 5. 合計と3枚目との整合

- 4枚目の API費用合計＋計算資源費用合計 ＝ 3枚目「その他」欄の該当金額
- 3枚目「その他」欄の「事項」には「AWS計算資源費用（詳細は4枚目）」「Bedrock API利用料（詳細は4枚目）」等と記載。
- 3枚目 B27（その他総計）と4枚目合計は**厳密に一致**させる。

## よくある設計パターン

### パターン A：大規模ファインチューニング中心（500 万円上限使い切り）

- p4d.24xlarge × 80h/月 × 6ヶ月 ≈ 230 万円
- 推論用 g5.xlarge × 200h/月 × 6ヶ月 ≈ 15 万円
- Bedrock Claude ≈ 10 万円
- S3/EBS/Transfer ≈ 20 万円
- バッファ（リトライ・実験失敗）：20%織込み
- 合計 計算資源：450 万円（その他に計上）、API：10 万円（その他）
- 残り 40 万円で消耗品・旅費

### パターン B：推論中心・データ解析

- g5.12xlarge × 100h/月 × 6ヶ月 ≈ 80 万円
- Bedrock Claude（文献要約） ≈ 50 万円
- SageMaker 処理ジョブ ≈ 20 万円
- S3 ≈ 5 万円
- 合計：155 万円程度

### パターン C：Bedrock 主体（API型）

- Bedrock Claude Sonnet 大量推論：300 万円
- Lambda/Step Functions：10 万円
- S3：5 万円
- OpenSearch (ベクター検索)：50 万円
- 合計：365 万円程度

## 注意

- 物理的に GPU サーバを購入する場合は**設備備品費（3枚目）** に計上し、4枚目「計算資源費用」には含めない（公募要領の留意事項タブに明記あり）。
- **為替変動リスク**は選定理由か算定根拠に一言触れておくと親切（「為替150円/USDで算定。変動時は費目間流用で吸収」等）。
- **xlsx 様式**の金額欄は**千円単位・千円未満切り捨て**（様式固定）。**ヒアリングや内部メモは万円単位**で扱い、書き込む時だけ `×10` で千円に換算する。
