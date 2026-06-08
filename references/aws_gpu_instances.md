# AWS GPUインスタンス 選定ガイド（SPReAD 向け）

研究用途別の代表的なインスタンスと選定指針。**価格は変動するので必ず `scripts/fetch_aws_price.py`（AWS 公開料金 API 経由）で最新を確認すること**（この表はヒアリングの起点用の目安）。

## 学習用（Training）

| インスタンス | GPU | vCPU | メモリ | 用途目安 |
|---|---|---|---|---|
| `p5.48xlarge` | H100 80GB × 8 | 192 | 2 TB | 大規模LLMフルFT、マルチノード分散学習 |
| `p5e.48xlarge` | H200 141GB × 8 | 192 | 2 TB | 超大規模LLM、70B+ |
| `p4de.24xlarge` | A100 80GB × 8 | 96 | 1.1 TB | 10〜70B LLM FT、マルチモーダル |
| `p4d.24xlarge` | A100 40GB × 8 | 96 | 1.1 TB | 〜10B LLM FT、CV学習 |
| `g6e.12xlarge` | L40S 48GB × 4 | 48 | 384 GB | 中規模学習、コスト重視 |
| `g6.12xlarge` | L4 24GB × 4 | 48 | 192 GB | 軽量FT、LoRA |

## 推論用（Inference）

| インスタンス | GPU | 用途目安 |
|---|---|---|
| `g5.xlarge` | A10G 24GB × 1 | 小規模モデル推論、〜7B |
| `g5.12xlarge` | A10G × 4 | 30B 級推論 |
| `g6.xlarge` | L4 × 1 | 小規模推論、コスト重視 |
| `inf2.xlarge` | Inferentia2 × 1 | 専用チップ、大規模バッチ推論 |
| `p4d.24xlarge` | A100 × 8 | 超大規模モデル推論、研究用途 |

## マネージド・代替

### Amazon Bedrock
- Claude / Titan / Llama / Mistral 等を**トークン課金**で利用
- 学習が不要、推論だけでよい場合は EC2 GPU より圧倒的に安く、運用負担ゼロ
- ベースモデル学習（事前学習）は不可、ファインチューニングは一部モデルで対応

### Amazon SageMaker
- Training Jobs / Processing Jobs / Inference Endpoints
- EC2 より 10〜40% 高い代わりに運用（監視・スケーリング）がラク
- 学生・小規模チームにおすすめ

## 付帯コスト（忘れがち）

- **S3 Standard**：約 $0.025/GB/月（ap-northeast-1）
- **EBS gp3**：約 $0.096/GB/月
- **Data Transfer Out**：最初の100GB/月無料、以降約 $0.114/GB（大量にダウンロードすると重い）
- **NAT Gateway**：時間課金＋データ処理料金
- **CloudWatch**：ログ・メトリクス蓄積で発生

研究計画では **総計の 10〜20%程度を付帯コストとしてバッファ** することを推奨。

## 選定のフロー

```
Q1. モデルを学習する？
  Yes → Q2へ
  No (推論のみ) → Bedrock か g5/g6/inf2 インスタンス

Q2. 事前学習 or ファインチューニング？
  事前学習 → p5.48xlarge × 複数ノード（コスト高、SPReAD の 500 万円では小規模のみ可能）
  ファインチューニング → Q3へ

Q3. モデル規模は？
  〜7B → g6e.12xlarge or p4d.24xlarge (LoRA)
  7〜30B → p4d.24xlarge
  30〜70B → p4de.24xlarge or p5.48xlarge
  70B+ → p5.48xlarge マルチノード
```

## Spot / Savings Plans の検討

- **Spot**：最大70〜90%OFF。ただし中断リスク。**チェックポイント頻繁保存前提**でないと使えない。短期研究ではリスク／リターン検討要。
- **Savings Plans (1年)**：20〜30%OFF 相当。研究期間6ヶ月には**不向き**（残期間無駄）。
- **Reserved Instances**：同上、不向き。
- **結論**：SPReAD の6ヶ月研究では基本オンデマンド。Spot は **実験用途で中断許容できる場合のみ** 部分適用。

## `fetch_aws_price.py` を呼ぶときのテンプレート

```bash
# EC2 GPU インスタンスのオンデマンド時間単価
python3 scripts/fetch_aws_price.py ec2 p4d.24xlarge ap-northeast-1
python3 scripts/fetch_aws_price.py ec2 g5.xlarge ap-northeast-1

# Bedrock トークン単価（入出力）。
# デフォルトモデルは Claude Opus 4.7、デフォルトリージョンは ap-northeast-1（東京）。
# ユーザーが他モデル・他リージョンを明示した場合のみそちらを使う。
# 例: anthropic.claude-opus-4-7 / amazon.titan-text-express-v1 / meta.llama3-3-70b-instruct-v1:0 など
python3 scripts/fetch_aws_price.py bedrock anthropic.claude-opus-4-7 ap-northeast-1 --io input
python3 scripts/fetch_aws_price.py bedrock anthropic.claude-opus-4-7 ap-northeast-1 --io output

# S3 Standard
python3 scripts/fetch_aws_price.py s3 standard ap-northeast-1
```

内部では `https://pricing.us-east-1.amazonaws.com/offers/v1.0/aws/index.json` からメタデータを辿り、必要なサービスの regional JSON を ijson でストリーム取得、対象 SKU が見つかり次第ダウンロードを中断する（全体を落とさない）。取得した単価は1週間キャッシュされる。

## ベストプラクティス（AWS 公式ドキュメント参照）

- Spot の中断リスクとチェックポイント戦略：https://docs.aws.amazon.com/AWSEC2/latest/UserGuide/using-spot-instances.html
- Savings Plans / Reserved の使い分け：https://docs.aws.amazon.com/savingsplans/latest/userguide/what-is-savings-plans.html
- SageMaker と EC2 の使い分け：https://docs.aws.amazon.com/sagemaker/latest/dg/whatis.html

SPReAD の研究期間（最大6ヶ月）では、1年以上コミットする Savings Plans/Reserved は不向き。基本オンデマンド、Spot は中断許容できる学習ワークロードでのみ部分適用が現実解。

## 補足：HPCI 特定研究有償課題（参考、本スキルでは扱わない）

第2回公募要領には、SPReAD 採択者向けに **HPCI（革新的ハイパフォーマンス・コンピューティング・インフラ）** の **「特定研究有償課題」** が新設されたことが明記された（令和8年6月上旬募集開始予定）。HPCI は富岳・東大／京大／東北大スパコン等を共通窓口でまとめて使える共用プラットフォームで、本枠は通常の課題審査を経ずに最低限の資格審査のみで1週間以内に利用開始でき、利用報告書・成果公開義務も原則免除される。詳細は HPCI ポータルサイトの「特定研究有償課題」ページを参照（公募要領に URL 記載あり）。

本スキルは AWS を計算資源とする前提のため HPCI ヒアリング分岐は持たないが、4枚目の **計算資源費用テーブルの「選定理由」(E列)** で「データセンター連携／急ぎの利用開始／既存の AWS ベースのワークフロー との統合容易性」等を理由に AWS を選んだ旨を補強する材料として使うことができる。
