---
name: spread-proposal-assistant
description: AI for Science萌芽的挑戦研究創出事業（SPReAD）第2回公募への応募者向けに、研究計画のヒアリングから「第2回_様式1_研究計画調書.xlsx」の作成までを伴走支援するスキル。研究者はAWSを計算資源として利用する前提。第2回公募要領・様式の内容を踏まえ、ヒアリングで情報を引き出し、文字数制限・費目区分・AWS利用料の算定根拠までを整えた上でExcelに埋め込む。日本語版・英語版テンプレートに対応。ユーザーが「SPReAD」「AI for Science萌芽」「研究計画調書を作りたい」「様式1を埋めたい」と言及した場合に起動する。
license: Proprietary
---

# SPReAD 研究計画調書 作成支援スキル（第2回公募版）

## 使用タイミング

以下のいずれかに該当する場合に本スキルを使う。

- ユーザーが SPReAD（AI for Science 萌芽的挑戦研究創出事業）第2回公募に応募したい
- ユーザーが「第2回_様式1_研究計画調書.xlsx」を埋めたい／下書きしたい
- AWS を計算資源として使う前提で費用見積もりが必要

第2回公募の応募期間は **令和8年6月2日（火）〜令和8年7月3日（金）正午**、研究期間は交付決定日〜令和9年2月5日。

## 第1回公募からの主要な変更点（重要）

第1回公募（令和8年4月）は終了しており、本スキルは第2回専用に書き換えられた。

- **AIインタビューの完全廃止**：旧5枚目（AIインタビュー実施完了報告書）が様式から削除。メールアドレス一致要件も撤廃。
- **様式1のセル位置がほぼ全面変更**：1〜4枚目のすべてのセル位置が刷新された。
- **新規項目「成果の公開方針（任意）」**が2枚目に追加（D13、日本語≤150字／英語≤90 words）。
- **「区分」が2項目に分離**：「所属機関の区分」（C21）と「応募者属性の区分」（C23、新設）。
- **3枚目は行追加禁止**：各テーブル20行固定枠、設備備品費の金額はテンプレ式 `=単価×数量`。
- **サブユースケース／AI活用度は `Y` 文字方式**（旧 True/False から変更）。
- **業績欄は1セル改行5件 → 5セルに1件ずつ**（D14:D18）。本人を含む業績のみ。
- **英語版様式 `2nd_Form1_Research Plan.xlsx` が新規提供**：日本語または英語のいずれかで応募可。
- **公式形式チェックツール `research_plan_self_check_v1.py` が文科省から配布**：本スキルの `validate.py` の後に実行を案内する。
- **重複制限**：第1回採択者は第2回応募不可。第1回不採択者は再応募可。ARiSE（共同代表者）との重複応募不可。

## スキルの全体方針

本スキルは**対話ヒアリング→ドラフト生成→xlsx 書き込み→検証**の4段構えで進める。ユーザーの時間を尊重し、1回に聞くのは1テーマに絞る。不明点は仮置きし、最後に一括で確認を取る。

### 利用するLLMモデルの方針

研究計画調書は**研究費獲得を左右する重要書類**であり、ヒアリング応答の解釈・本文ドラフトの作成・要約・選定理由の作文には**そのエージェント環境で利用可能な最新の Claude モデル**を使用すること。

- 本文ドラフト作成（研究目的・研究方法・AI利活用の妥当性・達成目標・ノウハウ・選定理由・算定根拠の補足説明等）は、必ず**最新の Claude フラッグシップモデル**（その時点での最新の Claude フロンティアモデル、執筆時点では Claude 4 系統または以降）で行う。
- Haiku などの小型・高速モデルは、**正確性と論理性が要求される本文の作文には使わない**（要約や定型変換のみに限定）。
- モデルID をハードコードしない。エージェント環境（Claude Quick Desktop, Claude Code 等）が選択した最新モデルを使う。ユーザーが「軽量モデルでよい」と明示しない限り、品質優先で動作する。
- スキル内のヘルパー関数（`window.cowork.askClaude` 等が利用可能な場合や、エージェントから別モデルへの委譲手段が露出している場合）でモデルを指定する場合も、最新の Claude モデルを指定する。

### 前提・ファイル配置

```
spread-proposal-assistant/
├── SKILL.md                                ← 本ファイル
├── references/                             ← ガイド・マスタ情報
│   ├── overview.md                         (調書全シートの構造・セル一覧・数式)
│   ├── lists.md                            (研究領域・ユースケース・所属機関の区分・応募者属性の区分)
│   ├── sheet1_basic.md                     (1枚目の記入項目とヒアリング順)
│   ├── sheet2_research.md                  (2枚目のプロンプト集・文字数制限)
│   ├── sheet3_budget.md                    (3枚目の費目設計ガイド)
│   ├── sheet4_aws_cost.md                  (4枚目 AWS 費用算定プロトコル)
│   └── aws_gpu_instances.md                (代表的GPUインスタンスと選定指針＋HPCI補足)
├── scripts/                                ← xlsx 書込・費用計算・検証スクリプト
│   ├── compute_aws_cost.py                 (ヒアリング値から費用表を生成)
│   ├── fetch_aws_price.py                  (AWS Price List Bulk API から単価取得)
│   ├── fill_workbook.py                (様式1への書き込み、XML文字列置換版、--language ja|en)
│   └── validate.py                         (文字数・必須項目・合計一致の検証)
└── call_materials/                         (公募要領等)
    ├── 2nd_spread_call_for_proposals.pdf    (第2回公募要領)
    ├── application_forms_ja/               ← 日本語版様式
    │   ├── 2nd_Form0_Application Form Checklist.docx
    │   ├── 2nd_Form1_Research Plan.xlsx     ← テンプレート本体（日本語版）
    │   ├── 2nd_Form2_Consent Confirmation Form Regarding Review Methods and Handling of Application Information.docx
    │   ├── 2nd_Form3_Student Application Consent Confirmation Form.docx
    │   └── 2nd_Form4_Advisor Consent Confirmation Form.docx
    ├── application_forms_en/               ← 英語版様式
    │   ├── 2nd_Form0_Application Form Checklist.docx
    │   ├── 2nd_Form1_Research Plan.xlsx     ← テンプレート本体（英語版）
    │   ├── 2nd_Form2_Consent Confirmation Form Regarding Review Methods and Handling of Application Information.docx
    │   ├── 2nd_Form3_Student Application Consent Confirmation Form.docx
    │   └── 2nd_Form4_Advisor Consent Confirmation Form.docx
    └── e_rad_guide/
        └── e_rad_operation_guide.pdf       (e-Rad 操作・入力ガイド、第2回別紙)
```

**交付要綱**（次世代人工知能技術等研究開発拠点形成事業費補助金 交付要綱、令和8年1月9日文部科学大臣決定）は第2回公募の配布物には同梱されていない。採択後の交付申請手続で参照する規程で、文部科学省 SPReAD ウェブサイト（https://www.mext.go.jp/aifors_spread/）から別途取得できる。スキルでは同梱しない。

**形式チェックツール**（`research_plan_self_check_v1.py` / `form_self_check_v1.py`）も再配布禁止のライセンスのため本スキルには同梱しない。提出前の最終確認で必ずユーザーにダウンロードを案内すること（Step 7 参照）。配布元 Box フォルダ：https://mext.ent.box.com/s/qf8vbuj3pso1hj9mwp1vs6rx2hashpuc/folder/385379226260

**重要**：xlsx の**シート名**は Excel ファイル**内部**の名称なので、zip ファイル名の互換性とは無関係。日本語版テンプレートでは `研究計画調書_1枚目` 等、英語版テンプレートでは `Research Plan_Sheet 1` 等。

パスの扱い：

- `SKILL.md`・`references/`・`scripts/`・`call_materials/` はすべて**同じ親ディレクトリ**配下（以下これを `{SKILL_ROOT}` と呼ぶ）
- スクリプトは `__file__` から `{SKILL_ROOT}` を自動解決するので、絶対パスを渡さなくても動く（`fill_workbook.py` の `--template` は省略可、`--language` だけで日英切替可能）
- 最終成果物は `{SKILL_ROOT}/../`（ユーザー作業フォルダ直下）に保存

その他の前提：

- xlsx 書き込みは **`scripts/fill_workbook.py`** を使う。テンプレートの構造を完全保持しながら対象シートの空セルだけに値を挿入する方式。詳細仕様と動作は **Step 6** を参照。
- AWS 費用は `scripts/fetch_aws_price.py`（AWS 公開料金 JSON を都度取得）を利用。MCP は使わない。

### 出力物

1. 埋め込み済みの `第2回_様式1_研究計画調書_{e-Rad所属機関コード}_{ローマ字氏名}.xlsx`（英語版を選ぶ場合は `2nd_Form1_Research Plan_{e-Rad Institution code}_{Name}.xlsx`）
2. ヒアリング内容を要約した `hearing_notes.md`（内部メモ）
3. AWS 費用算定の内訳 `aws_cost_breakdown.md`（4枚目の算定根拠ソース）

中間生成物は `{WORK_DIR}`（エージェント作業領域）に保存。最終 xlsx は**エージェントが現在書き込み可能な作業フォルダ**（ユーザーが選択しているフォルダ／プロジェクトルート等）に保存し、`computer://` リンクで提示する。エージェント環境によってはユーザー指定の任意パス（`~/Documents/SPReAD/` 等）に書き込めないため、**出力先をユーザーに確認しない**ことを推奨する（書き込み失敗の原因になる）。

> ⚠️ **生成 xlsx をユーザーに引き渡す際の注意**：本スキルが生成する xlsx は数式キャッシュ（文字数カウント・経費合計）が `<v>0</v>` のままで、Excel が起動時の自動再計算で正しい値を表示する設計（`fullCalcOnLoad="1"` 付き）。ユーザー自身に「**最終提出・公式形式チェックツール実行の前に、必ず一度 Excel で開いて『名前を付けて保存／コピーを保存』で同じファイル名で保存し直してください**」と案内すること。再保存なしに公式チェックツールにかけると「文字数 0」「合計 0 円」と誤判定される。詳細は Step 7-2C 参照。

## 推奨ワークフロー

### Step 0: 準備

1. **スキル同梱ファイル `{SKILL_ROOT}` の所在とアクセス方法を確定する**（最重要）。
   - エージェント実行環境によって、スキルは異なるパス配下に配置される。代表例：
     - **Claude Quick Desktop**：`~/.quickwork/profiles/federate-prod/skills/spread-proposal-assistant/`
     - **Claude Code（プラグイン）**：`~/.claude/plugins/cache/<owner>/<plugin>/skills/spread-proposal-assistant/`
     - **その他**：エージェントごとに異なるので、まず実際のパスを特定する
   - **アクセス方法に注意**：Quick Desktop など一部の環境では、ファイルツール（`Read`, `Glob`, `Grep`, `fd`, `ripgrep` 等）は「許可フォルダ」制限により当該パスを読めないことがある。一方、Python 経由（`run_python` / 同等の Python 実行ツール）の `os.path.exists()` `open()` `subprocess.run` は通る。アクセス可否は環境依存。
   - **手順**：以下の順で `{SKILL_ROOT}` を確定する。
     1. ファイルツールで `references/overview.md` を Read してみる。読めればファイルツール経由で進める。
     2. 読めなければ Python から `os.path.exists("{SKILL_ROOT}/references/overview.md")` を確認する。Python 経由で読み込めれば、以降の Read 操作は Python の `open()` と `Path.read_text()` を使う。
     3. どちらも不可ならユーザーにスキル本体のパスを確認する。
   - 以降の Step では `{SKILL_ROOT}/references/...` `{SKILL_ROOT}/scripts/...` `{SKILL_ROOT}/call_materials/...` を参照するが、上で確定した方法で読む／実行する。

2. **出力先 `{OUTPUT_DIR}` の決定（ユーザーには確認しない）**。
   - エージェントが現在書き込み可能な作業フォルダ（ユーザーが選択しているフォルダ、プロジェクトルート、エージェントの作業領域直下等）を `{OUTPUT_DIR}` として使う。
   - **ユーザーに任意のパスを聞かない**こと。`~/Documents/SPReAD/` のような外部パスは多くのエージェント環境で書き込み許可がなく、書き込み失敗の原因になる。
   - 中間生成物（`payload.json` `hearing_notes.md` `aws_cost_breakdown.md` 等）は `{WORK_DIR}` 配下（エージェントの作業領域）でよい。最終 xlsx も同じく書き込み可能な領域に出力する。
   - 完成後、`computer://` リンクで提示する。ユーザーが別の場所に保存したい場合は手動でコピーしてもらう。

3. `{SKILL_ROOT}/references/overview.md` を読み、調書の全シート構成・記入項目・文字数制約を把握する。

4. **応募言語（日本語／英語）を確認**する。多くは日本語なのでデフォルトで日本語を提案。英語応募の場合は `--language en` でテンプレートを切り替える。

5. 公募要領PDFの重要箇所は必要に応じて参照する：`pdftotext -layout "{SKILL_ROOT}/call_materials/2nd_spread_call_for_proposals.pdf" -`

### Step 1: キックオフヒアリング（AskUserQuestion）

最初に以下5点をまとめて AskUserQuestion で聞く：

1. 研究課題名（仮で可）
2. 研究領域（以下の11区分から選択。**このリストをそのまま提示せよ。記憶や推測で代替してはならない。**）
   1. 臨床科学
   2. 生命科学・薬学
   3. 化学
   4. 機械・社会基盤・エネルギー工学
   5. 材料・プロセス・応用医工学
   6. 電気工学・電子工学・情報科学・コンピューターサイエンス
   7. 数学・物理学・地球科学
   8. 農学・環境学・生態学
   9. 社会科学
   10. 芸術・人文科学
   11. その他
3. メインユースケース分類（9区分から選択、「9.その他」は最後の手段）
4. **第1回公募への応募経験**（採択 → 応募不可、不採択 → 再応募可、未応募 → 通常通り）
5. ARiSE への共同代表者としての応募予定の有無（あれば本事業 SPReAD への応募は審査対象外になる）

「9.その他」を選んだ場合は自由記述（C31）も後で必ず聞く。

### Step 2: 基本情報の確認

`references/sheet1_basic.md` を Read し、記入項目を順にヒアリングする。e-Rad 研究者番号、所属機関、メールアドレス、生年月日等は**必須だが推測禁止**なのでユーザーに直接聞く。

第2回公募で「区分」が2項目に分離されたので、必ず両方を確認：
- **所属機関の区分（C21）**：大学／高等専門学校／公的研究機関／民間企業／非営利団体・公益法人／その他
- **応募者属性の区分（C23）**：教員／研究員(ポスドク含む)／博士課程学生／修士課程学生／学部学生／技術職員・URA等／技術者・開発者／その他

学生（博士／修士／学部）の場合は L20 に「Y」を入れ、別途 様式3・様式4 の提出が必要であることを案内。

### Step 3: 研究内容ヒアリング（2枚目）

`references/sheet2_research.md` のプロンプト集に従い、以下を順番にヒアリングする。各項目は**文字数下限・上限が厳密**なので、ヒアリング後にドラフトを生成し、文字数を確認してから書き込む。

- 研究目的（D8、日本語80〜400字／英語48〜240 words）
- 研究方法（D9、日本語160〜800字／英語96〜480 words）
- AI 利活用の妥当性・実現可能性（D10、日本語160〜800字／英語96〜480 words）
- 達成目標（D11、日本語100〜500字／英語60〜300 words）
  - **第2回からは「最終的な到達目標を簡潔に示した上で、3か月／6か月の中間・終了時目標を区分」する書きぶり**
- ノウハウ抽出・共有計画（D12、日本語60〜300字／英語36〜180 words）
- **成果の公開方針（D13、任意、日本語≤150字／英語≤90 words）**：第2回新設項目
- 研究業績等（D14, D15, D16, D17, D18、最大5件、各セルに1件、本人を含むもののみ）

ヒアリング時は open question→深掘り→要約確認の順で行う。1項目につき2〜3往復で収束させる。

### Step 4: 研究経費の設計（3枚目）

`references/sheet3_budget.md` を Read。直接経費 **10万円〜500万円** の制約下で、費目（設備備品費／消耗品費／謝金／国内旅費／外国旅費／その他）の配分をヒアリングする。**人件費は対象外**。本事業では計算資源（AWS）は主に「その他」に計上し、詳細を4枚目に記載する。GPUカード等を物品として購入する場合は設備備品費に計上した上で、4枚目「計算資源費用」テーブルにも詳細を書く。

**単位ルール**：ユーザーとの会話・ドラフトは **万円** で統一。xlsx 様式の金額欄は「千円」単位で固定されているので、書き込む時のみ `千円 = 万円 × 10` に換算し、千円未満は切り捨て。

**3枚目は第2回から行追加禁止**（各テーブル20行固定枠）。20行を超える明細が必要な場合はユーザーに事項を集約してもらう。

設備備品費の金額（J列）は `=単価(G列)×数量(I列)` のテンプレ式が入っているので、`G{row}` と `I{row}` だけ書き込む。

**必要性記入欄（C34／C63／C92）は3枚とも必ず埋める**。空欄のまま提出すると「審査員が経費の妥当性を読み取れない」という形式不備になる。割合が90%超でなくても**簡潔な必要性説明を記載**する。文字数制限はないが、それぞれ2〜4文程度で具体的に書く：

- **`C34` 設備備品費・消耗品費の必要性**：当該物品が研究遂行に不可欠である理由（性能要件・代替手段との比較等）。当該費目を計上していない場合は「本研究では設備備品費・消耗品費の計上はない」と一行記載する。
- **`C63` 謝金・旅費の必要性**：謝金・旅費の使途と研究遂行上の必要性（学会参加目的・現地調査の意義等）。計上していない場合は「本研究では謝金・旅費の計上はない」と一行記載する。
- **`C92` その他費用の必要性**：SPReAD では AWS 計算資源費用が「その他」に集中して大半を占めるケースが多いため、**特に丁寧に記載**する。AWS を計算資源として選定した理由、研究遂行上不可欠である根拠、4枚目（API・計算資源費用）との対応関係を明示する。

ヒアリング時に各費目の使途を聞いた直後、その内容をまとめて payload の `equipment_consumables_necessity` / `honorarium_travel_necessity` / `other_necessity` に格納する。ユーザーが理由を言語化していない場合はエージェント側で要約してドラフトを提示し、確認を取る。

### Step 5: API・計算資源の算定（4枚目）

`references/sheet4_aws_cost.md` を Read。本スキルの中核ステップ。

> ⚠️ **重大な原則**：4枚目の API 費用テーブルに記載できるのは **AWS サービスのみ**（Bedrock, EC2 推論エンドポイント, Comprehend, Translate 等）。**Elsevier Text Mining API、Qdrant Cloud、OpenAI 直接、Pinecone、Hugging Face Inference Endpoints などのサードパーティ API は、ユーザーが自発的に「使いたい」と明示した場合のみ**記載する。エージェント側から「研究内容に合いそうだから」という理由で**提案・創作・追加することは厳禁**。SPReAD は AWS 利用が前提の制度設計であり、勝手にサードパーティ API を埋めると審査時に疑義を招く。詳細は `references/sheet4_aws_cost.md` の冒頭ガード句を参照。

#### 5-1. AWS 利用想定ヒアリング

以下を AskUserQuestion で確認：

- 主な用途（LLM ファインチューニング／推論／データ前処理／ベクター検索 等）
- 想定 GPU インスタンス（A100／H100／L40S／推論用 等、未定なら「要選定」）
- 見込み稼働時間（時間/月 × 月数、**研究期間は最大約6ヶ月**。12ヶ月や年額ベースで計算しないこと）
- データ容量（S3 保管量・転送量）
- 使用予定 AWS サービス（EC2, SageMaker, Bedrock, S3, EFS, ECR 等）
- **APIを使うか**：API を使う／使わない（GPU 学習・推論のみ）のどちらか。**API利用そのものの要否だけ**を聞く（GPU での自前学習・推論のみで完結する研究では API 費用テーブル＝行9〜18 は空欄のまま提出）。
- **モデル・リージョン選定はヒアリングしない**（質問項目を増やさないため）。API を使う場合の **デフォルトは `Amazon Bedrock 上の Claude Sonnet 4.6`、リージョンは `ap-northeast-1`（東京）**。算定根拠の `[備考]` に「研究者は具体モデル未定。本算定では Claude Sonnet 4.6（東京リージョン、JP CRIS対応。CRIS利用時もソースリージョン=東京の単価で課金）を仮置きし、応募後にモデル選定を詰める想定」と明記する。ユーザーが自発的に具体モデル名・他リージョンを出した場合のみ、そちらを優先する。
- **サードパーティ API（AWS 以外）は提案禁止**：Elsevier・Qdrant・OpenAI・Anthropic 直接・Pinecone・Weaviate・Cohere・Hugging Face Inference Endpoints 等は、ユーザーが「○○を使いたい」と自発的に言わない限り**追加しない**。研究内容上必要そうに思えても、エージェント側から提案しない。必要かどうかを確認したい場合は AskUserQuestion で「この研究で AWS 以外の外部 API を使う予定はありますか？」と明示的に聞く。

#### 5-2. AWS 公開料金 API から**必要なサービスだけ**取得

**MCP サーバーは使わない**（本スキル環境で利用不可のため）。代わりに AWS Price List Bulk API（認証不要、公開エンドポイント）を利用する。アクセスフロー詳細と CLI 例は `references/sheet4_aws_cost.md` 参照。

```bash
python3 spread-proposal-assistant/scripts/fetch_aws_price.py ec2 p4de.24xlarge ap-northeast-1
# Bedrock を使う場合のみ。デフォルトは Claude Sonnet 4.6 / ap-northeast-1（東京）
python3 spread-proposal-assistant/scripts/fetch_aws_price.py bedrock anthropic.claude-sonnet-4-6-v1 ap-northeast-1 --io input
python3 spread-proposal-assistant/scripts/fetch_aws_price.py bedrock anthropic.claude-sonnet-4-6-v1 ap-northeast-1 --io output
python3 spread-proposal-assistant/scripts/fetch_aws_price.py s3 standard ap-northeast-1
```

**使うサービスに絞って都度取得**する方針。取得できない場合（新しい SKU 等）は AWS 公式料金ページをフォールバックとし、**参照日と情報源を算定根拠欄に明記**する。

#### 5-3. 算定根拠を表に整形

`scripts/compute_aws_cost.py` を呼び出し、ヒアリング値＋`fetch_aws_price.py` から取得した単価を掛け合わせて月額／総額（千円単位）を算出。同スクリプトが `aws_cost_breakdown.md` を生成し、4枚目に書き込むテキストを準備する：

- API費用テーブル（行9〜18、列 D処理対象 / E金額 / F算定根拠）— 該当時のみ（Bedrock 等）
- 計算資源費用テーブル（行22〜31、列 D GPU種類 / E選定理由 / F金額 / G算定根拠）

選定理由では「学習データ量・モデル規模から A100 80GB 以上が必要」等、研究要件と結びつけて書く。算定根拠では「$X/時 × Y時間/月 × Zヶ月 × 為替150円/USD = ◯千円」のように式を明示する。**月数Zは研究期間（交付決定日〜令和9年2月5日、最大約6ヶ月）を超えないこと。「×12ヶ月」や年額ベースの計算は誤り。** **為替レートはユーザーに確認（所属機関の会計基準に従う）。**

### Step 6: xlsx への書き込み

`scripts/fill_workbook.py` を使う。テンプレ xlsx を ZIP として開き、対象シートの空セルだけを XML 文字列置換で書き換える。テンプレの構造（条件付き書式・データ検証・名前空間・customXml・printerSettings 等）はバイト単位でそのまま保持される。

> なお `anthropic-skills:xlsx` の `scripts/recalc.py`（LibreOffice 経由の再計算）や、openpyxl ベースの load→save は**使わない**。いずれも xlsx の OPC 構造や条件付き書式・リッチテキストを破壊し、Excel が起動時に「ファイルレベルの検証で問題が見つかりました」修復ダイアログを出す原因になる。

#### `fill_workbook.py` の動作

スクリプトは以下を自動で行う：

1. **空セル `<c r="..." s="..."/>` への値挿入**：
   - 文字列値は `<c r="..." s="..." t="inlineStr"><is><t xml:space="preserve">VALUE</t></is></c>` に置換
   - 数値値（e-Rad番号、生年月日、金額千円等）は `<c r="..." s="..."><v>VALUE</v></c>` に置換
   - 既値セル・数式セル・他セルは一切触らない
2. **`xl/workbook.xml` の `<calcPr>` に `fullCalcOnLoad="1"` を強制付与**：これにより Excel が開いた瞬間に文字数カウント（`=LEN(C35)` 等）・経費総計（`=SUM(E48:L48)` 等）・設備備品費の各行金額（`=G*I`）がすべて自動再計算される。
3. **数式インジェクション対策**：`=` `+` `-` `@` `\t` `\r` `\n` で始まる文字列値は `'` で前置エスケープ（CSV インジェクション類似攻撃を防止）。
4. **テンプレ全パーツのバイト保持**：上記で書き換える `xl/worksheets/sheet3.xml`（1枚目）〜 `sheet6.xml`（4枚目）と `xl/workbook.xml` 以外は ZIP コピーのみ。

呼び出し例（`{OUTPUT_DIR}` はエージェントが書き込み可能な作業フォルダ）：

```bash
# 日本語版（デフォルト）
python3 "{SKILL_ROOT}/scripts/fill_workbook.py" \
  --data "{WORK_DIR}/payload.json" \
  --output "{OUTPUT_DIR}/第2回_様式1_研究計画調書_{機関コード}_{氏名}.xlsx"

# 英語版
python3 "{SKILL_ROOT}/scripts/fill_workbook.py" \
  --language en \
  --data "{WORK_DIR}/payload.json" \
  --output "{OUTPUT_DIR}/2nd_Form1_ResearchPlan_{Institution code}_{Name}.xlsx"
```

`payload.json` の構造はスクリプト内の項目割当を参照（1枚目: `name_kanji`, `email`, `title` 等／2枚目: `purpose`, `method`, `ai_rationale`, `goals`, `knowhow`, `publication_policy`, `achievements` 等／3枚目: `equipment_rows`, `consumables_rows`, `honorarium_rows`, `domestic_travel_rows`, `foreign_travel_rows`, `other_rows`, `*_necessity` 等／4枚目: `api_rows`, `compute_rows`）。

#### ⚠️ ファイル名の規則（公式形式チェックツールが厳密に検査する）

文部科学省の公式形式チェックツールは以下の正規表現で厳密にチェックする。違反すると「ファイル名 NG」となり書類不備扱いになる：

- 日本語版: `^第2回_様式1_研究計画調書_<半角数字>_<氏名>$`
- 英語版: `^2nd_form1_researchplan_<半角数字>_<Name>$`（大文字小文字無視）

**重要：氏名部分にアンダースコア `_` を入れてはいけない**。`_` は固定区切り4個ぶんしか許容されないため、`_DRAFT` `_v2` `_山田_太郎` のように追加 `_` が入ると NG 扱いになる。

具体例（**○ = 通る／× = NG**）：

```
○  第2回_様式1_研究計画調書_1234567890_山田太郎.xlsx
○  第2回_様式1_研究計画調書_1234567890_YamadaTaro.xlsx
○  2nd_Form1_ResearchPlan_1234567890_YamadaTaro.xlsx
×  第2回_様式1_研究計画調書_1234567890_YAMADA_Taro.xlsx       ← 姓名間の _
×  第2回_様式1_研究計画調書_1234567890_山田太郎_DRAFT.xlsx     ← 末尾の _DRAFT
×  第2回_様式1_研究計画調書_1234567890_山田 太郎.xlsx          ← 半角スペース
```

氏名は **`_` を含めない／スペースを含めない** 連結形に整形してからファイル名に組み込むこと。下書きを示すサフィックス（_DRAFT 等）は **ファイル名ではなく拡張子前の括弧書き** などにして、提出前に必ず削除する。あるいは下書きフェーズはファイル名を仮で扱い、提出直前に正規ファイル名へリネームする運用を推奨する。

重要：
- 既存セルの数式・書式は絶対に変更しない（`references/overview.md` の数式一覧を参照）
- 未入力セル（薄オレンジ）は入力必須ではない項目もあるので、ユーザーが空欄を選んだ項目はそのまま
- 文字数カウント列（E列, N列）は既存数式を保持
- サブユースケース・AI活用度は **`Y` 文字** を入れる方式（True/False ではない）

### Step 7: 検証

最後に以下を順に実施：

1. **本スキル同梱の `scripts/validate.py` を実行**
   ```bash
   python3 "{SKILL_ROOT}/scripts/validate.py" "{WORK_DIR}/第2回_様式1_研究計画調書_…xlsx"
   ```
   - 各文字数制限の上下限チェック（成果の公開方針も含む）
   - 必須項目の未入力検出
   - 直接経費の上下限（10万円〜500万円）
   - 1枚目総計と3枚目各費目小計の整合
   - 90%超費目の必要性記載確認

2. **文部科学省公式の形式チェックツールを実行**（提出直前の最終確認、強く推奨）

   公式ツール `research_plan_self_check_v1.py`（様式1用）と `form_self_check_v1.py`（様式0/2/3/4 用）は文科省が再配布禁止のライセンスで配布しているため、**本スキルには同梱しない**。ユーザー自身に Box フォルダからダウンロードしてもらう。

   #### 2-A. ダウンロード（初回のみ、5分）

   - **配布先**：第2回 SPReAD 公募「形式チェックツール関連」フォルダ
     - URL: https://mext.ent.box.com/s/qf8vbuj3pso1hj9mwp1vs6rx2hashpuc/folder/385379226260
   - フォルダ内の以下5ファイルをまとめてダウンロード（Box の「ダウンロード」ボタンで ZIP 一括取得が早い）：
     - `README_v1.md`（手順書）
     - `【形式チェックツール】利用ガイド.pdf`（日本語ガイド）
     - `User Guide for the Formality Check Tool.pdf`（英語ガイド）
     - `research_plan_self_check_v1.py`（様式1用）
     - `form_self_check_v1.py`（様式0/2/3/4 用）

   #### 2-B. 環境準備（初回のみ、10分）

   - **Python 3.11 以上**をインストール
     - Windows: https://www.python.org/downloads/ から最新版をダウンロード→**「Add python.exe to PATH」に必ずチェック**を入れてインストール
     - Mac: 同サイトから `.pkg` をダウンロード→ダブルクリックでインストール
   - **必要ライブラリをインストール**（コマンドプロンプト／ターミナルで実行）
     - Windows: `pip install openpyxl pdfminer.six pypdf`
     - Mac: `pip3 install openpyxl pdfminer.six pypdf`

   #### 2-C. 様式1チェックの実行（毎回、3分）

   > ⚠️ **【重要】チェックツールにかける前に必ず Excel で一度開いて再保存する**
   >
   > 本スキルが生成した直後の xlsx は、文字数カウント（`=LEN(C35)` 等）や経費総計（`=SUM(...)`）の数式が**まだ計算されていない状態**でファイルに格納されている（`fullCalcOnLoad="1"` により Excel で開いた瞬間に計算される設計のため）。
   >
   > このまま公式形式チェックツールにかけると、ツールは数式キャッシュをそのまま読むため、**実際には記入済みの内容が「文字数 0」「合計 0 円」と判定される**。これを避けるには、ツール実行前に必ず以下の手順で Excel に**保存**させる必要がある：
   >
   > 1. 生成された `第2回_様式1_研究計画調書_<機関コード>_<氏名>.xlsx` を **Excel で開く**
   > 2. 内容に問題がないか目視確認（文字数カウント・経費合計が正しく表示されることもこの時点で確認できる）
   > 3. **「ファイル」→「名前を付けて保存」または「コピーを保存」**で、提出用フォルダ（例：「様式1チェック」）に **同じファイル名のまま保存**する（Excel が再計算した結果のキャッシュを xlsx に書き戻す目的）
   >    - **重要**：ファイル名規則を守るため、保存ダイアログで `_DRAFT` `_v2` のような追加サフィックスを付けない。提出時の正規ファイル名（`第2回_様式1_研究計画調書_<10桁機関コード>_<氏名>.xlsx`）と完全に一致させる
   > 4. Excel を閉じる
   >
   > これで再保存後の xlsx は数式キャッシュが確定し、公式チェックツールでも正しく評価される。

   1. デスクトップなどに「**様式1チェック**」フォルダを作成
   2. 以下2ファイルを同フォルダに入れる
      - `research_plan_self_check_v1.py`（ダウンロードしたツール、ファイル名から `_v1` を取らずそのままで動く。SKILL.md 上の表記は `research_plan_self_check.py` だが配布版は `_v1` 付きのまま使ってよい）
      - **【上記の手順で Excel から再保存した】** `第2回_様式1_研究計画調書_<機関コード>_<氏名>.xlsx`
   3. そのフォルダでコマンドプロンプト／ターミナルを開く
      - Windows: フォルダのアドレスバーに `cmd` と入力して Enter
      - Mac: ターミナルで `cd ` の後にフォルダを Finder からドラッグ＆ドロップして Enter
   4. 実行コマンド：
      - Windows: `python research_plan_self_check_v1.py`
      - Mac: `python3 research_plan_self_check_v1.py`
   5. 同じフォルダに `研究計画調書_セルフチェック結果_(YYYYMMDD_HHMM).xlsx` が出力される

   #### 2-D. 結果の解釈

   出力 Excel には **「判定サマリー」** と **「詳細チェック」** の2シートがある：

   - **判定サマリー**で各ファイルの **「OK」**／**「要確認」** を見る
   - **「要確認」**なら**詳細チェック**で項目別の指摘を確認し、本スキルで生成元 `payload.json` を修正→再生成→再実行
   - 全項目「OK」になれば形式面は提出可能な状態

   #### 2-E. 様式0/2/3/4 のチェックも忘れずに

   研究代表者本人が記入する **様式0（チェックリスト）・様式2（同意確認書）** および学生応募時の **様式3・様式4** は別途PDFで提出する。これらも上記と**別フォルダ**（例：「様式0234チェック」）を作って `form_self_check_v1.py` で同様にチェックする：

   - Windows: `python form_self_check_v1.py`
   - Mac: `python3 form_self_check_v1.py`

   ※ 様式1とPDF系を**同じフォルダで実行しない**こと（結果が混在する）。
   ※ 文字を選択できるPDFが必要。スキャン画像PDFは「判断不可」と判定される。

   #### 2-F. ありがちなエラーと対処

   - `python is not recognized`：Python の PATH 設定漏れ（Windows）。インストールやり直し or Mac は `python3` を使う
   - `ModuleNotFoundError`：`pip install ...` の実行漏れ。再実行
   - ダブルクリックで一瞬閉じる：必ずコマンドラインから `python ...` で実行
   - **様式1で「未記入」「文字数不足」「合計0円」と誤判定**：本スキル生成直後の xlsx は数式キャッシュが空（`<v>0</v>`）の状態。**2-C の冒頭で説明した「Excel で開いて再保存」を必ず実施**してからツールを再実行すること。Excel が `fullCalcOnLoad="1"` で再計算した結果は、明示的に保存しない限り xlsx ファイルには書き戻されない

3. ユーザーに最終レビュー依頼：
   - 様式0チェックリストの記入と署名（電子署名可、直筆／Word記名も可）
   - 様式2の同意確認書の署名
   - 学生応募の場合は様式3・様式4の同意確認書
   - 機関等への提出（e-Rad 経由、令和8年7月3日金 12:00 厳守）

## 重要な制約・注意

- **調書は PDF 化せず xlsx のまま提出**（公募要領に明記）
- **ファイル名（厳密）**：日本語版 `第2回_様式1_研究計画調書_<半角数字 e-Rad機関コード>_<氏名>.xlsx` ／ 英語版 `2nd_Form1_ResearchPlan_<半角数字>_<Name>.xlsx`。**氏名部分にアンダースコア `_` を含めない／半角スペースを含めない**（`YAMADA_Taro`, `山田 太郎`, `_DRAFT` 接尾辞などはすべて NG）。`validate.py` がファイル名規則を検査する。
- **出力先はエージェントが書き込み可能な作業フォルダ**を使う。ユーザーに任意のパスを確認しない（`~/Documents/SPReAD/` 等は多くのエージェント環境で書き込み許可がなく、書き込み失敗の原因になる）。`computer://` リンクで提示すれば、ユーザーが必要に応じて手動で別の場所にコピーできる。
- **直接経費 10 万円以上 500 万円以下**、間接経費は直接経費の 30%（機関配分）
- 金額はユーザー会話では **万円**、xlsx 書込は **千円**（`千円 = 万円 × 10`、千円未満切り捨て）
- **人件費は対象外**
- **設備備品費・旅費・謝金のいずれかが90%超の場合は必要性を3枚目に記載**
- 色を付した図や文字もそのまま審査に付される
- 2枚目の図は最大1枚（行20以下に貼付、セル内挿入は不要）
- 様式0〜4の他書類も必要だが、本スキルは様式1のみ扱う（他は手動案内）
- **第1回公募への採択者は第2回応募不可**（不採択者は再応募可）
- **ARiSE 共同代表者との重複応募不可**（応募した場合は SPReAD は審査対象外）
- **電子署名推奨**（直筆／Word記名も可、いずれも本人意思に基づくこと）
- **応募締切**：令和8年7月3日（金）正午（厳守、引き戻し・再提出不可）
- **報告書提出**：令和9年3月上旬

## 参照ファイル

- `references/overview.md` — 調書全シートの構造・セル一覧・数式
- `references/lists.md` — 研究領域・ユースケース・所属機関の区分・応募者属性の区分のマスタ
- `references/sheet1_basic.md` — 1枚目の記入項目とヒアリング順
- `references/sheet2_research.md` — 2枚目のプロンプト集・文字数制限
- `references/sheet3_budget.md` — 3枚目の費目設計ガイド・府省共通経費取扱区分
- `references/sheet4_aws_cost.md` — 4枚目 AWS 費用算定プロトコル
- `references/aws_gpu_instances.md` — SPReAD向け代表的GPUインスタンスと選定指針＋HPCI補足
- `scripts/compute_aws_cost.py` — ヒアリング値から費用表を生成
- `scripts/fetch_aws_price.py` — AWS Price List Bulk API から単価取得
- `scripts/fill_workbook.py` — 様式1への書き込み（XML 文字列置換、`--language ja|en`）
- `scripts/validate.py` — 文字数・必須項目・合計一致の検証
