"""AWS 費用算定ヘルパー

ヒアリング値と MCP から取得した単価を受け取り、4枚目の
API費用テーブル / 計算資源費用テーブルに入れる値を計算する。

使い方:
    from compute_aws_cost import ComputeCost, ApiCost, summarize
    compute = [
        ComputeCost(
            gpu_type="NVIDIA A100 80GB × 8 (p4de.24xlarge)",
            rationale="70B LLM のフルファインチューニングに必要な80GB VRAM。",
            hourly_usd=40.96,
            hours_per_month=80,
            months=6,
            fx_jpy_per_usd=150,
            retry_buffer=0.20,
            notes="学習データ10GB × 3エポック、再実行率20%込み"
        ),
    ]
    apis = [
        ApiCost(
            target="科学文献の要約と情報抽出",
            input_tokens_per_month=500_000,
            output_tokens_per_month=100_000,
            input_usd_per_mtok=3.0,
            output_usd_per_mtok=15.0,
            months=6,
            fx_jpy_per_usd=150,
            notes="Bedrock の Claude Opus 4.7（ap-northeast-1）を使用（デフォルト）"
        ),
    ]
    print(summarize(compute, apis))

本スクリプトは純粋な計算用。価格の取得は `fetch_aws_price.py`
（AWS Price List Bulk API の JSON を都度ストリーム取得）で実行する。

単位について：
- 4枚目 xlsx テーブルは「金額(千円)」と様式で固定されているため、
  `total_thousand_jpy` が xlsx 書き込み用の値（千円単位）。
- 一方、人が読むメモ・サマリ（`render_markdown` 出力）は
  **万円単位**で表示する（千円だと桁が紛らわしいため）。
"""

from dataclasses import dataclass, field
from typing import List
import math


@dataclass
class ComputeCost:
    # 表示・書き込み用
    service_name: str    # 例 "Amazon EC2 p4de.24xlarge (A100 80GB×8)"
    gpu_type: str        # 例 "NVIDIA A100 80GB × 8 (p4de.24xlarge)"
    rationale: str
    # 算定パラメータ
    hourly_usd: float
    hours_per_month: float
    months: float
    region: str = "ap-northeast-1"
    fx_jpy_per_usd: float = 150.0
    instance_count: int = 1
    retry_buffer: float = 0.0  # 0.20 = 20%バッファ
    price_source: str = ""  # 単価の情報源（例 "AWS Price List Bulk API, 2026-04-15取得"）
    notes: str = ""

    @property
    def total_jpy(self) -> float:
        raw = (
            self.hourly_usd
            * self.hours_per_month
            * self.months
            * self.instance_count
            * self.fx_jpy_per_usd
        )
        return raw * (1 + self.retry_buffer)

    @property
    def total_thousand_jpy(self) -> int:
        """千円単位、千円未満切り捨て（xlsx 書き込み用）"""
        return math.floor(self.total_jpy / 1000)

    @property
    def total_man_jpy(self) -> float:
        """万円単位、小数1桁（表示用）"""
        return round(self.total_jpy / 10000, 1)

    @property
    def calc_breakdown(self) -> str:
        """xlsx の算定根拠欄に貼付するテンプレ形式文字列"""
        parts = [
            f"${self.hourly_usd:.3f}/時",
            f"× {self.hours_per_month:g}時間/月",
            f"× {self.months:g}ヶ月",
        ]
        if self.instance_count > 1:
            parts.append(f"× {self.instance_count}台")
        parts.append(f"× {self.fx_jpy_per_usd:.0f}円/USD")
        if self.retry_buffer > 0:
            parts.append(f"× (1+{self.retry_buffer:.0%}再実行バッファ)")
        calc = " ".join(parts)

        price_src = self.price_source or "AWS公式料金ページ参照"
        lines = [
            f"[用途] {self.rationale}",
            f"[サービス] {self.service_name}（{self.gpu_type}）",
            f"[リージョン] {self.region}、オンデマンド",
            f"[単価] ${self.hourly_usd:.3f}/時 ({price_src})",
            f"[稼働時間] {self.hours_per_month:g}時間/月 × {self.months:g}ヶ月"
            + (f" × {self.instance_count}台" if self.instance_count > 1 else ""),
        ]
        if self.retry_buffer > 0:
            lines.append(f"[再実行率] {self.retry_buffer:.0%} バッファ")
        lines.append(f"[計算] {calc} = {self.total_man_jpy:,.1f}万円")
        if self.notes:
            lines.append(f"[備考] {self.notes}")
        return "\n".join(lines)


@dataclass
class ApiCost:
    # 表示・書き込み用
    # api_type_ja: 審査員向けの「日本語で端的な API 種別」。
    #   例："大規模言語モデル API（文章要約用）"
    #      "画像生成 API（図解の試作用）"
    #      "文献メタデータ取得 API（DOI 解決）"
    api_type_ja: str
    service_name: str      # 裏付け用のサービス名（デフォルト "Amazon Bedrock (Claude Opus 4.7)"、ユーザー指定があればそちら）。AWS サービスのみ記載可。サードパーティ API は禁止。
    target: str            # 処理対象（B列に入る、例 "科学文献の要約"）
    # 算定パラメータ
    input_tokens_per_month: float
    output_tokens_per_month: float
    input_usd_per_mtok: float   # 入力 $/1M tokens
    output_usd_per_mtok: float  # 出力 $/1M tokens
    months: float
    hours_per_month: float = 0.0  # 稼働時間/月（API が時間ベースの場合のみ使う、0 なら出力しない）
    region: str = "ap-northeast-1"
    fx_jpy_per_usd: float = 150.0
    retry_buffer: float = 0.0
    price_source: str = ""
    notes: str = ""

    @property
    def total_jpy(self) -> float:
        in_usd = self.input_tokens_per_month / 1_000_000 * self.input_usd_per_mtok
        out_usd = self.output_tokens_per_month / 1_000_000 * self.output_usd_per_mtok
        monthly_usd = in_usd + out_usd
        total_usd = monthly_usd * self.months * (1 + self.retry_buffer)
        return total_usd * self.fx_jpy_per_usd

    @property
    def total_thousand_jpy(self) -> int:
        """千円単位（xlsx 書き込み用）"""
        return math.floor(self.total_jpy / 1000)

    @property
    def total_man_jpy(self) -> float:
        """万円単位、小数1桁（表示用）"""
        return round(self.total_jpy / 10000, 1)

    @property
    def calc_breakdown(self) -> str:
        monthly_usd = (
            self.input_tokens_per_month / 1_000_000 * self.input_usd_per_mtok
            + self.output_tokens_per_month / 1_000_000 * self.output_usd_per_mtok
        )
        price_src = self.price_source or "AWS公式料金ページ参照"
        lines = [
            f"[API種類] {self.api_type_ja}",
            f"[サービス] {self.service_name}",
            f"[リージョン] {self.region}",
            f"[単価] 入力 ${self.input_usd_per_mtok}/1M tokens、"
            f"出力 ${self.output_usd_per_mtok}/1M tokens ({price_src})",
            f"[想定量] 入力 {self.input_tokens_per_month:,.0f} tok/月、"
            f"出力 {self.output_tokens_per_month:,.0f} tok/月 × {self.months:g} ヶ月",
        ]
        if self.hours_per_month > 0:
            lines.append(f"[稼働時間] {self.hours_per_month:g}時間/月 × {self.months:g}ヶ月")
        if self.retry_buffer > 0:
            lines.append(f"[再実行率] {self.retry_buffer:.0%} バッファ")
        lines.append(
            f"[計算] 月額 ${monthly_usd:.4f} × {self.months:g}ヶ月 × "
            f"{self.fx_jpy_per_usd:.0f}円/USD"
            + (f" × (1+{self.retry_buffer:.0%})" if self.retry_buffer > 0 else "")
            + f" = {self.total_man_jpy:,.1f}万円"
        )
        if self.notes:
            lines.append(f"[備考] {self.notes}")
        return "\n".join(lines)


def summarize(computes: List[ComputeCost], apis: List[ApiCost]) -> dict:
    """4枚目テーブルに書き込む辞書を生成する"""
    compute_rows = []
    for c in computes:
        compute_rows.append({
            "service_name": c.service_name,
            # xlsx 4枚目 B列(GPU種類) にはサービス名＋GPU型番を書く
            "gpu_type": f"{c.service_name} / {c.gpu_type}",
            "rationale": c.rationale,
            "amount_thousand_jpy": c.total_thousand_jpy,  # xlsx書き込み用
            "amount_man_jpy": c.total_man_jpy,             # 表示用
            "basis": c.calc_breakdown,
        })
    api_rows = []
    for a in apis:
        api_rows.append({
            "api_type_ja": a.api_type_ja,
            "service_name": a.service_name,
            # xlsx 4枚目 B列(処理対象) にはサービス名を書く。
            # 日本語のAPI種類は算定根拠側（D列）の [API種類] 行で表現。
            "target": a.service_name,
            "amount_thousand_jpy": a.total_thousand_jpy,
            "amount_man_jpy": a.total_man_jpy,
            "basis": a.calc_breakdown,
        })
    total_compute = sum(r["amount_thousand_jpy"] for r in compute_rows)
    total_api = sum(r["amount_thousand_jpy"] for r in api_rows)
    return {
        "compute_rows": compute_rows,
        "api_rows": api_rows,
        "total_compute_thousand_jpy": total_compute,
        "total_api_thousand_jpy": total_api,
        "grand_total_thousand_jpy": total_compute + total_api,
        "total_compute_man_jpy": round(total_compute / 10, 1),
        "total_api_man_jpy": round(total_api / 10, 1),
        "grand_total_man_jpy": round((total_compute + total_api) / 10, 1),
    }


def render_markdown(summary: dict) -> str:
    """内部メモ用に Markdown 形式で出力（人が読む金額は万円単位）

    注：xlsx 様式の「金額(千円)」欄に書き込む値は `amount_thousand_jpy`。
    """
    lines = ["# AWS 費用算定内訳", ""]
    lines.append("※ 表示は万円単位。xlsx 書き込み時は千円単位に自動換算。")
    lines.append("")
    lines.append("## API費用（Bedrock/その他）")
    lines.append("")
    lines.append("| # | 処理対象 | 金額(万円) | 算定根拠 |")
    lines.append("|---|---|---|---|")
    for i, r in enumerate(summary["api_rows"], 1):
        basis = r["basis"].replace("\n", " / ")
        lines.append(
            f"| {i} | {r['target']} | {r['amount_man_jpy']:,.1f} | {basis} |"
        )
    lines.append(
        f"\n**API費用 小計：{summary['total_api_man_jpy']:,.1f} 万円** "
        f"（xlsx 書込値：{summary['total_api_thousand_jpy']:,}千円）"
    )
    lines.append("")
    lines.append("## 計算資源費用")
    lines.append("")
    lines.append("| # | GPU種類 | 選定理由 | 金額(万円) | 算定根拠 |")
    lines.append("|---|---|---|---|---|")
    for i, r in enumerate(summary["compute_rows"], 1):
        basis = r["basis"].replace("\n", " / ")
        lines.append(
            f"| {i} | {r['gpu_type']} | {r['rationale']} | "
            f"{r['amount_man_jpy']:,.1f} | {basis} |"
        )
    lines.append(
        f"\n**計算資源費用 小計：{summary['total_compute_man_jpy']:,.1f} 万円** "
        f"（xlsx 書込値：{summary['total_compute_thousand_jpy']:,}千円）"
    )
    lines.append("")
    lines.append(
        f"## 総計：{summary['grand_total_man_jpy']:,.1f} 万円 "
        f"（xlsx 書込値：{summary['grand_total_thousand_jpy']:,}千円）"
    )
    lines.append("（3枚目「その他」合計と一致させること。上限500万円。）")
    return "\n".join(lines)


if __name__ == "__main__":
    # 動作確認用
    compute = [
        ComputeCost(
            service_name="Amazon EC2 p4de.24xlarge",
            gpu_type="NVIDIA A100 80GB × 8",
            rationale="70B LLM のフルFTに80GB VRAM×8が必要。H100(p5)より安価なA100で要件を満たす",
            hourly_usd=40.96,
            hours_per_month=80,
            months=6,
            region="ap-northeast-1",
            fx_jpy_per_usd=150,
            retry_buffer=0.20,
            price_source="AWS Price List Bulk API, 2026-04-15取得",
            notes="学習データ 10GB × 3 エポック",
        ),
    ]
    apis = [
        ApiCost(
            api_type_ja="大規模言語モデルAPI（文献要約・情報抽出用）",
            service_name="Amazon Bedrock (Claude Opus 4.7)",
            target="科学文献の要約と情報抽出",
            input_tokens_per_month=500_000,
            output_tokens_per_month=100_000,
            input_usd_per_mtok=3.0,
            output_usd_per_mtok=15.0,
            months=6,
            region="ap-northeast-1",
            fx_jpy_per_usd=150,
            retry_buffer=0.20,
            price_source="AWS Price List Bulk API, 2026-04-15取得",
            notes="PubMed 論文のメタデータ抽出",
        ),
    ]
    s = summarize(compute, apis)
    print(render_markdown(s))
