"""AWS 公開料金情報取得ヘルパー

AWS Price List Bulk API（認証不要・公開）にアクセスし、
**必要なサービス・リージョン・SKUだけ**を取得する。

アクセスフロー：

1. https://pricing.us-east-1.amazonaws.com/offers/v1.0/aws/index.json
   を一度取得（小さいメタデータ、ローカルにキャッシュ）
2. 対象サービスの currentRegionIndexUrl から region_index.json を取得
3. 対象リージョンの currentVersionUrl を特定
4. regional 本体 JSON を **HTTP ストリーム** で取得し、ijson で
   products を1件ずつパース。対象 SKU を見つけた時点で接続を閉じる
   （数百MB の JSON を全ダウンロードしない）
5. 取得した単価は結果キャッシュ（小さな JSON）に保存。2回目以降は本体に触らない

使い方（CLI）::

    python3 fetch_aws_price.py ec2 p4de.24xlarge ap-northeast-1
    python3 fetch_aws_price.py bedrock anthropic.claude-sonnet-4-20250514 us-east-1 --io output
    python3 fetch_aws_price.py s3 standard ap-northeast-1

ライブラリとして::

    from fetch_aws_price import get_ec2_ondemand_usd_hour
    usd_hour = get_ec2_ondemand_usd_hour("p4de.24xlarge", "ap-northeast-1")

依存： `ijson`（pip install ijson）

キャッシュ先はデフォルト `/tmp/aws_price_cache/`、環境変数
`AWS_PRICE_CACHE_DIR` で上書き可。
"""

from __future__ import annotations
import argparse
import json
import os
import sys
import time
import urllib.request
from pathlib import Path
from typing import Optional

try:
    import ijson
except ImportError:
    print(
        "ERROR: ijson required. pip install ijson --break-system-packages",
        file=sys.stderr,
    )
    raise

BASE = "https://pricing.us-east-1.amazonaws.com"
INDEX = f"{BASE}/offers/v1.0/aws/index.json"
CACHE_DIR = Path(os.environ.get("AWS_PRICE_CACHE_DIR", "/tmp/aws_price_cache"))
META_TTL = 24 * 60 * 60       # metadata 1 day
RESULT_TTL = 7 * 24 * 60 * 60  # SKU 単価結果 1週間（変わりにくい）


# ------------- metadata (small) -------------


def _fetch_json(url: str, cache_name: str, ttl: int = META_TTL) -> dict:
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    p = CACHE_DIR / cache_name
    if p.exists() and time.time() - p.stat().st_mtime < ttl:
        try:
            with open(p, "r", encoding="utf-8") as f:
                return json.load(f)
        except json.JSONDecodeError:
            p.unlink(missing_ok=True)
    with urllib.request.urlopen(url, timeout=60) as r:
        data = json.loads(r.read().decode("utf-8"))
    with open(p, "w", encoding="utf-8") as f:
        json.dump(data, f)
    return data


def _service_region_url(service_code: str, region: str) -> str:
    idx = _fetch_json(INDEX, "index.json")
    svc = idx["offers"].get(service_code)
    if not svc:
        raise ValueError(f"service {service_code!r} not in index")
    region_index = _fetch_json(
        BASE + svc["currentRegionIndexUrl"],
        f"{service_code}_region_index.json",
    )
    r = region_index["regions"].get(region)
    if not r:
        raise ValueError(f"region {region!r} not in {service_code}")
    return BASE + r["currentVersionUrl"]


# ------------- streaming (large) -------------


def _stream_products(url: str):
    """Yield (sku, product_dict) by streaming HTTP, stop when caller breaks."""
    req = urllib.request.Request(url)
    resp = urllib.request.urlopen(req, timeout=300)
    try:
        for sku, product in ijson.kvitems(resp, "products"):
            yield sku, product
    finally:
        resp.close()


def _stream_ondemand_price_for_sku(url: str, sku: str) -> Optional[dict]:
    """Stream terms.OnDemand and return the first dimension dict for the SKU."""
    req = urllib.request.Request(url)
    resp = urllib.request.urlopen(req, timeout=300)
    try:
        for tsku, term in ijson.kvitems(resp, "terms.OnDemand"):
            if tsku != sku:
                continue
            # term is dict of offerTermCode -> {...}
            for _, tbody in term.items():
                for _, pd in tbody.get("priceDimensions", {}).items():
                    return pd
            return None
    finally:
        resp.close()
    return None


# ------------- result cache -------------


def _result_cache(key: str) -> Optional[float]:
    p = CACHE_DIR / "results" / f"{key}.json"
    if p.exists() and time.time() - p.stat().st_mtime < RESULT_TTL:
        with open(p) as f:
            return json.load(f).get("price")
    return None


def _result_save(key: str, price: float, meta: dict):
    d = CACHE_DIR / "results"
    d.mkdir(parents=True, exist_ok=True)
    with open(d / f"{key}.json", "w") as f:
        json.dump({"price": price, **meta, "fetched": int(time.time())}, f)


# ------------- price queries -------------


def get_ec2_ondemand_usd_hour(
    instance_type: str,
    region: str,
    os_: str = "Linux",
    tenancy: str = "Shared",
    preinstalled_sw: str = "NA",
) -> Optional[float]:
    """EC2 On-Demand USD/hour (Linux/Shared/no SW/Used capacity)."""
    key = f"ec2_{region}_{instance_type}_{os_}_{tenancy}_{preinstalled_sw}"
    cached = _result_cache(key)
    if cached is not None:
        return cached

    url = _service_region_url("AmazonEC2", region)
    # pass 1: locate SKU
    sku = None
    for psku, p in _stream_products(url):
        a = p.get("attributes") or {}
        if (
            a.get("instanceType") == instance_type
            and a.get("operatingSystem") == os_
            and a.get("tenancy") == tenancy
            and a.get("preInstalledSw") == preinstalled_sw
            and a.get("capacitystatus") == "Used"
        ):
            sku = psku
            break
    if not sku:
        return None
    # pass 2: price for SKU
    pd = _stream_ondemand_price_for_sku(url, sku)
    if not pd:
        return None
    price = float(pd["pricePerUnit"]["USD"])
    _result_save(key, price, {"sku": sku, "unit": pd.get("unit")})
    return price


def get_bedrock_token_usd_per_mtok(
    model_id: str, region: str, io: str = "input"
) -> Optional[float]:
    """Bedrock USD per 1M tokens (input/output)."""
    key = f"bedrock_{region}_{model_id}_{io}"
    cached = _result_cache(key)
    if cached is not None:
        return cached

    url = _service_region_url("AmazonBedrock", region)
    target = "input-tokens" if io == "input" else "output-tokens"
    sku = None
    unit_found = ""
    for psku, p in _stream_products(url):
        a = p.get("attributes") or {}
        model = str(a.get("model") or a.get("modelId") or "")
        if model_id.lower() not in model.lower():
            continue
        usage = str(a.get("operation") or a.get("usagetype") or "").lower()
        if target in usage:
            sku = psku
            break
    if not sku:
        return None
    pd = _stream_ondemand_price_for_sku(url, sku)
    if not pd:
        return None
    raw = float(pd["pricePerUnit"]["USD"])
    unit = str(pd.get("unit", "")).lower()
    # normalize to $/1M tokens
    if "1k" in unit or "1,000 tokens" in unit:
        price = raw * 1000
    elif "1m" in unit or "1,000,000 tokens" in unit:
        price = raw
    else:
        price = raw  # best-effort
    _result_save(
        key, price, {"sku": sku, "unit": pd.get("unit"), "raw": raw}
    )
    return price


def get_s3_standard_usd_gb_month(region: str) -> Optional[float]:
    key = f"s3_standard_{region}"
    cached = _result_cache(key)
    if cached is not None:
        return cached

    url = _service_region_url("AmazonS3", region)
    sku = None
    for psku, p in _stream_products(url):
        a = p.get("attributes") or {}
        if (
            a.get("storageClass") == "General Purpose"
            and a.get("volumeType") == "Standard"
        ):
            sku = psku
            break
    if not sku:
        return None
    pd = _stream_ondemand_price_for_sku(url, sku)
    if not pd:
        return None
    price = float(pd["pricePerUnit"]["USD"])
    _result_save(key, price, {"sku": sku, "unit": pd.get("unit")})
    return price


# ------------- CLI -------------


def main():
    p = argparse.ArgumentParser()
    p.add_argument("kind", choices=["ec2", "bedrock", "s3"])
    p.add_argument("identifier")
    p.add_argument("region", nargs="?", default="ap-northeast-1")
    p.add_argument("--io", choices=["input", "output"], default="input")
    args = p.parse_args()

    if args.kind == "ec2":
        v = get_ec2_ondemand_usd_hour(args.identifier, args.region)
        if v is None:
            print(f"ERROR: {args.identifier} in {args.region} not found")
            sys.exit(1)
        print(f"{v:.4f} USD/hour ({args.identifier}, {args.region}, Linux/Shared)")
    elif args.kind == "bedrock":
        v = get_bedrock_token_usd_per_mtok(args.identifier, args.region, io=args.io)
        if v is None:
            print(f"ERROR: {args.identifier} ({args.io}) in {args.region} not found")
            sys.exit(1)
        print(
            f"{v:.4f} USD per 1M tokens ({args.identifier}, {args.io}, {args.region})"
        )
    elif args.kind == "s3":
        v = get_s3_standard_usd_gb_month(args.region)
        if v is None:
            print(f"ERROR: S3 Standard in {args.region} not found")
            sys.exit(1)
        print(f"{v:.6f} USD per GB-month (S3 Standard, {args.region})")


if __name__ == "__main__":
    main()
