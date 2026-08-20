#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
京华云采全量商品索引构建

用法：
    python build_sku_index.py [--out sku_index.json] [--refresh]

作用：
    遍历台/笔/服/印全品类全部商品，构建 {skuId: {skuName, brandCh, brandEn, cid}} 索引，
    供经销商画像（dealer_profile.py）反查商品名/品牌时直接匹配，无需再遍历商品。

缓存：
    - 已存在索引则增量更新（只抓新 skuId，快）
    - --refresh 强制全量重抓
"""
import argparse
import json
import os
import time

import requests

BASE = 'https://mkt-bjzc.zhongcy.com'
SEARCH_URL = BASE + '/proxy/trade-service/mall/search/querySkuListFromEs'
PLATFORM_ID = 20
PUBLISH_TYPE = 10024

CATEGORIES = {
    "台": 1000011,
    "笔": 1000014,
    "服": 1000128,
    "印": [1000023, 1000025, 1000027, 1000029, 1000041],
}

SESSION = requests.Session()
SESSION.trust_env = False
SESSION.headers.update({
    "Content-Type": "application/json",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Referer": BASE + "/mall-view/product/search",
    "Origin": BASE,
})


def get_products(cid):
    """获取某 cid 下所有商品（含 skuId/skuName/brand）。cid=None 时不带 cid 全量抓。"""
    products = []
    page = 1
    page_size = 100
    while True:
        payload = {
            "queryPage": {"platformId": PLATFORM_ID, "pageSize": page_size, "pageNum": page},
            "homeType": 110, "isAggregation": False, "isHometypes": True,
            "businessType": 1, "publishType": PUBLISH_TYPE,
        }
        if cid is not None:
            payload["cid"] = cid
        try:
            d = SESSION.post(SEARCH_URL, json=payload, timeout=30).json()
        except Exception as e:
            print(f"  ⚠️ 请求失败[{type(e).__name__}]，重试...")
            time.sleep(1)
            continue
        if not d or not d.get("data"):
            break
        rl = d["data"]["itemList"]["resultList"]
        if not rl:
            break
        for it in rl:
            products.append({
                "skuId": it.get("skuId"),
                "skuName": it.get("skuName", ""),
                "brandCh": it.get("brandNameCh", ""),
                "brandEn": it.get("brandNameEn", ""),
                "cid": it.get("cid", cid),
            })
        total = int(d["data"].get("totalCount", 0) or 0)
        if page * page_size >= total:
            break
        page += 1
        time.sleep(0.2)
    return products


def get_cid_list():
    cids = []
    for k, v in CATEGORIES.items():
        if isinstance(v, list):
            cids.extend(v)
        else:
            cids.append(v)
    return cids


def main():
    parser = argparse.ArgumentParser(description="京华云采全量商品索引构建")
    parser.add_argument("--out", default="sku_index.json", help="索引输出路径")
    parser.add_argument("--refresh", action="store_true", help="强制全量重抓（默认增量）")
    parser.add_argument("--all", action="store_true", help="全品类模式：不带 cid 抓全部商品（含耗材，约1万条）")
    args = parser.parse_args()

    index = {}
    if os.path.exists(args.out) and not args.refresh:
        with open(args.out, "r", encoding="utf-8") as f:
            index = json.load(f)
        print(f"复用已有索引 {len(index)} 条，增量更新")

    total_new = 0
    if args.all:
        # 全品类模式：一次遍历全部商品（约 1 万条，几分钟）
        products = get_products(None)
        new = [p for p in products if str(p["skuId"]) not in index]
        for p in new:
            index[str(p["skuId"])] = {
                "skuName": p["skuName"], "brandCh": p["brandCh"],
                "brandEn": p["brandEn"], "cid": p["cid"],
            }
        total_new = len(new)
        print(f"全品类：{len(products)} 个商品，新增 {len(new)} 条")
    else:
        for cid in get_cid_list():
            products = get_products(cid)
            new = [p for p in products if str(p["skuId"]) not in index]
            for p in new:
                index[str(p["skuId"])] = {
                    "skuName": p["skuName"], "brandCh": p["brandCh"],
                    "brandEn": p["brandEn"], "cid": p["cid"],
                }
            total_new += len(new)
            print(f"cid={cid}：{len(products)} 个商品，新增 {len(new)} 条")

    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(index, f, ensure_ascii=False, indent=0)
    print(f"\n索引构建完成：共 {len(index)} 个商品（本次新增 {total_new}）")
    print(f"已保存：{os.path.abspath(args.out)}")

    # 品牌统计
    from collections import Counter
    brands = Counter((v.get("brandCh") or "") + (v.get("brandEn") or "") for v in index.values())
    print("\n品牌分布 TOP10：")
    for b, c in brands.most_common(10):
        if b:
            print(f"  {b}: {c} 个商品")


if __name__ == "__main__":
    main()
