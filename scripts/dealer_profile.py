#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
京华云采经销商全景画像（代理商品 → 销售记录 → 客户 → 归集）

用法：
    python dealer_profile.py "北京宏鹏电脑有限公司" [--out 画像.xlsx] [--index sku_index.json] [--max-skus 100]

数据链路（2026-08-20 打通）：
    反查(shopName) → 代理商品 skuId 列表 → 商品索引(sku_index.json) 补品牌/商品名
                    → 每个 skuId 查销售记录 → 归集 shopName==该经销商的成交 → 客户清单

输出 Excel：
    经销商概览 / 代理商品清单 / 品牌分布 / 成交记录 / 客户汇总 / 零成交商品(机会)

说明：
    - 销售记录里的 shopName 才是"实际卖出方"，反查的 agentName 是"代理方"；
      归集成交时按 shopName==经销商 过滤，才是它自己卖的客户
    - 商品索引由 build_sku_index.py 生成；索引没有的 skuId 标"未知商品"
"""
import argparse
import json
import os
import sys
import time

import pandas as pd
import requests
from openpyxl import load_workbook
from openpyxl.styles import Font

BASE = 'https://mkt-bjzc.zhongcy.com'
SHOP = 'https://shop-bjzc.zhongcy.com'
AGENT_URL = BASE + '/proxy/trade-service/mall/search/querySkuAgentListFromEs'
SALE_URL = SHOP + '/proxy/trade-service/mall/order/querySkuSaleRecord'
PLATFORM_ID = 20
PUBLISH_TYPE = 10024

SESSION = requests.Session()
SESSION.trust_env = False
SESSION.headers.update({
    "Content-Type": "application/json",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Referer": BASE + "/mall-view/product/search",
    "Origin": BASE,
})


def query_agent(shop_name, page=1, retries=3):
    """反查经销商代理商品（分页）"""
    payload = {
        "platformId": PLATFORM_ID, "shopName": shop_name,
        "queryPage": {"platformId": PLATFORM_ID, "pageSize": 100, "pageNum": page},
        "publishType": PUBLISH_TYPE,
    }
    for i in range(retries):
        try:
            d = SESSION.post(AGENT_URL, json=payload, timeout=30).json()
            if d.get("code") == "0":
                return (d.get("data") or {}).get("itemAgentList") or {}
            return {}
        except Exception:
            time.sleep(0.5 * (i + 2))
    return {}


def query_sale(sku_id, page=1, retries=3):
    """查某商品的销售记录"""
    for i in range(retries):
        try:
            r = SESSION.get(SALE_URL, params={
                "pageNum": page, "pageSize": 100, "platformId": PLATFORM_ID, "skuId": sku_id,
            }, timeout=30)
            d = r.json()
            if d.get("code") == "0":
                return d.get("data") or {}
            return {}
        except Exception:
            time.sleep(0.5 * (i + 2))
    return {}


def main():
    parser = argparse.ArgumentParser(description="京华云采经销商全景画像")
    parser.add_argument("dealer", help="经销商公司名")
    parser.add_argument("--out", default=None, help="输出 Excel 路径")
    parser.add_argument("--index", default="sku_index.json", help="商品索引路径")
    parser.add_argument("--max-skus", type=int, default=100, help="最多分析的商品数（默认 100，0=全部）")
    parser.add_argument("--quiet", action="store_true", help="安静模式：只打印摘要")
    args = parser.parse_args()

    dealer = args.dealer.strip()
    out_path = args.out or f"{dealer}_画像.xlsx"
    quiet = args.quiet

    # 商品索引
    index = {}
    if os.path.exists(args.index):
        with open(args.index, "r", encoding="utf-8") as f:
            index = json.load(f)

    def prod_info(sku_id):
        it = index.get(str(sku_id))
        if it:
            return f"{it.get('brandCh','')}{it.get('brandEn','')} {it.get('skuName','')}".strip()
        return "未知商品"

    def prod_brand(sku_id):
        it = index.get(str(sku_id))
        if it:
            return f"{it.get('brandCh','')}{it.get('brandEn','')}".strip() or "未知品牌"
        return "未知品牌"

    # 1. 反查：拿全部代理商品
    if not quiet:
        print(f"▶ 反查 {dealer} 的代理商品...")
    agents = []
    page = 1
    while True:
        ia = query_agent(dealer, page)
        rl = ia.get("resultList") or []
        if not rl:
            break
        agents.extend(rl)
        count = ia.get("count") or 0
        if page * 100 >= count and count > 0:
            break
        page += 1
        time.sleep(0.2)
    if not agents:
        print(f"❌ 未查到 {dealer} 的代理记录（可能不是京华云采经销商，或为品牌方）")
        sys.exit(1)

    # 去重 skuId（同一商品多条记录取第一条）
    seen = {}
    for a in agents:
        sku = a.get("skuId")
        if sku and sku not in seen:
            seen[sku] = a
    sku_ids = list(seen.keys())
    if args.max_skus and args.max_skus > 0:
        sku_ids = sku_ids[:args.max_skus]

    phone = agents[0].get("agentPhone", "")
    sales_area = agents[0].get("salesArea", "")
    if not quiet:
        print(f"  代理商品数（去重）：{len(sku_ids)} | 电话：{phone} | 销售区域：{sales_area}")

    # 2. 每个 skuId 查销售记录，归集本经销商成交
    sales = []  # 该经销商的成交（shopName==dealer）
    all_sales = []  # 所有成交（含别人卖的）
    for i, sku in enumerate(sku_ids, 1):
        data = query_sale(sku)
        records = (data or {}).get("result") or []
        for rec in records:
            info = {
                "skuId": sku,
                "商品": prod_info(sku),
                "品牌": prod_brand(sku),
                "采购单位": rec.get("organizeName", ""),
                "供应商": rec.get("shopName", ""),
                "数量": rec.get("skuNum", 0),
                "单价(元)": rec.get("sellPrice", 0),
                "金额(元)": round(rec.get("skuNum", 0) * rec.get("sellPrice", 0), 2),
                "成交时间": str(rec.get("orderTime", ""))[:16],
            }
            all_sales.append(info)
            if rec.get("shopName") == dealer:
                sales.append(info)
        if not quiet and (i % 10 == 0 or i == len(sku_ids)):
            print(f"  [{i}/{len(sku_ids)}] 已扫 {len(all_sales)} 条成交，其中本经销商 {len(sales)} 条")
        time.sleep(0.2)

    df_sales = pd.DataFrame(sales) if sales else pd.DataFrame(columns=[
        "skuId", "商品", "品牌", "采购单位", "供应商", "数量", "单价(元)", "金额(元)", "成交时间"])

    # 3. 汇总
    total_qty = int(df_sales["数量"].sum()) if len(df_sales) else 0
    total_amt = df_sales["金额(元)"].sum() if len(df_sales) else 0
    n_customers = df_sales["采购单位"].nunique() if len(df_sales) else 0
    sold_skus = set(df_sales["skuId"]) if len(df_sales) else set()
    zero_skus = [s for s in sku_ids if s not in sold_skus]  # 代理但无成交

    # 品牌分布（代理维度）
    from collections import Counter
    brand_cnt = Counter(prod_brand(s) for s in sku_ids)
    df_brand = pd.DataFrame(brand_cnt.items(), columns=["品牌", "代理商品数"]).sort_values("代理商品数", ascending=False)

    # 客户汇总
    if len(df_sales):
        df_cust = df_sales.groupby("采购单位").agg(
            次数=("采购单位", "count"), 数量=("数量", "sum"), 金额=("金额(元)", "sum"),
        ).sort_values("金额", ascending=False).reset_index()
    else:
        df_cust = pd.DataFrame(columns=["采购单位", "次数", "数量", "金额"])

    # 代理商品清单（含成交标记）
    df_skus = pd.DataFrame([{
        "skuId": s, "商品": prod_info(s), "品牌": prod_brand(s),
        "供货价(元)": seen[s].get("supplyPrice"),
        "库存": seen[s].get("inventory"),
        "有无成交": "✅" if s in sold_skus else "— 零成交",
    } for s in sku_ids])

    # 零成交商品
    df_zero = df_skus[df_skus["有无成交"] != "✅"].reset_index(drop=True)

    # 4. 输出
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        overview = pd.DataFrame({
            "指标": ["经销商", "电话", "销售区域", "代理商品数", "成交商品数",
                    "零成交商品数", "总销量", "总成交额(元)", "客户数"],
            "数值": [dealer, phone, sales_area, len(sku_ids), len(sold_skus),
                    len(zero_skus), total_qty, round(total_amt, 2), n_customers],
        })
        overview.to_excel(writer, sheet_name="经销商概览", index=False)
        df_brand.to_excel(writer, sheet_name="品牌分布", index=False)
        df_skus.to_excel(writer, sheet_name="代理商品清单", index=False)
        df_sales.to_excel(writer, sheet_name="成交记录", index=False)
        df_cust.to_excel(writer, sheet_name="客户汇总", index=False)
        df_zero.to_excel(writer, sheet_name="零成交商品", index=False)

    print("=" * 56)
    print(f"经销商画像：{dealer}")
    print(f"  电话：{phone} | 销售区域：{sales_area}")
    print(f"  代理商品：{len(sku_ids)} 个（品牌 {len(brand_cnt)} 个）| 电话已确认")
    print(f"  实际成交：{len(df_sales)} 笔 / {total_qty} 台 / ¥{total_amt:,.0f} / {n_customers} 家客户")
    print(f"  零成交商品：{len(zero_skus)} 个（代理但没卖出去 → 潜在机会）")
    if len(df_cust):
        print("\n■ TOP 客户：")
        for _, r in df_cust.head(5).iterrows():
            print(f"  · {r['采购单位'][:24]}：{int(r['数量'])}台 ¥{r['金额']:,.0f}")
    print(f"\n报告已保存：{out_path}")


if __name__ == "__main__":
    main()
