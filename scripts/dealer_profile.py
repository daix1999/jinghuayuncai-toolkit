#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
京华云采经销商业务状态穿透（代理商品 → 交易记录 → 卖了多少 → 卖给谁 → 占比）

用法：
    python dealer_profile.py "北京宏鹏电脑有限公司" [--out 画像.xlsx] [--index sku_index.json] [--max-skus 100]

数据链路（2026-08-20 打通）：
    反查(shopName) → 代理商品 skuId 列表
                  → 每个 skuId 查销售记录：统计「商品总成交」和「该经销商成交」
                  → 算出该经销商在代理商品盘子里的份额（台数/金额占比）
                  → 客户归集（卖给谁）

核心输出：
    - 业务状态穿透：代理商品总盘子 vs 该经销商成交 → 份额（台数%/金额%）
    - 商品级占比：每个代理商品上，该经销商占多少（代理不卖/主卖/被抢）
    - 成交记录/客户清单/零成交商品

口径：
    - 销售记录 shopName = 实际卖出方；只统计 shopName==该经销商 的成交，才是它自己卖的
    - "代理商品总成交" = 该经销商代理的全部商品在平台上的成交总量（它可参与的总盘子）
"""
import argparse
import json
import os
import sys
import time
from collections import Counter

import pandas as pd
import requests
from openpyxl import load_workbook

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
    parser = argparse.ArgumentParser(description="京华云采经销商业务状态穿透")
    parser.add_argument("dealer", help="经销商公司名")
    parser.add_argument("--out", default=None, help="输出 Excel 路径")
    parser.add_argument("--index", default="sku_index.json", help="商品索引路径")
    parser.add_argument("--max-skus", type=int, default=100, help="最多分析的代理商品数（默认 100，0=全部）")
    parser.add_argument("--quiet", action="store_true", help="安静模式：只打印摘要")
    args = parser.parse_args()

    dealer = args.dealer.strip()
    out_path = args.out or f"{dealer}_业务穿透.xlsx"
    quiet = args.quiet

    index = {}
    if os.path.exists(args.index):
        with open(args.index, "r", encoding="utf-8") as f:
            index = json.load(f)

    def prod_info(sku_id):
        it = index.get(str(sku_id))
        if it:
            return f"{it.get('brandCh','')}{it.get('brandEn','')} {it.get('skuName','')}".strip()
        return "已下架商品"  # 不在索引 = 已彻底下架，平台不再提供名称

    def prod_brand(sku_id):
        it = index.get(str(sku_id))
        if it:
            return f"{it.get('brandCh','')}{it.get('brandEn','')}".strip() or "未知品牌"
        return "已下架"

    # 1. 反查全部代理商品
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

    seen = {}
    for a in agents:
        sku = a.get("skuId")
        if sku and sku not in seen:
            seen[sku] = a
    all_sku_ids = list(seen.keys())

    # 排除已彻底下架商品（不在索引 = 平台已移除、无商品名，不参与统计，降低数据量）
    sku_ids = [s for s in all_sku_ids if str(s) in index]
    n_offline = len(all_sku_ids) - len(sku_ids)
    if args.max_skus and args.max_skus > 0:
        sku_ids = sku_ids[:args.max_skus]

    phone = agents[0].get("agentPhone", "")
    sales_area = agents[0].get("salesArea", "")
    if not quiet:
        print(f"  代理商品数（去重）：{len(sku_ids)} | 电话：{phone} | 销售区域：{sales_area}")

    # 2. 每个 skuId 查销售记录：累计「商品总成交」与「该经销商成交」
    sku_stat = {}      # skuId -> {总销量,总金额,X销量,X金额}
    dealer_sales = []  # 该经销商实际卖出的订单
    for i, sku in enumerate(sku_ids, 1):
        data = query_sale(sku)
        records = (data or {}).get("result") or []
        st = sku_stat.setdefault(sku, {"总销量": 0, "总金额": 0.0, "X销量": 0, "X金额": 0.0})
        for rec in records:
            qty = int(rec.get("skuNum") or 0)
            amt = qty * float(rec.get("sellPrice") or 0)
            st["总销量"] += qty
            st["总金额"] += amt
            if rec.get("shopName") == dealer:
                st["X销量"] += qty
                st["X金额"] += amt
                dealer_sales.append({
                    "skuId": sku, "商品": prod_info(sku), "品牌": prod_brand(sku),
                    "采购单位": rec.get("organizeName", ""),
                    "数量": qty, "单价(元)": rec.get("sellPrice", 0),
                    "金额(元)": round(amt, 2), "成交时间": str(rec.get("orderTime", ""))[:16],
                })
        if not quiet and (i % 10 == 0 or i == len(sku_ids)):
            print(f"  [{i}/{len(sku_ids)}] 已扫 {len(sku_stat)} 个商品，该经销商累计 {len(dealer_sales)} 笔")
        time.sleep(0.2)

    # 3. 汇总：总盘子 vs 该经销商
    total_qty = sum(s["总销量"] for s in sku_stat.values())      # 代理商品总盘子（台）
    total_amt = sum(s["总金额"] for s in sku_stat.values())      # 代理商品总盘子（元）
    x_qty = sum(s["X销量"] for s in sku_stat.values())           # 该经销商卖出（台）
    x_amt = sum(s["X金额"] for s in sku_stat.values())           # 该经销商卖出（元）
    share_qty = x_qty / total_qty * 100 if total_qty else 0
    share_amt = x_amt / total_amt * 100 if total_amt else 0

    sold_skus = {s for s, st in sku_stat.items() if st["X销量"] > 0}
    n_customers = len({r["采购单位"] for r in dealer_sales})

    df_sales = pd.DataFrame(dealer_sales) if dealer_sales else pd.DataFrame(
        columns=["skuId", "商品", "品牌", "采购单位", "数量", "单价(元)", "金额(元)", "成交时间"])

    # 客户汇总
    if len(df_sales):
        df_cust = df_sales.groupby("采购单位").agg(
            次数=("采购单位", "count"), 数量=("数量", "sum"), 金额=("金额(元)", "sum"),
        ).sort_values("金额", ascending=False).reset_index()
    else:
        df_cust = pd.DataFrame(columns=["采购单位", "次数", "数量", "金额"])

    # 商品级占比表（核心：每个代理商品上这家占多少）
    df_share = pd.DataFrame([{
        "skuId": s, "商品": prod_info(s), "品牌": prod_brand(s),
        "商品总销量": st["总销量"], "该经销商销量": st["X销量"],
        "销量占比%": round(st["X销量"] / st["总销量"] * 100, 1) if st["总销量"] else 0,
        "商品总金额": round(st["总金额"], 0), "该经销商金额": round(st["X金额"], 0),
        "金额占比%": round(st["X金额"] / st["总金额"] * 100, 1) if st["总金额"] else 0,
        "角色": ("主卖" if (st["X销量"] > 0 and st["X销量"] / st["总销量"] > 0.5)
                else ("有卖" if st["X销量"] > 0 else "代理未卖")),
    } for s, st in sku_stat.items()]).sort_values("该经销商金额", ascending=False)

    # 品牌分布
    brand_cnt = Counter(prod_brand(s) for s in sku_ids)
    df_brand = pd.DataFrame(brand_cnt.items(), columns=["品牌", "代理商品数"]).sort_values("代理商品数", ascending=False)

    # 4. 输出
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        overview = pd.DataFrame({
            "指标": ["经销商", "电话", "销售区域", "代理商品数(统计)", "已下架商品(已排除)",
                    "有成交商品", "零成交商品",
                    "代理商品总成交(台)", "代理商品总成交(元)",
                    "该经销商成交(台)", "该经销商成交(元)",
                    "台数份额%", "金额份额%", "客户数"],
            "数值": [dealer, phone, sales_area, len(sku_ids), n_offline,
                    len(sold_skus), len(sku_ids) - len(sold_skus),
                    total_qty, round(total_amt, 2), x_qty, round(x_amt, 2),
                    round(share_qty, 1), round(share_amt, 1), n_customers],
        })
        overview.to_excel(writer, sheet_name="业务状态穿透", index=False)
        df_share.to_excel(writer, sheet_name="商品成交占比", index=False)
        df_brand.to_excel(writer, sheet_name="品牌分布", index=False)
        df_sales.to_excel(writer, sheet_name="成交记录", index=False)
        df_cust.to_excel(writer, sheet_name="客户汇总", index=False)

    # 控制台摘要
    print("=" * 60)
    print(f"业务穿透：{dealer}")
    print(f"  电话：{phone} | 销售区域：{sales_area}")
    print(f"  代理商品：共 {len(sku_ids) + n_offline} 个（已排除彻底下架 {n_offline} 个，统计 {len(sku_ids)} 个）")
    if len(sku_ids) == 0:
        print("  ⚠️ 该经销商代理的商品均已彻底下架，无在售数据可统计")
        print(f"\n报告已保存：{out_path}")
        return
    print(f"  其中实际卖出 {len(sold_skus)} 个（{len(sold_skus)/len(sku_ids)*100:.0f}%）")
    print(f"  📦 代理商品总盘子：{total_qty} 台 / ¥{total_amt:,.0f}")
    print(f"  💰 该经销商卖出：{x_qty} 台 / ¥{x_amt:,.0f}（占盘子 {share_qty:.1f}% 台 / {share_amt:.1f}% 金额）")
    print(f"  客户：{n_customers} 家")
    if len(df_cust):
        print("\n■ TOP 客户：")
        for _, r in df_cust.head(5).iterrows():
            print(f"  · {r['采购单位'][:24]}：{int(r['数量'])}台 ¥{r['金额']:,.0f}")
    main_roles = df_share["角色"].value_counts()
    print("\n■ 商品角色分布：", dict(main_roles))
    print(f"\n报告已保存：{out_path}")


if __name__ == "__main__":
    main()
