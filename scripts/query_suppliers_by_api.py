#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
京华云采供应商电话 API 查询脚本
通过抓包逆向的 API 端点，遍历各品类商品的经销商列表，
按 agentName 精确匹配名单中的供应商，返回其 agentPhone。

用法：
  python query_suppliers_by_api.py <供应商名单.json> [--categories 台,笔,服,印] [--out 结果.json]

供应商名单.json 格式：["公司A", "公司B", ...]（公司名列表）

API 端点（抓包所得）：
  - 商品搜索: POST {BASE}/proxy/trade-service/mall/search/querySkuListFromEs
  - 经销商列表: POST {BASE}/proxy/trade-service/mall/search/querySkuAgentListFromEs
    （直接返回 agentName + agentPhone，无需点击弹窗）
"""
import argparse
import json
import os
import time
import requests

BASE = "https://mkt-bjzc.zhongcy.com"
SEARCH_URL = BASE + "/proxy/trade-service/mall/search/querySkuListFromEs"
AGENT_URL = BASE + "/proxy/trade-service/mall/search/querySkuAgentListFromEs"
PLATFORM_ID = 20
PUBLISH_TYPE = 10024

# 品类 cid 对照表
CATEGORIES = {
    "台": {"cid": 1000011, "name": "台式计算机"},
    "笔": {"cid": 1000014, "name": "便携式计算机"},
    "服": {"cid": 1000128, "name": "服务器"},
    "印": {"cid": 1000023, "name": "打印机"},  # 打印机还有 25/27/29/41 等子类
}

# 打印机扩展 cid（喷墨/激光/针式/多功能/复印等）
PRINTER_CIDS = [1000023, 1000025, 1000027, 1000029, 1000041]

SESSION = requests.Session()
# 绕过系统代理直连（避免代理导致 zhongcy.com TLS 握手失败）
SESSION.trust_env = False
SESSION.headers.update({
    "Content-Type": "application/json",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Referer": BASE + "/mall-view/product/search",
    "Origin": BASE,
})


def post_json(url, payload, retries=3):
    """POST JSON，带重试。返回 dict 或 None"""
    for i in range(retries):
        try:
            r = SESSION.post(url, json=payload, timeout=30)
            d = r.json()
            if d.get("code") == "0":
                return d
            # 系统异常码可重试
            if "99999" in str(d.get("code", "")):
                time.sleep(0.5 * (i + 2))
                continue
            return d
        except Exception as e:
            time.sleep(0.5 * (i + 2))
    return None


def get_cid_list(category_keys):
    """根据品类键返回要遍历的 cid 列表"""
    cids = []
    for k in category_keys:
        if k == "印":
            cids.extend(PRINTER_CIDS)
        elif k in CATEGORIES:
            cids.append(CATEGORIES[k]["cid"])
    # 去重保序
    seen = set()
    result = []
    for c in cids:
        if c not in seen:
            seen.add(c)
            result.append(c)
    return result


def get_products(cid):
    """获取某 cid 下所有商品 skuId 列表"""
    products = []
    page = 1
    page_size = 100
    while True:
        payload = {
            "queryPage": {"platformId": PLATFORM_ID, "pageSize": page_size, "pageNum": page},
            "homeType": 110,
            "isAggregation": False,
            "isHometypes": True,
            "cid": cid,
            "businessType": 1,
            "publishType": PUBLISH_TYPE,
        }
        d = post_json(SEARCH_URL, payload)
        if not d or not d.get("data"):
            break
        rl = d["data"]["itemList"]["resultList"]
        if not rl:
            break
        for it in rl:
            products.append({
                "skuId": it["skuId"],
                "skuName": it.get("skuName", ""),
                "brandName": (it.get("brandNameCh") or "") + (it.get("brandNameEn") or ""),
            })
        total = int(d["data"].get("totalCount", 0) or 0)
        if page * page_size >= total:
            break
        page += 1
        time.sleep(0.2)
    return products


def get_dealers(sku_id):
    """获取某商品的全部经销商（含 agentPhone）"""
    dealers = []
    page = 1
    page_size = 100  # 上限 100，超限触发服务端排序异常
    while True:
        payload = {
            "platformId": PLATFORM_ID,
            "skuId": sku_id,
            "shopName": "",
            "queryPage": {"platformId": PLATFORM_ID, "pageSize": page_size, "pageNum": page},
            "publishType": PUBLISH_TYPE,
        }
        d = post_json(AGENT_URL, payload)
        if not d or not d.get("data"):
            break
        ia = d["data"]["itemAgentList"]
        rl = ia.get("resultList", [])
        dealers.extend(rl)
        total = ia.get("count", 0)
        if page * page_size >= total or not rl:
            break
        page += 1
        time.sleep(0.25)
    return dealers


def main():
    parser = argparse.ArgumentParser(description="京华云采供应商电话 API 查询")
    parser.add_argument("supplier_list", help="供应商名单 JSON：公司名列表")
    parser.add_argument("--categories", default="台,笔,服,印", help="品类键（台/笔/服/印），逗号分隔")
    parser.add_argument("--out", default=None, help="结果 JSON 输出路径")
    args = parser.parse_args()

    with open(args.supplier_list, "r", encoding="utf-8") as f:
        suppliers = json.load(f)
    supplier_set = set(str(s).strip() for s in suppliers if str(s).strip())
    print(f"待匹配供应商：{len(supplier_set)} 家")

    cids = get_cid_list([k.strip() for k in args.categories.split(",") if k.strip()])
    print(f"遍历品类 cid：{cids}")

    result = {}  # 公司名 -> {phone, product, category, brandName, is_lenovo}

    for cid in cids:
        products = get_products(cid)
        print(f"\ncid={cid} 共 {len(products)} 个商品")
        for idx, p in enumerate(products, 1):
            if not (supplier_set - set(result.keys())):
                print("  所有供应商已匹配，提前结束")
                break
            sku_id = p["skuId"]
            brand = p["brandName"]
            is_lenovo = "联想" in brand or "lenovo" in brand.lower()
            dealers = get_dealers(sku_id)
            for dl in dealers:
                aname = (dl.get("agentName") or "").strip()
                if aname in supplier_set:
                    phone = dl.get("agentPhone", "") or ""
                    prev = result.get(aname)
                    # 联想产品优先覆盖非联想结果
                    if prev and prev.get("is_lenovo") and not is_lenovo:
                        continue
                    result[aname] = {
                        "phone": str(phone),
                        "product": p["skuName"],
                        "category": next((v["name"] for v in CATEGORIES.values() if v["cid"] == cid), str(cid)),
                        "brandName": brand,
                        "is_lenovo": is_lenovo,
                    }
            if idx % 20 == 0:
                print(f"  [{idx}/{len(products)}] 已匹配 {len(result)}/{len(supplier_set)}")
            time.sleep(0.15)

    print(f"\n匹配完成：{len(result)}/{len(supplier_set)} 家")
    unmatched = supplier_set - set(result.keys())
    if unmatched:
        print(f"未匹配 {len(unmatched)} 家：")
        for u in sorted(unmatched):
            print(f"  - {u}")

    out_path = args.out or "api_match_result.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f"结果已保存：{out_path}")


if __name__ == "__main__":
    main()
