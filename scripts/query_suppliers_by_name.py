#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
京华云采供应商电话定向反查（快路径，替代遍历全部商品）

用法：
    python query_suppliers_by_name.py <供应商名单.json> [--out api_match_result.json] [--max-pages 1]

原理（2026-08-20 实测发现）：
    querySkuAgentListFromEs 接口支持 shopName 参数——不带 skuId、只带 shopName，
    可直接按公司名全局反查，一家公司 1-2 次请求就能拿到 agentPhone，
    无需遍历各品类全部商品（原方案 500+ 次请求 → 现在 = 名单公司数 × 1-2 次）。

输出：
    --out 指定的 JSON：{公司名: {phone, source, product}}（命中）
    未命中的公司打印清单，供公开搜索兜底。

备注：
    - 反查结果不含商品名（skuName 为空），source 标注「京华云采API(反查)」
    - 联想官方等品牌方不在经销商体系，会命中 0 条，属正常
"""
import argparse
import json
import os
import sys
import time

import requests

BASE = 'https://mkt-bjzc.zhongcy.com'
AGENT_URL = BASE + '/proxy/trade-service/mall/search/querySkuAgentListFromEs'
PLATFORM_ID = 20
PUBLISH_TYPE = 10024

SESSION = requests.Session()
SESSION.trust_env = False  # 绕过系统代理直连（避免代理导致 TLS 握手失败）
SESSION.headers.update({
    "Content-Type": "application/json",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Referer": BASE + "/mall-view/product/search",
    "Origin": BASE,
})


def normalize_name(name):
    """名称归一化：去首尾空白、全半角括号转换（用于尝试变体匹配）"""
    s = str(name).strip()
    s = s.replace('（', '(').replace('）', ')')
    s = s.replace(' ', '')
    return s


def query_by_name(shop_name, page=1, retries=3):
    """按公司名反查，返回 resultList（可能为空）"""
    payload = {
        "platformId": PLATFORM_ID,
        "shopName": shop_name,
        "queryPage": {"platformId": PLATFORM_ID, "pageSize": 100, "pageNum": page},
        "publishType": PUBLISH_TYPE,
    }
    last_err = "未知错误"
    for i in range(retries):
        try:
            r = SESSION.post(AGENT_URL, json=payload, timeout=30)
            d = r.json()
            if d.get("code") == "0":
                return (d.get("data") or {}).get("itemAgentList") or {}
            if "99999" in str(d.get("code", "")):
                time.sleep(0.5 * (i + 2))
                continue
            return {}
        except requests.exceptions.Timeout:
            last_err = "网络超时"
            time.sleep(0.5 * (i + 2))
        except requests.exceptions.ConnectionError:
            last_err = "网络连接失败"
            time.sleep(0.5 * (i + 2))
        except requests.exceptions.SSLError:
            last_err = "SSL/代理问题"
            time.sleep(0.5 * (i + 2))
        except Exception as e:
            last_err = type(e).__name__
            time.sleep(0.5 * (i + 2))
    print(f"⚠️ 请求失败[{last_err}]（已重试 {retries} 次）：{shop_name}，请检查网络或代理")
    return {}


def main():
    parser = argparse.ArgumentParser(description="京华云采供应商电话定向反查（快路径）")
    parser.add_argument("supplier_list", help="供应商名单 JSON：公司名列表")
    parser.add_argument("--out", default=None, help="结果 JSON 输出路径")
    parser.add_argument("--max-pages", type=int, default=1,
                        help="每家公司最多翻页数（默认 1，第 1 页就有电话）")
    args = parser.parse_args()

    with open(args.supplier_list, "r", encoding="utf-8") as f:
        suppliers = json.load(f)
    supplier_set = [str(s).strip() for s in suppliers if str(s).strip()]
    print(f"待反查供应商：{len(supplier_set)} 家")

    result = {}
    unmatched = []
    seen_variants = set()

    for idx, name in enumerate(supplier_set, 1):
        if name in seen_variants:
            continue
        seen_variants.add(name)

        ia = query_by_name(name)
        rl = ia.get("resultList") or []
        if rl:
            phone = str(rl[0].get("agentPhone", "") or "")
            result[name] = {
                "phone": phone,
                "source": "京华云采API(反查)",
                "product": "",
            }
            # 若 phone 为空，翻几页再找（极少见）
            if not phone:
                for p in range(2, args.max_pages + 1):
                    ia2 = query_by_name(name, page=p)
                    rl2 = ia2.get("resultList") or []
                    if rl2 and rl2[0].get("agentPhone"):
                        result[name]["phone"] = str(rl2[0]["agentPhone"])
                        break
                    time.sleep(0.2)
        else:
            # 尝试归一化变体（去空格/全半角括号）
            norm = normalize_name(name)
            if norm != name:
                ia3 = query_by_name(norm)
                rl3 = ia3.get("resultList") or []
                if rl3:
                    result[name] = {
                        "phone": str(rl3[0].get("agentPhone", "") or ""),
                        "source": "京华云采API(反查-归一化)",
                        "product": "",
                    }
                else:
                    unmatched.append(name)
            else:
                unmatched.append(name)

        if idx % 10 == 0 or idx == len(supplier_set):
            print(f"  [{idx}/{len(supplier_set)}] 已命中 {len(result)} 家")
        time.sleep(0.2)

    print(f"\n反查完成：命中 {len(result)}/{len(supplier_set)} 家")
    if unmatched:
        print(f"未命中 {len(unmatched)} 家（可走风鸟/公开搜索兜底）：")
        for u in unmatched:
            print(f"  - {u}")

    out_path = args.out or "api_match_result.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f"结果已保存：{out_path}")


if __name__ == "__main__":
    main()
