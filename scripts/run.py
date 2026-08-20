#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
京华云采工具包 - 一键补电话入口

用法：
    python run.py <名单.xlsx> [--categories 台,笔,服,印] [--name-col 供应商名称] [--phone-col 联系电话]

一键完成：
    1. 读名单，自动找电话为空的供应商
    2. 官网 API 查询这些供应商的电话
    3. 填充电话 + 标注来源 + 备注联系人
    4. 输出 原名_已填充.xlsx + 打印完整摘要 + 列出未匹配清单

说明：
    - 本脚本只做「API 查询 + 填充」，API 查不到的公司会标注「未匹配」，
      需后续配合 AI 走公开搜索（风鸟 / WebSearch）补漏
    - 列参数支持列名（如 供应商名称）或列号（如 A）
"""
import argparse
import json
import os
import subprocess
import sys
import tempfile

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
QUERY_SCRIPT = os.path.join(HERE, "query_suppliers_by_api.py")          # 兜底：遍历全部商品
QUERY_BY_NAME_SCRIPT = os.path.join(HERE, "query_suppliers_by_name.py")  # 快路径：按公司名定向反查
FILL_SCRIPT = os.path.join(HERE, "fill_phones.py")


def _resolve_col(df, col_spec):
    """解析列参数：支持列名(供应商名称)、字母列号(A)、数字列号(1)。返回列名。"""
    if isinstance(col_spec, int):
        return df.columns[col_spec - 1]
    s = str(col_spec).strip()
    if s.isdigit():
        return df.columns[int(s) - 1]
    if s.isalpha() and len(s) <= 3:
        idx = 0
        for ch in s.upper():
            idx = idx * 26 + (ord(ch) - ord('A') + 1)
        return df.columns[idx - 1]
    # 否则当作列名，尝试精确/模糊匹配
    if s in df.columns:
        return s
    for c in df.columns:
        if s in str(c):
            return c
    raise SystemExit(
        f"❌ 找不到列「{s}」。实际列名：{list(df.columns)}\n"
        f"   请用 --name-col / --phone-col 指定正确的列名或列号。"
    )


def find_empty_phones(input_path, name_col, phone_col):
    """读名单，找出电话为空的供应商，返回公司名列表"""
    df = pd.read_excel(input_path)
    name_c = _resolve_col(df, name_col)
    phone_c = _resolve_col(df, phone_col)

    empty_names = []
    for _, row in df.iterrows():
        name = str(row[name_c]).strip() if pd.notna(row[name_c]) else ''
        phone = row[phone_c]
        is_empty = pd.isna(phone) or str(phone).strip() in ('', 'nan', 'None', '无', 'null')
        if name and name != 'nan' and is_empty:
            empty_names.append(name)
    return empty_names


def main():
    parser = argparse.ArgumentParser(description="京华云采一键补电话")
    parser.add_argument("input", help="输入名单 xlsx")
    parser.add_argument("--categories", default="台,笔,服,印", help="品类键（台/笔/服/印），逗号分隔")
    parser.add_argument("--name-col", default="供应商名称", help="名称列（列名或列号）")
    parser.add_argument("--phone-col", default="联系电话", help="电话列（列名或列号）")
    parser.add_argument("--out", default=None, help="输出路径（默认 原名_已填充.xlsx）")
    args = parser.parse_args()

    input_path = os.path.abspath(args.input)
    if not os.path.exists(input_path):
        print(f"❌ 文件不存在：{input_path}")
        sys.exit(1)

    # 1. 找空号
    print("=" * 56)
    print("步骤 1/3：读取名单，找电话为空的供应商...")
    try:
        empty_names = find_empty_phones(input_path, args.name_col, args.phone_col)
    except SystemExit:
        raise
    except Exception as e:
        print(f"❌ 读取名单失败：{type(e).__name__} - {e}")
        sys.exit(1)
    if not empty_names:
        print("✅ 没有发现电话为空的供应商，无需填充")
        return
    print(f"  发现 {len(empty_names)} 家电话为空")

    # 2. API 查询（快路径优先：按公司名反查，一家 1-2 次请求，不遍历全部商品）
    print("\n步骤 2/3：官网 API 查询电话...")
    tmp = tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False, encoding='utf-8')
    json.dump(empty_names, tmp, ensure_ascii=False)
    tmp.close()
    names_file = tmp.name
    api_result_file = names_file.replace('.json', '_api.json')

    # 2a. 快路径：按公司名定向反查（新，快 50-100 倍）
    subprocess.run(
        [sys.executable, QUERY_BY_NAME_SCRIPT, names_file, "--out", api_result_file],
        check=False,
    )
    api_result = {}
    if os.path.exists(api_result_file):
        with open(api_result_file, 'r', encoding='utf-8') as f:
            api_result = json.load(f)

    # 2b. 兜底：快路径未命中的公司，再遍历商品（原方案，仅少量）
    remain = [n for n in empty_names if n not in api_result]
    if remain:
        print(f"  ⏳ 快路径未命中 {len(remain)} 家，用遍历商品兜底...")
        remain_file = names_file.replace('.json', '_remain.json')
        with open(remain_file, 'w', encoding='utf-8') as f:
            json.dump(remain, f, ensure_ascii=False)
        full_file = api_result_file.replace('.json', '_full.json')
        subprocess.run(
            [sys.executable, QUERY_SCRIPT, remain_file,
             "--categories", args.categories, "--out", full_file],
            check=False,
        )
        if os.path.exists(full_file):
            with open(full_file, 'r', encoding='utf-8') as f:
                api_result.update(json.load(f))
            os.remove(full_file)
        os.remove(remain_file)

    # 转成 mapping 格式（未命中的也保留，标注"未匹配"）
    mapping = {}
    for name in empty_names:
        if name in api_result:
            info = api_result[name]
            mapping[name] = {
                "phone": str(info.get("phone", "") or ""),
                "source": f"京华云采API（{info.get('product', '')}）",
                "contact": f"品牌:{info.get('brandName', '')}",
            }
        else:
            mapping[name] = {"phone": "", "source": "未匹配(API未命中)", "contact": ""}

    mapping_file = names_file.replace('.json', '_mapping.json')
    with open(mapping_file, 'w', encoding='utf-8') as f:
        json.dump(mapping, f, ensure_ascii=False, indent=2)

    # 3. 填充
    print("\n步骤 3/3：填充电话并生成副本...")
    out_path = args.out or input_path.rsplit('.', 1)[0] + '_已填充.xlsx'
    subprocess.run(
        [sys.executable, FILL_SCRIPT, input_path, mapping_file,
         "--name-col", args.name_col, "--phone-col", args.phone_col,
         "--out", out_path],
        check=False,
    )

    # 清理临时文件
    for f in (names_file, api_result_file, mapping_file):
        try:
            os.remove(f)
        except OSError:
            pass

    # 最终提示
    api_hit = sum(1 for n in empty_names if n in api_result)
    miss = len(empty_names) - api_hit
    print("\n" + "=" * 56)
    print(f"一键补电话完成：{len(empty_names)} 家空号，API 命中 {api_hit} 家，未命中 {miss} 家")
    if miss > 0:
        print(f"💡 未命中的 {miss} 家需公开搜索补充——可对我说「用风鸟批量查这些公司」并提供名单继续")
        print("  未命中清单：")
        for n in empty_names:
            if n not in api_result:
                print(f"    - {n}")
    print(f"输出：{out_path}")


if __name__ == "__main__":
    main()
