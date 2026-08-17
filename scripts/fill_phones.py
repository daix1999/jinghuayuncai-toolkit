#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
京华云采名单电话填充脚本（通用版）
用法：
  python fill_phones.py <输入xlsx路径> <映射json路径> [--name-col A] [--phone-col D]
                                          [--remark-col F] [--out <输出路径>]

功能：
  - 读取供应商名单 Excel，找出电话列为空的行
  - 根据映射 JSON（公司名 -> {phone, source, contact}）填充电话
  - 新增"电话来源"列和"可能联系人及身份"列
  - 颜色标注：绿色=API匹配、黄色=公开搜索、橙色=未找到
  - 默认输出到"原名_已填充.xlsx"副本，不覆盖原文件
"""
import argparse
import json
import sys
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter, column_index_from_string


def col_to_idx(col):
    """把 'A'/'B' 或数字转换成 openpyxl 列索引（1-based）"""
    if isinstance(col, int):
        return col
    return column_index_from_string(col.strip().upper())


def is_empty(val):
    if val is None:
        return True
    s = str(val).strip()
    return s in ('', 'nan', 'None', '无', 'null')


def fill_phones(input_path, mapping, name_col='A', phone_col='D',
                remark_col=None, source_col_name='电话来源',
                contact_col_name='可能联系人及身份', out_path=None):
    wb = load_workbook(input_path)
    ws = wb.active

    name_idx = col_to_idx(name_col)
    phone_idx = col_to_idx(phone_col)
    remark_idx = col_to_idx(remark_col) if remark_col else None

    green_fill = PatternFill('solid', fgColor='C6EFCE')   # API 匹配
    yellow_fill = PatternFill('solid', fgColor='FFF2CC')  # 公开搜索
    orange_fill = PatternFill('solid', fgColor='FCE4D6')  # 未找到

    # 找表头行：扫描前 3 行，找到含名称列的行作为表头
    header_row = 1
    for r in range(1, 4):
        if ws.cell(row=r, column=name_idx).value:
            header_row = r
            break

    # 确定新增列位置
    max_col = ws.max_column
    source_idx = max_col + 1
    contact_idx = max_col + 2
    ws.cell(row=header_row, column=source_idx, value=source_col_name).font = Font(bold=True)
    ws.cell(row=header_row, column=contact_idx, value=contact_col_name).font = Font(bold=True)
    ws.column_dimensions[get_column_letter(source_idx)].width = 28
    ws.column_dimensions[get_column_letter(contact_idx)].width = 32

    filled = 0
    not_found = 0
    skipped = 0

    for row_idx in range(header_row + 1, ws.max_row + 1):
        name_cell = ws.cell(row=row_idx, column=name_idx)
        name = str(name_cell.value).strip() if name_cell.value else ''
        if not name or name in ('nan', 'None'):
            continue

        # 用规范化后的名称匹配
        key = name
        if key not in mapping:
            # 尝试去除首尾空白后匹配
            key = name.strip()
        if key not in mapping:
            continue

        phone_cell = ws.cell(row=row_idx, column=phone_idx)
        if not is_empty(phone_cell.value):
            skipped += 1
            continue

        entry = mapping[key]
        phone = entry.get('phone', '')
        source = entry.get('source', '')
        contact = entry.get('contact', '')

        if phone:
            phone_cell.value = phone
            phone_cell.fill = yellow_fill if ('搜索' in source or '公开' in source or 'web' in source.lower()) else green_fill
            filled += 1
        else:
            phone_cell.fill = orange_fill
            not_found += 1

        ws.cell(row=row_idx, column=source_idx).value = source or '需进一步获取'
        if contact:
            ws.cell(row=row_idx, column=contact_idx).value = contact

        if remark_idx and contact:
            rc = ws.cell(row=row_idx, column=remark_idx)
            existing = str(rc.value).strip() if rc.value and str(rc.value).strip() not in ('nan', 'None') else ''
            if existing and contact not in existing:
                rc.value = existing + '; ' + contact
            elif not existing:
                rc.value = contact

    if out_path is None:
        out_path = input_path.rsplit('.', 1)[0] + '_已填充.xlsx'
    wb.save(out_path)
    print(f'填充完成：{filled} 家已填，{not_found} 家未找到，{skipped} 家已有电话跳过')
    print(f'输出：{out_path}')
    return {'filled': filled, 'not_found': not_found, 'skipped': skipped, 'output': out_path}


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='京华云采名单电话填充')
    parser.add_argument('input', help='输入 xlsx 路径')
    parser.add_argument('mapping', help='映射 JSON 路径：{公司名: {phone, source, contact}}')
    parser.add_argument('--name-col', default='A', help='供应商名称列（默认 A）')
    parser.add_argument('--phone-col', default='D', help='电话列（默认 D）')
    parser.add_argument('--remark-col', default=None, help='备注列（可选，用于追加联系人）')
    parser.add_argument('--out', default=None, help='输出路径（默认 原名_已填充.xlsx）')
    args = parser.parse_args()

    with open(args.mapping, 'r', encoding='utf-8') as f:
        mapping = json.load(f)

    fill_phones(args.input, mapping, args.name_col, args.phone_col,
                args.remark_col, out_path=args.out)
