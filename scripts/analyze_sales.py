#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
京华云采销售数据分析

用法：
    python analyze_sales.py <销售记录.xlsx> [--out 分析报告.xlsx]

输入：fetch_sale_records.py 输出的销售记录 Excel（采购单位/采购数量/采购单价/供应商/成交时间）
输出：多 Sheet 分析报告 Excel + 控制台摘要

分析维度：
    1. 买方单位类型分布（政府/学校/医院/企业/其他）—— 看清政采盘 vs 企业盘
    2. 单位×供应商多次往来 —— 发现固定供货关系（壁垒/突破口）
    3. 供应商客户占比 —— 是否有垄断者（分散=有切入空间）
    4. 商品销售增长（按月） —— 判断市场热度与时机
    5. TOP 采购单位 —— 大客户在哪
    6. 价格分析（均价/区间/波动） —— 报价卡位依据
    7. 型号销量排行 —— 哪款是走量款
    8. 采购频次 —— 大单 vs 频次单
"""
import argparse
import os
import sys

import pandas as pd


def find_col(df, *candidates):
    """在 DataFrame 中按候选列名模糊匹配，返回真实列名"""
    for cand in candidates:
        for c in df.columns:
            if cand in str(c):
                return c
    return None


def classify_org(name):
    """按单位名称关键词分类买方类型（学校 > 医院 > 政府 > 企业 > 其他）"""
    school_kw = ['大学', '学院', '中学', '小学', '学校', '幼儿园', '党校',
                 '职业', '技师', '培训学校', '实验学校']
    hospital_kw = ['医院', '卫生院', '卫生服务中心', '疾控', '口腔',
                   '医疗', '妇幼保健', '康复中心', '卫生所']
    gov_kw = ['人民政府', '政府', '街道', '镇', '乡', '局', '委员会', '委',
              '法院', '检察院', '公安', '税务', '财政', '审计', '纪检',
              '人大', '政协', '机关事务', '管委会', '办公室', '监督管理']
    company_kw = ['有限公司', '集团', '股份', '公司', '厂', '合作社', '商贸', '科技']

    if any(k in name for k in school_kw):
        return '学校/教育'
    if any(k in name for k in hospital_kw):
        return '医院/医疗'
    if any(k in name for k in gov_kw):
        return '政府/机关'
    if any(k in name for k in company_kw):
        return '企业'
    return '其他/事业单位'


def analyze(input_path, out_path):
    df = pd.read_excel(input_path)
    if df.empty:
        print('❌ 数据为空，无法分析')
        sys.exit(1)

    # 列名定位（兼容不同版本输出）
    org_col = find_col(df, '采购单位')
    shop_col = find_col(df, '供应商')
    num_col = find_col(df, '采购数量', '数量')
    price_col = find_col(df, '采购单价', '单价')
    time_col = find_col(df, '成交时间', '下单时间', '时间')
    model_col = find_col(df, '型号', '商品')
    if not (org_col and shop_col and num_col and price_col and time_col):
        print(f'❌ 缺少必要列。当前列：{list(df.columns)}')
        print('   需要：采购单位 / 供应商 / 采购数量 / 采购单价 / 成交时间')
        sys.exit(1)

    df['采购数量'] = pd.to_numeric(df[num_col], errors='coerce').fillna(0)
    df['采购单价'] = pd.to_numeric(df[price_col], errors='coerce').fillna(0)
    df['成交金额'] = df['采购数量'] * df['采购单价']
    df['成交时间'] = pd.to_datetime(df[time_col], errors='coerce')
    df['月份'] = df['成交时间'].dt.strftime('%Y-%m')
    df['单位类型'] = df[org_col].astype(str).apply(classify_org)

    n = len(df)
    total_qty = int(df['采购数量'].sum())
    total_amt = df['成交金额'].sum()
    n_org = df[org_col].nunique()
    n_shop = df[shop_col].nunique()

    print('=' * 56)
    print('京华云采销售数据分析报告')
    print(f'  记录数：{n} 条 | 采购单位：{n_org} 家 | 供应商：{n_shop} 家')
    print(f'  总销量：{total_qty} 台 | 总金额：¥{total_amt:,.0f}')
    print('=' * 56)

    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
        # 1. 单位类型分布
        type_stat = df.groupby('单位类型').agg(
            采购单位数=('采购单位', 'nunique'),
            销量=('采购数量', 'sum'),
            金额=('成交金额', 'sum'),
            记录数=('采购单位', 'count'),
        ).sort_values('金额', ascending=False)
        type_stat['金额占比'] = (type_stat['金额'] / total_amt * 100).round(1)
        type_stat.to_excel(writer, sheet_name='单位类型分布')
        print('\n■ 买方单位类型分布（按金额排序）：')
        for t, r in type_stat.iterrows():
            print(f'  {t:<8} {int(r["采购单位数"])}家单位 {int(r["销量"])}台 ¥{r["金额"]:,.0f} ({r["金额占比"]}%)')

        # 2. 单位×供应商多次往来
        pair = df.groupby([org_col, shop_col]).agg(
            次数=('采购单位', 'count'),
            销量=('采购数量', 'sum'),
            金额=('成交金额', 'sum'),
        ).reset_index()
        pair.columns = ['采购单位', '供应商', '次数', '销量', '金额']
        repeat = pair[pair['次数'] > 1].sort_values('次数', ascending=False)
        repeat.to_excel(writer, sheet_name='多次往来', index=False)
        print(f'\n■ 单位×供应商多次往来：{len(repeat)} 对')
        for _, r in repeat.head(8).iterrows():
            print(f'  {r["采购单位"][:20]} ↔ {r["供应商"][:16]}：{int(r["次数"])}次 {int(r["销量"])}台')

        # 3. 供应商客户占比
        shop_stat = df.groupby(shop_col).agg(
            客户数=('采购单位', 'nunique'),
            销量=('采购数量', 'sum'),
            金额=('成交金额', 'sum'),
        ).sort_values('金额', ascending=False)
        shop_stat['金额占比'] = (shop_stat['金额'] / total_amt * 100).round(1)
        shop_stat['销量占比'] = (shop_stat['销量'] / total_qty * 100).round(1)
        shop_stat.to_excel(writer, sheet_name='供应商占比')
        top1 = shop_stat.iloc[0]
        print(f'\n■ 供应商占比（TOP3）：')
        for i, (s, r) in enumerate(shop_stat.head(3).iterrows(), 1):
            print(f'  {i}. {s[:20]}：金额占{r["金额占比"]}% / 客户{int(r["客户数"])}家')
        if len(shop_stat) > 1:
            cr = shop_stat['金额占比'].head(3).sum()
            print(f'  → TOP3 集中度 {cr:.1f}%' + ('（较集中，头部有壁垒）' if cr > 60 else '（较分散，有切入空间）'))

        # 4. 销售增长（按月）
        monthly = df.groupby('月份').agg(
            销量=('采购数量', 'sum'),
            金额=('成交金额', 'sum'),
            记录数=('采购单位', 'count'),
        ).sort_index()
        monthly['销量环比%'] = monthly['销量'].pct_change().mul(100).round(1)
        monthly.to_excel(writer, sheet_name='月度销售趋势')
        print(f'\n■ 月度销售趋势（共 {len(monthly)} 个月）：')
        prev = None
        for m, r in monthly.iterrows():
            trend = ''
            if prev is not None:
                d = r['销量'] - prev
                trend = f'  ({"▲+" if d >= 0 else "▼"}{abs(d)}台)'
            print(f'  {m}：{int(r["销量"])}台 ¥{r["金额"]:,.0f}{trend}')
            prev = r['销量']

        # 5. TOP 采购单位
        org_stat = df.groupby(org_col).agg(
            次数=('采购单位', 'count'),
            销量=('采购数量', 'sum'),
            金额=('成交金额', 'sum'),
        ).sort_values('金额', ascending=False)
        org_stat['金额占比'] = (org_stat['金额'] / total_amt * 100).round(1)
        org_stat.to_excel(writer, sheet_name='TOP采购单位')
        print(f'\n■ TOP 采购单位（前5）：')
        for i, (o, r) in enumerate(org_stat.head(5).iterrows(), 1):
            print(f'  {i}. {o[:26]}：{int(r["次数"])}次 {int(r["销量"])}台 ¥{r["金额"]:,.0f}')

        # 6. 价格分析
        prices = df['采购单价'].dropna()
        price_stat = pd.DataFrame({
            '指标': ['均价', '最高价', '最低价', '中位数', '价格记录数'],
            '数值': [round(prices.mean(), 1), prices.max(), prices.min(),
                    round(prices.median(), 1), len(prices)],
        })
        price_stat.to_excel(writer, sheet_name='价格分析', index=False)
        print(f'\n■ 价格分析：均价 ¥{prices.mean():,.0f} | 区间 ¥{prices.min():,.0f}-{prices.max():,.0f}')

        # 7. 型号销量排行
        if model_col:
            model_stat = df.groupby(model_col).agg(
                销量=('采购数量', 'sum'),
                金额=('成交金额', 'sum'),
                记录数=('采购单位', 'count'),
            ).sort_values('销量', ascending=False)
            model_stat.to_excel(writer, sheet_name='型号排行')
            print(f'\n■ 型号销量排行 TOP3：')
            for i, (m, r) in enumerate(model_stat.head(3).iterrows(), 1):
                print(f'  {i}. {str(m)[:30]}：{int(r["销量"])}台')

        # 8. 采购频次
        freq = df.groupby(org_col).size().reset_index(name='采购次数')
        freq.columns = ['采购单位', '采购次数']
        freq_bin = pd.cut(freq['采购次数'], bins=[0, 1, 2, 4, 100],
                          labels=['1次(单次)', '2次', '3-4次', '5次+'])
        freq_stat = freq.groupby(freq_bin, observed=True).size().reset_index(name='单位数')
        freq_stat.columns = ['采购频次区间', '单位数']
        freq_stat.to_excel(writer, sheet_name='采购频次', index=False)
        print(f'\n■ 采购频次分布：')
        for _, r in freq_stat.iterrows():
            print(f'  {r["采购频次区间"]}：{int(r["单位数"])} 家单位')

    print(f'\n分析完成，报告已保存：{out_path}')
    print(f'  Sheet：{pd.ExcelFile(out_path).sheet_names}')


def main():
    parser = argparse.ArgumentParser(description='京华云采销售数据分析')
    parser.add_argument('input', help='销售记录 xlsx（fetch_sale_records.py 的输出）')
    parser.add_argument('--out', default=None, help='分析报告输出路径（默认 原名_分析报告.xlsx）')
    args = parser.parse_args()

    input_path = os.path.abspath(args.input)
    if not os.path.exists(input_path):
        print(f'❌ 文件不存在：{input_path}')
        sys.exit(1)

    out_path = args.out or input_path.rsplit('.', 1)[0] + '_分析报告.xlsx'
    analyze(input_path, out_path)


if __name__ == '__main__':
    main()
