#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
京华云采销售数据分析（v2：数据 + 说明型报告）

用法：
    python analyze_sales.py <销售记录.xlsx> [--out 分析报告.xlsx]

输入：fetch_sale_records.py 输出的销售记录 Excel（采购单位/采购数量/采购单价/供应商/成交时间）
输出：多 Sheet 分析报告 Excel，每个 Sheet = 数据表格 + 自动生成的文字分析说明

Sheet 结构：
    Sheet1  销售明细          —— 原始成交数据
    Sheet2  整体分析报告      —— 文字为主的总结报告（概览/品牌/买方/供应商/价格/趋势/建议）
    Sheet3  品牌对比分析      —— 不同品牌之间对比（数据 + 说明）
    Sheet4  型号分析          —— 品牌 × 型号对比（数据 + 说明）
    Sheet5  单位类型分布      —— 政府/学校/医院/企业（数据 + 说明）
    Sheet6  单位×供应商往来   —— 多次往来关系（数据 + 说明）
    Sheet7  供应商占比        —— 垄断度判断（数据 + 说明）
    Sheet8  月度销售趋势      —— 环比增长（数据 + 说明）
    Sheet9  TOP采购单位       —— 大客户排行（数据 + 说明）
    Sheet10 价格分析          —— 均价/区间/波动（数据 + 说明）
    Sheet11 采购频次          —— 大单 vs 频次单（数据 + 说明）
"""
import argparse
import os
import sys

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font


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


def append_insights(ws, start_row, title, insights, col=1):
    """在 sheet 的 start_row 下方写入分析说明文本块"""
    ws.cell(row=start_row, column=col, value=title).font = Font(bold=True, size=12)
    for i, line in enumerate(insights, 1):
        ws.cell(row=start_row + i, column=col, value=line)
    return start_row + len(insights) + 2


def write_data_plus_insights(writer, sheet_name, df_data, insights):
    """写一个分析 sheet：数据表 + 下方分析说明"""
    df_data.to_excel(writer, sheet_name=sheet_name, index=False)
    ws = writer.sheets[sheet_name]
    start = len(df_data) + 3
    append_insights(ws, start, '📌 分析说明', insights)
    return ws


def gen_overall_report(df, n, total_qty, total_amt, n_org, n_shop,
                       type_stat, shop_stat, monthly, org_stat, prices,
                       brand_stat, model_stat, freq_stat, model_col):
    """生成 Sheet2 整体分析报告（文字为主）"""
    L = []
    L.append('【京华云采销售数据分析报告】')
    L.append('')
    L.append('一、数据概览')
    L.append(f'  本次共分析 {n} 条成交记录，覆盖 {n_org} 家采购单位、{n_shop} 家供应商。')
    L.append(f'  总销量 {total_qty} 台，成交总金额 ¥{total_amt:,.0f}，均价 ¥{prices.mean():,.0f}。')
    L.append('')
    L.append('二、品牌格局')
    if brand_stat is not None and len(brand_stat) > 0:
        top_b = brand_stat.iloc[0]
        L.append(f'  共 {len(brand_stat)} 个品牌，主力品牌为【{top_b.name}】：{int(top_b["销量"])} 台、'
                 f'¥{top_b["金额"]:,.0f}（金额占比 {top_b["金额占比"]:.1f}%）。')
        if len(brand_stat) > 1:
            cr3 = brand_stat['金额占比'].head(3).sum()
            L.append(f'  TOP3 品牌金额集中度 {cr3:.1f}%' +
                     ('，品牌格局集中。' if cr3 > 70 else '，品牌格局较分散。'))
    L.append('')
    L.append('三、买方结构（单位类型）')
    t1 = type_stat.iloc[0]
    L.append(f'  成交最多的类型是【{t1.name}】：{int(t1["销量"])} 台、¥{t1["金额"]:,.0f}（金额占比 {t1["金额占比"]:.1f}%）。')
    if '学校/教育' in type_stat.index and type_stat.loc['学校/教育', '金额占比'] < 5:
        L.append('  学校/教育 渗透率 <5%，是潜在空白市场。')
    if '医院/医疗' in type_stat.index and type_stat.loc['医院/医疗', '金额占比'] < 5:
        L.append('  医院/医疗 渗透率 <5%，是潜在空白市场。')
    L.append('')
    L.append('四、供应商格局')
    s1 = shop_stat.iloc[0]
    L.append(f'  共 {len(shop_stat)} 家供应商供货，头部为【{s1.name}】：金额占比 {s1["金额占比"]:.1f}%。')
    cr3 = shop_stat['金额占比'].head(3).sum()
    L.append(f'  TOP3 供应商金额集中度 {cr3:.1f}%' +
             ('，较集中（头部有壁垒）。' if cr3 > 60 else '，较分散（有切入空间）。'))
    L.append('')
    L.append('五、价格信号')
    L.append(f'  成交均价 ¥{prices.mean():,.0f}，区间 ¥{prices.min():,.0f} - ¥{prices.max():,.0f}，'
             f'中位数 ¥{prices.median():,.0f}。')
    spread = (prices.max() - prices.min()) / prices.mean() * 100
    L.append(f'  价格波动幅度 {spread:.0f}%' + ('，价格弹性大（议价空间存在）。' if spread > 10 else '，价格稳定（市场定价成熟）。'))
    L.append('')
    L.append('六、趋势与节奏')
    if len(monthly) > 0:
        peak = monthly['销量'].idxmax()
        peak_row = monthly.loc[peak]
        L.append(f'  覆盖 {len(monthly)} 个月，销量峰值在【{peak}】：{int(peak_row["销量"])} 台、¥{peak_row["金额"]:,.0f}。')
        growth = monthly['销量'].pct_change().max()
        if pd.notna(growth) and growth > 0.5:
            L.append(f'  单月最大环比增长 {growth*100:.0f}%，存在明显的采购波峰。')
    L.append('')
    L.append('七、结论与建议')
    if '学校/教育' in type_stat.index and type_stat.loc['学校/教育', '金额占比'] < 5:
        L.append('  ▶ 学校/医院类单位渗透低，可作为新客户突破口。')
    if cr3 <= 60:
        L.append('  ▶ 供应商格局分散，暂无垄断者，适合切入供货。')
    if model_col and model_stat is not None and len(model_stat) > 0:
        top_m = model_stat.iloc[0]
        L.append(f'  ▶ 走量型号【{str(top_m.name)[:30]}】（{int(top_m["销量"])} 台），备货与报价优先覆盖。')
    L.append('  ▶ 建议按"类型空白 + 分散格局"两个信号，优先开发学校/医院类单位。')
    return L


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
    brand_col = find_col(df, '品牌')
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
    print('京华云采销售数据分析（v2）')
    print(f'  记录数：{n} 条 | 采购单位：{n_org} 家 | 供应商：{n_shop} 家')
    print(f'  总销量：{total_qty} 台 | 总金额：¥{total_amt:,.0f}')
    print('=' * 56)

    # ============ 各维度聚合 ============
    # 品牌对比
    brand_stat = None
    if brand_col:
        brand_stat = df.groupby(brand_col).agg(
            销量=('采购数量', 'sum'), 金额=('成交金额', 'sum'),
            客户数=('采购单位', 'nunique'), 记录数=('采购单位', 'count'),
        ).sort_values('金额', ascending=False)
        brand_stat['金额占比'] = (brand_stat['金额'] / total_amt * 100).round(1)

    # 型号分析（品牌×型号）
    model_stat = None
    if model_col:
        gb_cols = [brand_col, model_col] if brand_col else [model_col]
        model_stat = df.groupby(gb_cols).agg(
            销量=('采购数量', 'sum'), 金额=('成交金额', 'sum'),
            记录数=('采购单位', 'count'),
        ).sort_values('销量', ascending=False).reset_index()
        model_stat.columns = (['品牌', '型号', '销量', '金额', '记录数']
                              if brand_col else ['型号', '销量', '金额', '记录数'])

    # 单位类型
    type_stat = df.groupby('单位类型').agg(
        采购单位数=('采购单位', 'nunique'), 销量=('采购数量', 'sum'),
        金额=('成交金额', 'sum'), 记录数=('采购单位', 'count'),
    ).sort_values('金额', ascending=False)
    type_stat['金额占比'] = (type_stat['金额'] / total_amt * 100).round(1)

    # 多次往来
    pair = df.groupby([org_col, shop_col]).size().reset_index(name='次数')
    pair['销量'] = df.groupby([org_col, shop_col])['采购数量'].sum().values
    pair['金额'] = df.groupby([org_col, shop_col])['成交金额'].sum().values
    pair.columns = ['采购单位', '供应商', '次数', '销量', '金额']
    repeat = pair[pair['次数'] > 1].sort_values('次数', ascending=False)

    # 供应商占比
    shop_stat = df.groupby(shop_col).agg(
        客户数=('采购单位', 'nunique'), 销量=('采购数量', 'sum'),
        金额=('成交金额', 'sum'),
    ).sort_values('金额', ascending=False)
    shop_stat['金额占比'] = (shop_stat['金额'] / total_amt * 100).round(1)
    shop_stat['销量占比'] = (shop_stat['销量'] / total_qty * 100).round(1)

    # 月度趋势
    monthly = df.groupby('月份').agg(
        销量=('采购数量', 'sum'), 金额=('成交金额', 'sum'),
        记录数=('采购单位', 'count'),
    ).sort_index()
    monthly['销量环比%'] = monthly['销量'].pct_change().mul(100).round(1)

    # TOP 采购单位
    org_stat = df.groupby(org_col).agg(
        次数=('采购单位', 'count'), 销量=('采购数量', 'sum'),
        金额=('成交金额', 'sum'),
    ).sort_values('金额', ascending=False)
    org_stat['金额占比'] = (org_stat['金额'] / total_amt * 100).round(1)

    # 价格
    prices = df['采购单价'].dropna()
    price_stat = pd.DataFrame({
        '指标': ['均价', '最高价', '最低价', '中位数', '价格记录数'],
        '数值': [round(prices.mean(), 1), prices.max(), prices.min(),
                round(prices.median(), 1), len(prices)],
    })

    # 采购频次
    freq = df.groupby(org_col).size().reset_index(name='采购次数')
    freq.columns = ['采购单位', '采购次数']
    freq_bin = pd.cut(freq['采购次数'], bins=[0, 1, 2, 4, 100],
                      labels=['1次(单次)', '2次', '3-4次', '5次+'])
    freq_stat = freq.groupby(freq_bin, observed=True).size().reset_index(name='单位数')
    freq_stat.columns = ['采购频次区间', '单位数']

    # ============ 生成说明文本 ============
    def insights_brand():
        if brand_stat is None or len(brand_stat) == 0:
            return ['数据中无品牌字段，跳过品牌对比。']
        lines = [f'共 {len(brand_stat)} 个品牌参与成交。']
        top = brand_stat.iloc[0]
        lines.append(f'主力品牌【{top.name}】：{int(top["销量"])} 台、¥{top["金额"]:,.0f}，金额占比 {top["金额占比"]:.1f}%。')
        if len(brand_stat) > 1:
            cr3 = brand_stat['金额占比'].head(3).sum()
            lines.append(f'TOP3 品牌金额集中度 {cr3:.1f}%' + ('，品牌格局集中。' if cr3 > 70 else '，品牌格局较分散。'))
        return lines

    def insights_model():
        if model_stat is None:
            return ['数据中无型号字段，跳过型号分析。']
        lines = [f'共 {len(model_stat)} 个型号参与成交。']
        t1 = model_stat.iloc[0]
        lines.append(f'最走量型号【{str(t1["型号"])[:30]}】：{int(t1["销量"])} 台、¥{t1["金额"]:,.0f}。')
        if len(model_stat) > 1:
            top3_qty = model_stat['销量'].head(3).sum()
            lines.append(f'TOP3 型号销量合计 {int(top3_qty)} 台（占总量 {top3_qty/total_qty*100:.1f}%），'
                         + ('集中度高，主攻这几款即可。' if top3_qty/total_qty > 0.5 else '型号较分散，需多款备货。'))
        if brand_col:
            for b in brand_stat.index[:2]:
                sub = model_stat[model_stat['品牌'] == b]
                if len(sub):
                    m = sub.iloc[0]
                    lines.append(f'【{b}】下走量款：{str(m["型号"])[:24]}（{int(m["销量"])} 台）。')
        return lines

    def insights_type():
        lines = [f'共 {len(type_stat)} 类买方单位。']
        t1 = type_stat.iloc[0]
        lines.append(f'成交主力是【{t1.name}】：{int(t1["销量"])} 台、¥{t1["金额"]:,.0f}（{t1["金额占比"]:.1f}%）。')
        for t in ['学校/教育', '医院/医疗']:
            if t in type_stat.index:
                v = type_stat.loc[t, '金额占比']
                if v < 5:
                    lines.append(f'【{t}】金额占比仅 {v:.1f}%，渗透低，是新客户突破口。')
        return lines

    def insights_repeat():
        if len(repeat) == 0:
            return ['未发现同一单位×同一供应商多次成交，均为单次交易（关系尚未固化，正是切入时机）。']
        lines = [f'发现 {len(repeat)} 对固定供货关系（同一单位多次从同一供应商采购）。']
        for _, r in repeat.head(3).iterrows():
            lines.append(f'  · {r["采购单位"][:22]} ↔ {r["供应商"][:18]}：{int(r["次数"])} 次、{int(r["销量"])} 台')
        lines.append('这些是已固化的渠道关系；其余单位无固定供货商，是切入机会。')
        return lines

    def insights_shop():
        lines = [f'共 {len(shop_stat)} 家供应商参与供货。']
        s1 = shop_stat.iloc[0]
        lines.append(f'头部供应商【{s1.name}】：金额占比 {s1["金额占比"]:.1f}%、客户 {int(s1["客户数"])} 家。')
        cr3 = shop_stat['金额占比'].head(3).sum()
        lines.append(f'TOP3 供应商金额集中度 {cr3:.1f}%' +
                     ('，较集中（头部有壁垒，需从服务/价格切入）。' if cr3 > 60 else '，较分散（无垄断者，切入空间大）。'))
        return lines

    def insights_monthly():
        lines = [f'覆盖 {len(monthly)} 个月的成交。']
        if len(monthly) > 0:
            peak = monthly['销量'].idxmax()
            lines.append(f'销量峰值在【{peak}】：{int(monthly.loc[peak, "销量"])} 台、¥{monthly.loc[peak, "金额"]:,.0f}。')
            growth = monthly['销量'].pct_change()
            if growth.notna().any():
                gmax = growth.idxmax()
                lines.append(f'环比增长最大在【{gmax}】：{growth.loc[gmax]*100:.0f}%。')
        lines.append('若存在明显波峰，采购有周期性，建议在波峰前 1-2 个月做备货与报价准备。')
        return lines

    def insights_org():
        lines = []
        o1 = org_stat.iloc[0]
        lines.append(f'TOP 采购单位【{o1.name[:24]}】：{int(o1["次数"])} 次、{int(o1["销量"])} 台、¥{o1["金额"]:,.0f}（{o1["金额占比"]:.1f}%）。')
        top5_amt = org_stat['金额'].head(5).sum()
        lines.append(f'TOP5 单位合计 ¥{top5_amt:,.0f}，占总金额 {top5_amt/total_amt*100:.1f}%。')
        return lines

    def insights_price():
        lines = [f'成交均价 ¥{prices.mean():,.0f}，中位数 ¥{prices.median():,.0f}。',
                 f'价格区间 ¥{prices.min():,.0f} - ¥{prices.max():,.0f}。']
        spread = (prices.max() - prices.min()) / prices.mean() * 100
        lines.append(f'波动幅度 {spread:.0f}%' + ('，价格弹性大，存在议价空间。' if spread > 10 else '，价格稳定，市场定价成熟。'))
        return lines

    def insights_freq():
        lines = []
        single = int(freq_stat.loc[freq_stat['采购频次区间'] == '1次(单次)', '单位数'].sum())
        lines.append(f'单次采购单位 {single} 家（占 {single/n_org*100:.0f}%），属"碰运气型"成交。')
        high = int(freq_stat.loc[freq_stat['采购频次区间'].isin(['3-4次', '5次+']), '单位数'].sum())
        lines.append(f'高频复购单位（3 次以上）{high} 家，是值得长期维护的客户。')
        return lines

    # ============ 输出 Excel ============
    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
        # Sheet1 销售明细（原始数据）
        out_parts = []
        if brand_col:
            out_parts.append(('品牌', df[brand_col]))
        if model_col:
            out_parts.append(('型号', df[model_col]))
        out_parts.extend([
            ('采购单位', df[org_col]),
            ('供应商', df[shop_col]),
            ('采购数量', df[num_col]),
            ('采购单价(元)', df[price_col]),
            ('成交时间', df[time_col]),
        ])
        df_out = pd.DataFrame({k: v for k, v in out_parts})
        df_out.to_excel(writer, sheet_name='销售明细', index=False)

        # Sheet2 整体分析报告（文字为主）
        report = gen_overall_report(df, n, total_qty, total_amt, n_org, n_shop,
                                    type_stat, shop_stat, monthly, org_stat, prices,
                                    brand_stat, model_stat, freq_stat, model_col)
        ws2 = writer.book.create_sheet('整体分析报告')
        ws2.column_dimensions['A'].width = 110
        for i, line in enumerate(report, 1):
            ws2.cell(row=i, column=1, value=line).font = Font(size=11)
        ws2.cell(row=1, column=1).font = Font(bold=True, size=14)

        # Sheet3 品牌对比
        if brand_stat is not None:
            write_data_plus_insights(writer, '品牌对比分析', brand_stat.reset_index(), insights_brand())
        # Sheet4 型号分析
        if model_stat is not None:
            write_data_plus_insights(writer, '型号分析', model_stat, insights_model())
        # Sheet5 单位类型
        write_data_plus_insights(writer, '单位类型分布', type_stat.reset_index(), insights_type())
        # Sheet6 多次往来
        write_data_plus_insights(writer, '多次往来', repeat.reset_index(drop=True), insights_repeat())
        # Sheet7 供应商占比
        write_data_plus_insights(writer, '供应商占比', shop_stat.reset_index(), insights_shop())
        # Sheet8 月度趋势
        write_data_plus_insights(writer, '月度销售趋势', monthly.reset_index(), insights_monthly())
        # Sheet9 TOP 采购单位
        write_data_plus_insights(writer, 'TOP采购单位', org_stat.reset_index(), insights_org())
        # Sheet10 价格分析
        write_data_plus_insights(writer, '价格分析', price_stat, insights_price())
        # Sheet11 采购频次
        write_data_plus_insights(writer, '采购频次', freq_stat, insights_freq())

    # 控制台摘要
    print('\n■ 主要洞察：')
    for fn in (insights_brand, insights_type, insights_shop, insights_repeat):
        for line in fn():
            print('  ' + line)
        print()

    print(f'\n分析完成，报告已保存：{out_path}')
    with pd.ExcelFile(out_path) as xf:
        print(f'  Sheet：{xf.sheet_names}')


def main():
    parser = argparse.ArgumentParser(description='京华云采销售数据分析（数据+说明型报告）')
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
