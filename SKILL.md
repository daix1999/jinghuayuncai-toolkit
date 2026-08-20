---
name: jinghuayuncai-toolkit
slug: jinghuayuncai-toolkit
displayName: 京华云采工具包
version: 1.6.0
summary: 获取京华云采经销商联系方式与销售数据：名单补电话 + 按产品/品类/品牌查经销商 + 拉销售成交记录
description: 京华云采（北京市政府采购平台）经销商与成交数据工具。当用户需要补全供应商名单缺失电话、按产品/品类/品牌查经销商、或拉取机型历史成交记录时使用。三大能力：①名单补电话——先走官网 API 抓包查询（querySkuAgentListFromEs 直接返回 agentPhone），未命中再公开搜索，标注电话来源和联系人身份；②查经销商——按指定产品、品类或品牌，爬取全部经销商及联系方式，输出"品类、品牌、产品、经销商名称、电话"表；③销售记录——按型号、品牌或品类拉取历史成交，输出"型号、采购单位、采购数量、采购单价、供应商、成交时间"表。关键词：京华云采、补电话、填充电话、查经销商、获取联系方式、销售记录、成交记录、联想经销商、渠道拓展。
tags: [京华云采, 经销商, 联系方式, 电话, 销售记录, 成交数据, 渠道拓展, 政府采购, 供应商, 办公效率]
license: MIT
agent_created: true
---

# 京华云采工具包

京华云采（北京市政府采购平台）经销商与销售数据工具：**名单补电话、按维度圈经销商、拉销售成交记录并分析**。

## 四大能力（触发词）

| 能力 | 触发示例 | 核心命令 |
|------|---------|---------|
| **① 名单补电话** | "把这张名单电话为空的补一下" / "补全联系方式" | `python scripts/run.py 名单.xlsx --name-col 供应商名称 --phone-col 联系电话` |
| **② 查经销商** | "查联想所有产品的经销商" / "台式机品类有哪些经销商" | `python scripts/crawl_dealers.py --brand 联想 --out 联想经销商.xlsx` |
| **③ 销售记录+分析** | "查开天M70d的销售记录" / "分析这批成交" | `python scripts/fetch_sale_records.py --product 开天M70d --out 销售.xlsx` → `python scripts/analyze_sales.py 销售.xlsx --out 分析.xlsx` |
| **④ 经销商画像** | "查XX公司的全景数据" / "这家经销商代理什么、客户是谁" | `python scripts/dealer_profile.py "公司名" --index sku_index.json --out 画像.xlsx` |

## 快速上手

1. **补电话**：优先 `run.py` 一键（找空号→API查→填充）；API 未命中的公司 ≥5 家时**默认走风鸟批量查**（riskbird-cominfo-batch），**不要逐家 WebSearch**（省 token）
2. **查经销商**：`crawl_dealers.py --brand/--category/--product` 任意组合
3. **销售记录**：`fetch_sale_records.py --product/--brand/--category`
4. **分析**：`analyze_sales.py` 输出 12 个 Sheet，每个含「汇总+关键订单明细+分析说明」
5. **经销商画像**：先建商品索引 `python scripts/build_sku_index.py --out sku_index.json --all`（全品类约 1 万商品，几分钟），再 `dealer_profile.py "公司名"` 看代理商品/品牌/成交/客户

## 关键经验速查（务必遵守）

1. **副本操作**：只输出 `原名_已填充.xlsx`，绝不覆盖原始文件
2. **来源必标**：每个电话标注来源（京华云采API / 公开搜索 / 未找到）
3. **不猜号码**：打码无法确认时标"未找到"，错号比空号更伤信任
4. **API 限频**：请求间 sleep 0.15~0.3 秒，pageSize 固定 100
5. **列名或列号**：`--name-col 供应商名称` 或 `A` 都行，不用手动数列号
6. **型号拆配置**：价格分析按「型号+配置」分组，同配置明显波动（>100元或>2%）才是议价信号

## 文档索引（按需读，不用全读）

| 文件 | 内容 |
|------|------|
| `references/workflows.md` | **详细工作流**：三大能力完整步骤、检查点、异常处理、输入规范、端到端案例 |
| `references/api_docs.md` | API 端点、参数、品类 cid 对照表 |
| `references/search_tips.md` | 公开搜索渠道与技巧 |
| `FAQ.md` | 常见问题（输入/结果/网络/数据/效率/边界） |
| `README.md` | 介绍与效果示例 |

> 执行具体任务时，先看 `references/workflows.md` 对应能力的工作流。
