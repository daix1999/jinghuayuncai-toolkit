# 更新日志 (Changelog)

本项目的所有重要变更都记录在此文件中。格式遵循 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.0.0/)，版本号遵循 [语义化版本](https://semver.org/lang/zh-CN/)。

## [1.1.1] - 2026-08-17

### 修复

- **消除静默失败**：`fill_phones.py` 未匹配的公司不再静默跳过，改为统计并标注「未匹配(映射缺失)」；三个脚本请求彻底失败时打印 `⚠️ 请求失败（已重试 N 次）` 警告，网络波动不再静默丢失结果。
- **列位置错误提示**：填充时列参数填错，明确报错并列出真实表头，而非静默读错列。

### 新增

- **列名支持**：`fill_phones.py` 的 `--name-col` / `--phone-col` 支持列名（如 `供应商名称`）或列号（如 `A`），无需手动数列号。
- **完整结果摘要**：填充脚本结尾打印总行数 / 空号数 / 已填 / 未找到 / 未匹配 / 跳过，一眼判断结果完整性。

### 文档

- **SKILL.md 新增**：输入格式规范（Excel / 名单 JSON / 映射 JSON）、能力边界（API 依赖、品类覆盖、大批量耗时）、集中式常见问题 FAQ（7 条）、端到端使用案例。
- **安装说明同步**：补充列名支持与结果摘要说明，新增「找不到列」「判断成败」FAQ。

---

## [1.1.0] - 2026-08-17

### 新增

- **销售成交记录查询（能力三）**：新增 `scripts/fetch_sale_records.py`，按型号 / 品牌 / 品类拉取京华云采历史成交数据，输出「品类、品牌、型号、采购单位、采购数量、采购单价、供应商、成交时间」表，用于竞品分析、报价卡位、反向锁定供货渠道。
- **风鸟批量查询整合**：补电话环节中，官网 API 未命中的公司 ≥5 家时，优先调用 `riskbird-cominfo-batch` 技能走风鸟（riskbird.com）CDP 浏览器自动化批量提取电话/邮箱/法人/注册资本，替代逐家网页搜索；未安装该技能时自动降级为 WebSearch，主流程不受影响。
- **销售记录 API 文档**：`references/api_docs.md` 补充 `querySkuSaleRecord` 端点及双域名（mkt / shop）说明。

### 修复

- **代理绕过 SSL 问题**：三个脚本的 `requests.Session` 增加 `trust_env = False`，绕过系统本地代理（如 Clash 7897 端口），解决访问 zhongcy.com 时 TLS 握手失败的问题。

### 变更

- **技能更名**：显示名由「京华云采获取经销商联系方式」改为「京华云采工具包」，技术名/仓库名由 `jinghuayuncai-dealer-contact` 统一为 `jinghuayuncai-toolkit`。
- **文档同步**：README / SKILL.md / metadata.json / 安装使用说明 全部升级为三大能力说明，`search_tips.md` 新增「风鸟批量查询」渠道（第 0 优先级）。

### 技术细节

- 销售记录接口字段实测对齐：`organizeName`(采购单位)、`shopName`(供应商)、`skuNum`(数量)、`sellPrice`(单价)、`orderTime`(成交时间)，结果数组在 `data.result`，分页用 `data.totalPageCount`；记录本身不含型号字段，由商品搜索的 `skuName` 补全。

---

## [1.0.0] - 2026-08-13

### 新增

- **名单补电话（能力一）**：给定一份京华云采供应商名单（xlsx），补全缺失的联系电话。先走官网 API 抓包查询（`querySkuAgentListFromEs` 直接返回 `agentPhone`），未命中的再公开搜索，标注电话来源和联系人身份，输出副本不覆盖原文件。
- **经销商爬取（能力二）**：按产品 / 品类 / 品牌正向爬取京华云采上对应的所有经销商及联系方式，统一输出「品类、品牌、产品、经销商名称、电话」表。
- 脚本：`query_suppliers_by_api.py`（名单反向匹配）、`crawl_dealers.py`（经销商正向爬取）、`fill_phones.py`（Excel 填充 + 标注）。
- 参考文档：`api_docs.md`（API 端点/参数/cid 对照表）、`search_tips.md`（公开搜索渠道与技巧）。
