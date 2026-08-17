# 京华云采 API 端点与参数（抓包逆向所得）

## 基础信息

- **域名**：`https://mkt-bjzc.zhongcy.com`
- **认证**：无需登录，普通 HTTP 请求即可（带 Referer/Origin 头即可通过）
- **platformId**：固定 `20`
- **publishType**：固定 `10024`

## 关键端点

### 1. 商品搜索
```
POST /proxy/trade-service/mall/search/querySkuListFromEs
```
请求体：
```json
{
  "queryPage": {"platformId": 20, "pageSize": 100, "pageNum": 1},
  "homeType": 110,
  "isAggregation": false,
  "isHometypes": true,
  "cid": 1000011,
  "businessType": 1,
  "publishType": 10024
}
```
返回：`data.itemList.resultList` 是商品列表，含 `skuId`、`skuName`、`brandNameCh`、`brandNameEn`；`data.totalCount` 是总数。

### 2. 经销商列表（核心，直接返回电话）
```
POST /proxy/trade-service/mall/search/querySkuAgentListFromEs
```
请求体：
```json
{
  "platformId": 20,
  "skuId": "某商品skuId",
  "shopName": "",
  "queryPage": {"platformId": 20, "pageSize": 100, "pageNum": 1},
  "publishType": 10024
}
```
返回：`data.itemAgentList.resultList` 是经销商列表，**每条直接含 `agentName` 和 `agentPhone`**——无需点击弹窗，这是抓包的关键发现。

## 品类 cid 对照表

| 品类 | cid |
|------|-----|
| 台式计算机 | 1000011 |
| 便携式计算机 | 1000014 |
| 服务器 | 1000128 |
| 打印机（喷墨/激光/针式/多功能/复印等子类） | 1000023、1000025、1000027、1000029、1000041 |

## 关键约束

1. **pageSize 上限 100**：超过 100 会触发服务端排序异常（500 会报错），经销商分页必须用 100
2. **返回码**：`code == "0"` 表示成功；`"99999"` 开头的是系统异常，可重试
3. **限频**：请求间加 0.15~0.3 秒 sleep，避免被封
4. **Referer**：查商品和查经销商的 Referer 可以不同，但都要带 `https://mkt-bjzc.zhongcy.com` 域名

## 供应商名称匹配策略

- API 返回的 `agentName` 是全称（如"北京艾维克信息技术有限公司"）
- 名单中的公司名可能与 API 名称有细微差异（空格、括号全半角），建议先精确匹配，未命中再做归一化模糊匹配
- 一个供应商可能出现在多个商品下，联想产品的结果优先保留（`is_lenovo` 标记）
