# 京华云采 API 端点与参数（抓包逆向所得）

## 基础信息

- **两个域名**：
  - `https://mkt-bjzc.zhongcy.com` — 商品搜索、经销商查询（市场侧）
  - `https://shop-bjzc.zhongcy.com` — 销售记录查询（订单/成交侧）
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

### 3. 销售记录（shop 域名，GET 方法）

```
GET https://shop-bjzc.zhongcy.com/proxy/trade-service/mall/order/querySkuSaleRecord
```

Query 参数：
```
pageNum=1&pageSize=100&platformId=20&skuId=4272875
```

- **方法**：GET（query 参数），与其他两个 POST 接口不同
- **域名**：`shop-bjzc.zhongcy.com`（订单侧，非 mkt）
- 参数：`pageNum` 页码、`pageSize` 每页条数、`platformId=20`、`skuId` 商品ID

**返回结构（已实测）**：
```json
{
  "code": "0",
  "msg": "查询成功",
  "data": {
    "pageNum": 14,
    "pageSize": 10,
    "totalCount": 131,
    "totalPageCount": 14,
    "result": [
      {
        "organizeId": 19231,
        "organizeName": "北京市平谷区投资促进服务中心本级",   // 采购单位
        "shopId": 610,
        "shopName": "北京绿都畅达人才科技发展有限公司",       // 供应商
        "skuNum": 1,                                          // 采购数量
        "sellPrice": 5900.00000,                              // 采购单价(元)
        "orderTime": "2025-01-21 11:14:49"                    // 成交时间
      }
    ],
    "lastPage": true, "firstPage": false, "prevPage": 13, "nextPage": 15
  }
}
```

**字段映射**（销售记录 → 建表列）：
- `organizeName` → 采购单位
- `shopName` → 供应商
- `skuNum` → 采购数量
- `sellPrice` → 采购单价（元）
- `orderTime` → 成交时间
- 型号：记录里**没有型号字段**，需用商品搜索得到的 `skuName` 补

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
