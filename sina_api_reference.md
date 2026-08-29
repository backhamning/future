# 新浪财经 CFFEX 股指期货实时接口参考（CFF_RE_）

> 数据源：`hq.sinajs.cn/list=CFF_RE_<合约代码>`（多合约逗号分隔）
> 约定：本项目（cffex-futures 技能）**统一仅用新浪源**，不依赖 akshare / 东方财富。
> 字段映射经实测核对，核对日期：2026-08-28 / 2026-08-29（样本 IF2609 / IH2609）。

## 请求方式

```
GET https://hq.sinajs.cn/list=CFF_RE_IF2609,CFF_RE_IH2609
Header:
  Referer: https://finance.sina.com.cn/
  User-Agent: Mozilla/5.0 (必填，否则 403)
```

返回格式（每个合约一行，GBK 编码）：

```
var hq_str_CFF_RE_IF2609="今开,最高,最低,收盘,成交量,成交额,持仓量,...,昨收,昨结算,...,日期,时间,...,字段48,名称";
```

字段以逗号分隔，`var hq_str_CFF_RE_<code>="..."` 中的内容按索引取值。

## 字段映射（实测核对）

| 索引 | 字段 | 说明 | 备注 |
|------|------|------|------|
| [0]  | 今开 (open)        | 当日开盘价 | |
| [1]  | 最高 (high)        | 当日最高价 | |
| [2]  | 最低 (low)         | 当日最低价 | |
| [3]  | 收盘 (close)       | 当日收盘价 | = akshare `futures_zh_realtime` 的 close |
| [4]  | 成交量 (volume)    | 手 | 整数 |
| [5]  | 成交额 (amount)    | 元 | |
| [6]  | 持仓量 (open_interest / oi) | 持仓 | = open interest |
| [13] | 昨收 (preclose)    | 前一交易日收盘 | 已多合约验证 = 新浪日线前一交易日收盘 |
| [14] | 昨结算 (presettlement) | 前一交易日结算 | 与 akshare `presettlement` 一致 |
| [36] | 日期 (trade_date)  | YYYY-MM-DD | 格式如 `2026-08-28` |
| [37] | 时间               | HH:MM:SS | 如 `15:00:00` |
| [48] | ⚠️ 非结算价        | 实测 ≠ CFFEX 官方结算价 | 见下方说明 |
| [49] | 名称               | 合约中文名 | 如 `沪深300指数期货2609` |

## 重要澄清：结算价取不到

- **[48] 并非结算价**。实测 IF2609：Sina `[48]` = `4601.116`，而 CFFEX 官方结算价 = `4592.4`，二者不符。
- 盘后 akshare 实时 `settlement` 字段亦为 `0`（未回填）。
- 因此**本接口无法可靠提供结算价**，代码中 `settle` 字段统一置 `None`。如需结算价，应改用 CFFEX 官方结算参数或盘后官方文件。

## 本项目用法

- `fetch_close_prices.py`：`fetch_via_sina()` 解析 CFF_RE_，主源，带请求重试（3 次 / 退避 1.5s）。
- 收盘价（[3]）、昨收（[13]）、持仓量（[6]）、成交量（[4]）为可靠字段，邮件日报与历史导出均重度依赖。
- 结算价（[48]）不使用，相关列显示为 `—`。
- ETF / 指数实时行情走 `hq.sinajs.cn/list=sh510050` 等（不含 `CFF_RE_` 前缀），ETF 历史日线走 `money.finance.sina.com.cn` 的 `getKLineData`（scale=240）。
