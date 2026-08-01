---
name: cffex-futures
description: 中金所股指期货数据工具。获取 IH(上证50)/IC(中证500)/IF(沪深300)/IM(中证1000) 所有活跃合约的行情数据。支持每日收盘价自动获取、合约信息查询等。当用户需要查看股指期货收盘价、合约行情、或设置自动监控时触发。
agent_created: true
---

# CFFEX 股指期货数据工具

## Overview

覆盖中金所（CFFEX）四大股指期货品种（IH/IC/IF/IM）的行情数据获取。
合约月份规则：当月、下月、随后两个季月（3/6/9/12）。

## Data Source

**主数据源：akshare 实时接口** (`futures_zh_realtime`) — 返回当日实时行情，收盘后立即可取，一次拉整个品种所有合约。注意：结算价(settlement)在收盘后约16:00才公布，15:15时该字段为空。

**备用数据源 1：新浪 HTTP 直连** (`hq.sinajs.cn`) — 纯标准库，零外部依赖，当 akshare 不可用时自动回退。

**备用数据源 2：akshare 日线接口** (`futures_zh_daily_sina`) — 返回历史日线数据，有 1-2 小时延迟，作为最终兜底。

Python 环境：`C:/Users/LN/.workbuddy/binaries/python/envs/default/Scripts/python.exe`（安装了 akshare 1.18.78 的 venv）。

## Core Capabilities

### 1. 获取所有活跃合约收盘价

获取 IH/IC/IF/IM 全部活跃合约的最新交易日收盘价。自动发现合约、过滤已到期合约，只保留最新交易日数据。

**触发场景：**
- "获取股指期货收盘价"
- "IH/IC/IF/IM 今天收盘多少"
- "查看中金所期货行情"
- 配合自动化：每个交易日收盘后自动拉取

**执行方式：**

```bash
cd <skill_dir>
python scripts/fetch_close_prices.py              # 表格输出（需 venv 环境）
python scripts/fetch_close_prices.py --csv         # 保存 CSV
python scripts/fetch_close_prices.py --json        # JSON 输出
python scripts/fetch_close_prices.py --csv --output-dir ./data  # 指定输出目录
```

**Python 调用：**

```python
from scripts.fetch_close_prices import fetch_all

results = fetch_all()
for r in results:
    print(f"{r['code']}: close={r['close']}, volume={r['volume']}")
```

**输出：** 按品种分组的美化表格，包含合约、今开、最高、最低、收盘、结算、成交量、持仓量。CSV 文件以交易日命名（`cffex_futures_YYYY-MM-DD.csv`）。

**依赖：** Python 3.8+（主数据源需要 akshare；无 akshare 时自动回退到纯标准库模式）

### 合约发现规则

脚本自动查询未来 12 个月的所有合约代码，并按以下规则过滤：
1. 排除无成交量的合约
2. 排除到期月份距今超过 2 个月的合约
3. 仅保留最新交易日的数据（剔除已到期合约的历史数据）

## Automated Tasks

### 本地自动化（Excel 累积）

配合 WorkBuddy 自动化功能，每个交易日 15:15 运行：

```
Python: C:/Users/LN/.workbuddy/binaries/python/envs/default/Scripts/python.exe
任务: 运行 scripts/fetch_close_prices.py --excel --excel-file <数据目录>/cffex_daily.xlsx
调度: 每个交易日 15:15
输出: 累积 Excel 文件（汇总 sheet + 每日详情 sheet）
```

### 云端自动化（GitHub Actions + 邮件发送）

详见 `DEPLOY.md`。核心流程：

1. GitHub Actions cron 每交易日 07:01 UTC（= 15:01 CST，周一至周五）触发
2. 安装依赖 → 运行 `scripts/daily_collect_email.py --excel-file cffex_daily.xlsx`
3. 采集期货 + ETF 数据 → 生成 HTML 邮件 → SMTP 发送到 QQ 邮箱（凭据存于仓库 Secrets）
4. Server酱微信推送作为遗留步骤保留（`push_serverchan.py`，失败不阻断邮件）
5. 当日 Excel 作为 artifact 上传（保留 90 天）

所需 Secrets：
- `QQ_EMAIL_ACCOUNT` — QQ 邮箱账号
- `QQ_EMAIL_AUTH_CODE` — QQ 邮箱授权码
- `EMAIL_TO` — 收件人（可选，默认同账号）
- `SERVERCHAN_SENDKEY` — Server酱推送（遗留，可选）

文件清单：
- `scripts/fetch_close_prices.py` — 核心数据采集（支持 --json/--csv/--excel）
- `scripts/daily_collect_email.py` — 采集 + 生成 HTML 邮件 + SMTP 发送
- `scripts/push_serverchan.py` — Server酱推送脚本（遗留）
- `scripts/publish_draft.py` — 微信公众号草稿箱推送脚本（遗留）
- `scripts/export_history.py` — 历史数据导出为 Excel
- `.github/workflows/cffex-daily.yml` — GitHub Actions 配置
- `requirements.txt` — Python 依赖（akshare, pandas, openpyxl）
- `DEPLOY.md` — 部署指南

### 本地推送公众号草稿箱

本地 WorkBuddy 自动化可在生成 Excel 的同时，将格式化结果推送到个人公众号草稿箱。

```
Python: venv Python
任务: 运行 scripts/publish_draft.py --excel --excel-file <数据目录>/cffex_daily.xlsx
调度: 每个交易日 15:15（与 Excel 累积任务串联或独立执行）
```

前置条件：
1. 公众号后台获取 AppID/AppSecret（mp.weixin.qq.com → 开发 → 基本配置）
2. 将本机公网 IP 添加至 IP 白名单（`curl ifconfig.me` 查看）
3. 设置环境变量 `WECHAT_APP_ID` 和 `WECHAT_APP_SECRET`

操作步骤：
1. 自动化运行后，草稿自动出现在 mp.weixin.qq.com → 草稿箱
2. 手动点"发布"即可

环境变量：
- `SERVERCHAN_SENDKEY`（必填）— Server酱 SendKey
- `EXCEL_FILE`（可选）— Excel 输出路径
- `WECHAT_APP_ID`（可选）— 公众号 AppID，用于草稿箱推送
- `WECHAT_APP_SECRET`（可选）— 公众号 AppSecret，用于草稿箱推送
