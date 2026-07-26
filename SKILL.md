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

**主数据源：akshare** (`futures_zh_daily_sina`) — 返回结构化 DataFrame，含完整 OHLCV + settle。

**备用数据源：新浪 HTTP 直连** — 当 akshare 未安装时自动回退，纯 Python 标准库，零外部依赖。

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

### 云端自动化（GitHub Actions + Server酱 微信推送）

详见 `DEPLOY.md`。核心流程：

1. GitHub Actions cron 每交易日 07:15 UTC（= 15:15 CST）触发
2. 安装依赖 → 运行 `scripts/push_serverchan.py`
3. 采集数据 → 格式化为 Markdown → POST 到 Server酱 → 推送到微信
4. Excel 单日文件作为 artifact 上传（保留 90 天）

文件清单：
- `scripts/push_serverchan.py` — Server酱推送脚本
- `.github/workflows/cffex-daily.yml` — GitHub Actions 配置
- `requirements.txt` — Python 依赖（akshare, pandas, openpyxl）
- `DEPLOY.md` — 部署指南

环境变量：
- `SERVERCHAN_SENDKEY`（必填）— Server酱 SendKey
- `EXCEL_FILE`（可选）— Excel 输出路径
