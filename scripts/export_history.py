#!/usr/bin/env python3
"""
导出指定日期范围的 CFFEX 股指期货历史数据到 Excel。

用法:
    python export_history.py                           # 默认 2026-06-01 ~ 今天
    python export_history.py --start 2026-01-01 --end 2026-06-30
    python export_history.py --output ./my_data.xlsx
"""

import argparse
import json
import os
import re
import sys
import time
from datetime import date, datetime, timedelta

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

INSTRUMENTS = ["IH", "IF", "IC", "IM"]
INSTRUMENT_NAMES = {"IH": "上证50", "IF": "沪深300", "IC": "中证500", "IM": "中证1000"}

# 合约探测范围：当前月份前后各 12 个月
MONTH_RANGE = 12

HEADER_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
HEADER_FONT = Font(name="微软雅黑", bold=True, color="FFFFFF", size=10)
TITLE_FONT = Font(name="微软雅黑", bold=True, size=13)
DATA_FONT = Font(name="微软雅黑", size=10)
THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)
RED_FILL = PatternFill(start_color="FFF2F2", end_color="FFF2F2", fill_type="solid")  # 贴水行
COLORS = {
    "IH": "4472C4",
    "IF": "ED7D31",
    "IC": "70AD47",
    "IM": "FFC000",
}


def _safe_float(val):
    """安全转 float，失败返回 None。"""
    if val is None:
        return None
    try:
        return float(val)
    except (ValueError, TypeError):
        return None


def _safe_int(val):
    """安全转 int，失败返回 None。"""
    if val is None:
        return None
    try:
        return int(float(val))
    except (ValueError, TypeError):
        return None


def gen_monthly_contracts(year, month, count):
    """生成连续月份合约代码（如 IH2608）。"""
    codes = []
    for i in range(count):
        m = month + i
        y = year + (m - 1) // 12
        m = (m - 1) % 12 + 1
        for sym in INSTRUMENTS:
            codes.append(f"{sym}{y % 100:02d}{m:02d}")
    return codes


def fetch_sina_futures_kline(symbol, start_date, end_date, max_retry=3):
    """从新浪期货日 K 线获取合约历史。

    akshare 的 futures_zh_daily_sina 当前对所有 CFFEX 股指期货代码稳定抛
    IndexError（接口已不可用），改用新浪 InnerFuturesNewService.getDailyKLine。

    字段映射（已验证）：o=今开 h=最高 l=最低 c=收盘 v=成交量 p=持仓 s=结算。
    返回 DataFrame: date,open,high,low,close,volume,hold,settle（空则为空 DataFrame）。
    新浪接口易限流，空结果/异常自动退避重试，避免把"限流空"误判为"无数据"。
    """
    import pandas as pd
    import json as _json
    import time
    from urllib.request import Request, urlopen

    url = (
        "https://stock2.finance.sina.com.cn/futures/api/jsonp_v2.php/var%20_=/"
        f"InnerFuturesNewService.getDailyKLine?symbol={symbol}"
        f"&startDate={start_date}&endDate={end_date}"
    )
    for attempt in range(max_retry):
        try:
            req = Request(url, headers={"User-Agent": "Mozilla/5.0", "Referer": "https://finance.sina.com.cn/"})
            with urlopen(req, timeout=15) as resp:
                text = resp.read().decode("gbk", "replace")
            s = text.find("(")
            e = text.rfind(")")
            arr = _json.loads(text[s + 1:e]) if s >= 0 and e > s else _json.loads(text)

            rows = []
            for it in arr:
                d = str(it.get("d", ""))
                if not d or not (start_date <= d <= end_date):
                    continue
                s_val = _safe_float(it.get("s"))
                rows.append({
                    "date": d,
                    "open": _safe_float(it.get("o")),
                    "high": _safe_float(it.get("h")),
                    "low": _safe_float(it.get("l")),
                    "close": _safe_float(it.get("c")),
                    "volume": _safe_int(it.get("v")),
                    "hold": _safe_int(it.get("p")),
                    "settle": (s_val if s_val and s_val != 0 else None),
                })
            if rows:
                return pd.DataFrame(rows)
            # 空结果：多为限流，退避后重试
        except Exception:
            pass
        if attempt < max_retry - 1:
            time.sleep(1.0)
    return pd.DataFrame()


def discover_active_contracts(min_last_date=None):
    """查询并返回活跃的合约代码列表。"""
    from datetime import timedelta

    today = date.today()
    if min_last_date is None:
        min_last_date = (today - timedelta(days=90)).strftime("%Y-%m-%d")

    # 确定需要覆盖的合约月份范围：从 min_last_date 所在月到 today + MONTH_RANGE
    min_y, min_m = int(min_last_date[:4]), int(min_last_date[:7].split("-")[1])
    months_back = (today.year - min_y) * 12 + (today.month - min_m)
    total_months = months_back + MONTH_RANGE + 2  # +2 这个留裕量
    candidates = gen_monthly_contracts(min_y, min_m, total_months)

    active = set()
    for code in candidates:
        try:
            # 用新浪期货 K 线判断合约是否在范围内有数据（akshare 日线已坏）
            df = fetch_sina_futures_kline(code, min_last_date, today.isoformat())
            if df is None or df.empty:
                continue
            active.add(code)
        except Exception:
            pass

    # 排序：IH → IF → IC → IM，同品种按月份
    instr_order = {"IH": 0, "IF": 1, "IC": 2, "IM": 3}
    return sorted(active, key=lambda c: (instr_order.get(c[:2], 99), c[2:]))


def fetch_all_history(contracts, start_date, end_date):
    """拉取所有合约在日期范围内的日线数据（含当日实时补丁）。"""
    import pandas as pd

    all_data = {}  # {contract_code: DataFrame}
    for code in contracts:
        try:
            # akshare futures_zh_daily_sina 已坏，改用新浪期货日 K 线
            df = fetch_sina_futures_kline(code, start_date, end_date)
            if df is None or df.empty:
                print(f"  [SKIP] {code}: 无数据")
                continue
            df["date"] = df["date"].astype(str)
            mask = (df["date"] >= start_date) & (df["date"] <= end_date)
            df = df[mask].copy()
            if not df.empty:
                all_data[code] = df
                print(f"  [OK] {code}: {len(df)} 条记录 ({df['date'].iloc[0]} ~ {df['date'].iloc[-1]})")
            else:
                print(f"  [SKIP] {code}: 日期范围内无数据")
        except Exception as e:
            print(f"  [ERR] {code}: {e}")

    # 补充当日数据：新浪日 K 线接口在收盘后 1-2 小时内对部分新合约（如 2609/2610）
    # 暂未回填当日，故用 CFF_RE_ 实时接口补当日（与 fetch_close_prices 同一数据源，零 akshare）。
    # 注意：日线失败的合约即使实时有当日数据也必须补（新建记录），
    # 否则"当月"合约（如 IC2609/IM2609）缺失会导致期货/ETF 比率错用远月。
    try:
        from fetch_close_prices import fetch_all
        for r in fetch_all():
            code = r["code"]
            trade_date = r.get("trade_date", "")
            close_v = r.get("close")
            volume_v = r.get("volume")
            if not trade_date or close_v is None or volume_v is None or volume_v == 0:
                continue

            new_row = pd.DataFrame([{
                "date": trade_date,
                "open": r.get("open"),
                "high": r.get("high"),
                "low": r.get("low"),
                "close": close_v,
                "volume": volume_v,
                "hold": r.get("open_interest"),
                "settle": r.get("settle"),
            }])

            if code in all_data:
                existing_dates = set(all_data[code]["date"].values)
                if trade_date in existing_dates:
                    continue
                all_data[code] = pd.concat([all_data[code], new_row], ignore_index=True)
            else:
                # 日线缺失的新合约：用实时当日数据新建
                all_data[code] = new_row
            print(f"  [+] {code}: 补充当日 {trade_date} (Sina realtime)")
    except Exception as e:
        print(f"  [!] 当日 Sina 实时补充失败: {e}")

    return all_data


def build_summary_sheet(wb, all_data, start_date, end_date):
    """构建汇总 sheet：每个品种一行，列=日期，值=当月合约收盘价。"""
    ws = wb.active
    ws.title = "当月合约收盘价"

    # 收集所有交易日
    all_dates = set()
    for df in all_data.values():
        all_dates.update(df["date"].tolist())
    all_dates = sorted(all_dates)

    # 标题
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(all_dates) + 1)
    title_cell = ws.cell(row=1, column=1, value=f"CFFEX 股指期货当月合约收盘价  |  {start_date} ~ {end_date}")
    title_cell.font = TITLE_FONT
    title_cell.alignment = Alignment(horizontal="center")

    # 表头
    headers = ["品种"] + all_dates
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")
        cell.border = THIN_BORDER

    # 每品种找当月合约
    for row_idx, inst in enumerate(INSTRUMENTS, 3):
        ws.cell(row=row_idx, column=1, value=f"{inst} ({INSTRUMENT_NAMES[inst]})").font = DATA_FONT
        ws.cell(row=row_idx, column=1).border = THIN_BORDER

        # 找该品种最早到期月份
        inst_codes = sorted(
            [c for c in all_data if c.startswith(inst)],
            key=lambda c: c[2:],
        )
        if not inst_codes:
            continue
        front_code = inst_codes[0]
        df = all_data[front_code].set_index("date")

        for col, d in enumerate(all_dates, 2):
            cell = ws.cell(row=row_idx, column=col)
            cell.font = DATA_FONT
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="right")
            if d in df.index:
                cell.value = round(float(df.loc[d, "close"]), 1)

    # 列宽
    ws.column_dimensions["A"].width = 18
    for col in range(2, len(all_dates) + 2):
        ws.column_dimensions[get_column_letter(col)].width = 10

    return ws


def build_detail_sheet(wb, all_data, inst):
    """构建单品种明细 sheet：行为日期，列为各合约 OHLCV。"""
    ws = wb.create_sheet(title=f"{inst} 全合约")

    inst_codes = sorted([c for c in all_data if c.startswith(inst)], key=lambda c: c[2:])
    if not inst_codes:
        ws.cell(row=1, column=1, value="无数据")
        return ws

    # 收集所有日期
    all_dates = set()
    for code in inst_codes:
        all_dates.update(all_data[code]["date"].tolist())
    all_dates = sorted(all_dates)

    # 标题
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(inst_codes) * 5 + 1)
    title_cell = ws.cell(row=1, column=1, value=f"{inst} — {INSTRUMENT_NAMES[inst]}  各合约行情")
    title_cell.font = TITLE_FONT
    title_cell.alignment = Alignment(horizontal="center")

    # 表头第一行：合约名
    ws.cell(row=2, column=1, value="日期").font = HEADER_FONT
    ws.cell(row=2, column=1).fill = HEADER_FILL
    ws.cell(row=2, column=1).border = THIN_BORDER
    ws.cell(row=2, column=1).alignment = Alignment(horizontal="center")

    col = 2
    for code in inst_codes:
        ws.merge_cells(start_row=2, start_column=col, end_row=2, end_column=col + 4)
        hcell = ws.cell(row=2, column=col, value=code)
        hcell.font = HEADER_FONT
        hcell.fill = PatternFill(start_color=COLORS.get(inst, "666666"), end_color=COLORS.get(inst, "666666"), fill_type="solid")
        hcell.alignment = Alignment(horizontal="center")
        for c in range(col, col + 5):
            ws.cell(row=2, column=c).border = THIN_BORDER
        col += 5

    # 表头第二行：字段名
    ws.cell(row=3, column=1, value="").border = THIN_BORDER
    col = 2
    for code in inst_codes:
        for field in ["开盘", "最高", "最低", "收盘", "结算"]:
            cell = ws.cell(row=3, column=col, value=field)
            cell.font = Font(name="微软雅黑", bold=True, size=9)
            cell.fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
            cell.alignment = Alignment(horizontal="center")
            cell.border = THIN_BORDER
            col += 1

    # 数据行
    for row_idx, d in enumerate(all_dates, 4):
        date_cell = ws.cell(row=row_idx, column=1, value=d)
        date_cell.font = DATA_FONT
        date_cell.border = THIN_BORDER
        date_cell.alignment = Alignment(horizontal="center")

        col = 2
        for code in inst_codes:
            df = all_data[code].set_index("date")
            if d in df.index:
                row = df.loc[d]
                vals = [
                    round(float(row["open"]), 1),
                    round(float(row["high"]), 1),
                    round(float(row["low"]), 1),
                    round(float(row["close"]), 1),
                    round(float(row["settle"]), 1) if row.get("settle") and float(row["settle"]) != 0 else None,
                ]
            else:
                vals = [None] * 5

            for v in vals:
                cell = ws.cell(row=row_idx, column=col)
                cell.font = DATA_FONT
                cell.border = THIN_BORDER
                cell.alignment = Alignment(horizontal="right")
                if v is not None:
                    cell.value = v
                col += 1

    # 列宽
    ws.column_dimensions["A"].width = 12
    for col in range(2, len(inst_codes) * 5 + 2):
        ws.column_dimensions[get_column_letter(col)].width = 9

    return ws


def build_volume_oi_sheet(wb, all_data, inst):
    """构建成交量和持仓量 sheet。"""
    ws = wb.create_sheet(title=f"{inst} 量仓")

    inst_codes = sorted([c for c in all_data if c.startswith(inst)], key=lambda c: c[2:])
    if not inst_codes:
        return ws

    all_dates = set()
    for code in inst_codes:
        all_dates.update(all_data[code]["date"].tolist())
    all_dates = sorted(all_dates)

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(inst_codes) * 2 + 1)
    ws.cell(row=1, column=1, value=f"{inst} — {INSTRUMENT_NAMES[inst]}  成交量 / 持仓量").font = TITLE_FONT
    ws.cell(row=1, column=1).alignment = Alignment(horizontal="center")

    ws.cell(row=2, column=1, value="日期").font = HEADER_FONT
    ws.cell(row=2, column=1).fill = HEADER_FILL
    ws.cell(row=2, column=1).border = THIN_BORDER
    ws.cell(row=2, column=1).alignment = Alignment(horizontal="center")

    col = 2
    for code in inst_codes:
        ws.merge_cells(start_row=2, start_column=col, end_row=2, end_column=col + 1)
        hcell = ws.cell(row=2, column=col, value=code)
        hcell.font = HEADER_FONT
        hcell.fill = PatternFill(start_color=COLORS.get(inst, "666666"), end_color=COLORS.get(inst, "666666"), fill_type="solid")
        hcell.alignment = Alignment(horizontal="center")
        for c in range(col, col + 2):
            ws.cell(row=2, column=c).border = THIN_BORDER
        col += 2

    ws.cell(row=3, column=1, value="").border = THIN_BORDER
    col = 2
    for code in inst_codes:
        for field in ["量", "仓"]:
            cell = ws.cell(row=3, column=col, value=field)
            cell.font = Font(name="微软雅黑", bold=True, size=9)
            cell.fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
            cell.alignment = Alignment(horizontal="center")
            cell.border = THIN_BORDER
            col += 1

    for row_idx, d in enumerate(all_dates, 4):
        ws.cell(row=row_idx, column=1, value=d).font = DATA_FONT
        ws.cell(row=row_idx, column=1).border = THIN_BORDER
        ws.cell(row=row_idx, column=1).alignment = Alignment(horizontal="center")

        col = 2
        for code in inst_codes:
            df = all_data[code].set_index("date")
            if d in df.index:
                row = df.loc[d]
                for field in ["volume", "hold"]:
                    cell = ws.cell(row=row_idx, column=col)
                    cell.font = DATA_FONT
                    cell.border = THIN_BORDER
                    cell.alignment = Alignment(horizontal="right")
                    try:
                        cell.value = int(float(row[field]))
                        cell.number_format = "#,##0"
                    except (ValueError, TypeError):
                        pass
                    col += 1
            else:
                for _ in range(2):
                    ws.cell(row=row_idx, column=col).border = THIN_BORDER
                    col += 1

    ws.column_dimensions["A"].width = 12
    for col in range(2, len(inst_codes) * 2 + 2):
        ws.column_dimensions[get_column_letter(col)].width = 11

    return ws


def build_etf_sheet(wb, start_date, end_date, create_sheet=True):
    """构建 ETF 收盘价 sheet（或仅获取数据）。"""
    ws = wb.create_sheet(title="ETF收盘价") if create_sheet else None

    etf_map = {
        "510050": "IH_上证50ETF",
        "510300": "IF_沪深300ETF",
        "510500": "IC_中证500ETF",
        "512100": "IM_中证1000ETF",
    }

    all_dates = set()
    etf_data = {}
    for code, name in etf_map.items():
        df = None

        # 新浪 K 线接口（唯一 ETF 数据源；东方财富 push2his 已弃用/封锁）
        try:
            from urllib.request import Request, urlopen
            import json as _json
            import pandas as pd

            sina_code = f"sh{code}"
            url = (
                "https://money.finance.sina.com.cn/quotes_service/api/json_v2.php/"
                f"CN_MarketData.getKLineData?symbol={sina_code}&scale=240&ma=no&datalen=600"
            )
            req = Request(url, headers={"User-Agent": "Mozilla/5.0", "Referer": "https://finance.sina.com.cn/"})
            with urlopen(req, timeout=15) as resp:
                raw = _json.loads(resp.read().decode())

            rows = []
            for item in raw:
                d = item.get("day", "")
                if start_date <= d <= end_date:
                    rows.append({
                        "date": d,
                        "close": float(item["close"]),
                        "open": float(item["open"]),
                        "high": float(item["high"]),
                        "low": float(item["low"]),
                        "volume": float(item["volume"]),
                    })
            if rows:
                df = pd.DataFrame(rows)
        except Exception:
            pass

        if df is not None and not df.empty:
            all_dates.update(df["date"].tolist())
            etf_data[code] = df.set_index("date")
            print(f"  [OK] ETF {code} ({name}): {len(df)} 条 ({df['date'].iloc[0]} ~ {df['date'].iloc[-1]})")
        else:
            print(f"  [ERR] ETF {code}: 新浪 K 线获取失败")

    # 补充当日 ETF 数据：
    # 新浪 K 线可能不含当日（收盘前拉取），用 Sina 实时行情补（仅当日，>0 才采用）。
    from datetime import date as _dt
    today_str = _dt.today().strftime("%Y-%m-%d")

    def _today_etf_close_ok(code):
        if code not in etf_data or etf_data[code] is None or etf_data[code].empty:
            return False
        if today_str not in etf_data[code].index:
            return False
        try:
            v = float(etf_data[code].loc[today_str, "close"])
        except Exception:
            return False
        return v > 0

    try:
        from urllib.request import Request, urlopen
        import re
        import pandas as pd
        sina_map = {"510050": "sh510050", "510300": "sh510300", "510500": "sh510500", "512100": "sh512100"}
        name = {"510050": "IH_上证50ETF", "510300": "IF_沪深300ETF", "510500": "IC_中证500ETF", "512100": "IM_中证1000ETF"}
        # 仅对当日收盘无效的 ETF 补 Sina 实时（缺失 / 0 占位都补）
        need = [c for c in sina_map if not _today_etf_close_ok(c)]
        if need:
            url = "https://hq.sinajs.cn/list=" + ",".join(sina_map[c] for c in need)
            req = Request(url, headers={"Referer": "https://finance.sina.com.cn/", "User-Agent": "Mozilla/5.0"})
            with urlopen(req, timeout=15) as resp:
                text = resp.read().decode("gbk", errors="replace")
            rev = {v: k for k, v in sina_map.items()}
            for m in re.finditer(r'var hq_str_(\w+)="([^"]*)"', text):
                s_code = m.group(1)
                if s_code not in rev:
                    continue
                etf_code = rev[s_code]
                if etf_code not in need:
                    continue
                vals = m.group(2).split(",")
                if len(vals) < 4:
                    continue
                try:
                    close_v = float(vals[3])
                except (ValueError, TypeError):
                    continue
                if close_v <= 0:
                    continue
                new_row = pd.DataFrame([{"date": today_str, "close": close_v, "open": close_v, "high": close_v, "low": close_v, "volume": 0}]).set_index("date")
                if etf_code not in etf_data or etf_data[etf_code] is None or etf_data[etf_code].empty:
                    etf_data[etf_code] = new_row
                else:
                    # 丢弃当日旧值，用 Sina 实时覆盖当日
                    etf_data[etf_code] = etf_data[etf_code][etf_data[etf_code].index != today_str]
                    etf_data[etf_code] = pd.concat([etf_data[etf_code], new_row])
                all_dates.add(today_str)
                print(f"  [+] ETF {etf_code} ({name.get(etf_code, '')}): 补充/修正当日 {today_str} (Sina realtime) close={close_v}")
    except Exception as e:
        print(f"  [!] ETF 当日补丁失败: {e}")

    all_dates = sorted(all_dates)

    if ws is not None:
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(etf_data) + 1)
        ws.cell(row=1, column=1, value=f"对应 ETF 收盘价  |  {start_date} ~ {end_date}").font = TITLE_FONT
        ws.cell(row=1, column=1).alignment = Alignment(horizontal="center")

        ws.cell(row=2, column=1, value="日期").font = HEADER_FONT
        ws.cell(row=2, column=1).fill = HEADER_FILL
        ws.cell(row=2, column=1).border = THIN_BORDER
        ws.cell(row=2, column=1).alignment = Alignment(horizontal="center")

        etf_order = ["510050", "510300", "510500", "512100"]
        for col, code in enumerate(etf_order, 2):
            cell = ws.cell(row=2, column=col, value=f"{code} ({etf_map[code]})")
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="center")

        for row_idx, d in enumerate(all_dates, 3):
            ws.cell(row=row_idx, column=1, value=d).font = DATA_FONT
            ws.cell(row=row_idx, column=1).border = THIN_BORDER
            ws.cell(row=row_idx, column=1).alignment = Alignment(horizontal="center")

            for col, code in enumerate(etf_order, 2):
                cell = ws.cell(row=row_idx, column=col)
                cell.font = DATA_FONT
                cell.border = THIN_BORDER
                cell.alignment = Alignment(horizontal="right")
                if code in etf_data and d in etf_data[code].index:
                    try:
                        cell.value = round(float(etf_data[code].loc[d, "close"]), 3)
                    except Exception:
                        pass

        ws.column_dimensions["A"].width = 12
        for col in range(2, len(etf_order) + 2):
            ws.column_dimensions[get_column_letter(col)].width = 18

    return ws, etf_data


# ─────────────────────────────────────────────────────────────
# ETF 期权历史 + 合成期货 F
# ─────────────────────────────────────────────────────────────

SINA_OPT_RT_REF = "https://finance.sina.com.cn/"
SINA_OPT_HIST_REF = "https://stock.finance.sina.com.cn/option/quotes.html"

# 合成期货 F = K + (C - P) 需要 ETF 期权行情。
# 512100(IM/中证1000ETF) 在新浪期权体系无行情数据，IM 的合成F不可用，故跳过。
_OPT_ETF_CODES = ["510050", "510300", "510500"]


def _opt_get(url, ref, timeout=15):
    from urllib.request import Request, urlopen
    req = Request(url, headers={"User-Agent": "Mozilla/5.0", "Referer": ref})
    with urlopen(req, timeout=timeout) as r:
        return r.read().decode("gbk", "replace")


def _parse_opt_name(name):
    """从合约简称解析 (行权价K, 类型)。如 '50ETF购7月2700' -> (2.7, 'C')。"""
    if not name:
        return None, None
    typ = "C" if "购" in name else ("P" if ("沽" in name or "认沽" in name) else None)
    if typ is None:
        return None, None
    m = re.search(r"(\d+)\s*$", name)
    if not m:
        return None, None
    return float(m.group(1)) / 1000.0, typ


def fetch_etf_option_history(start_date, end_date, max_retry=6):
    """拉取 4 个 ETF 标的的期权历史日线。

    返回 {etf_code: {contract_code: {"K": float, "type": "C"/"P", "df": DataFrame}}}。
    合成期货 F = K + (C - P)，取每日最平值(ATM, K 最接近 ETF 收盘)的认购/认沽对计算。
    """
    import pandas as pd

    s = date.fromisoformat(start_date)
    e = date.fromisoformat(end_date)

    # 期权 ETF 到期日 = 每月第 4 个周三。
    # 只抓取在 [start, end] 内可能作为「当月合约」的期权月份：
    #   - 从起始月（其到期日通常 >= start）
    #   - 直到「到期日覆盖 end 的那个月」为止（更往后的月份在范围内永远不是当月，纯浪费）
    def _opt_exp(yy, mm):
        y = 2000 + yy
        first = date(y, mm, 1)
        days_to_first_wed = (2 - first.weekday()) % 7  # 周三 = 2
        first_wed = first + timedelta(days=days_to_first_wed)
        return first_wed + timedelta(days=21)  # 第 4 个周三

    # 覆盖 end_date 所需的最晚期权月份（到期日 >= end 的那个月）
    ey, em = e.year % 100, e.month
    if _opt_exp(ey, em) < e:
        em += 1
        if em > 12:
            em = 1
            ey += 1

    months = []
    y, m = s.year, s.month
    while True:
        if (y % 100, m) > (ey, em):
            break
        months.append(f"{y % 100:02d}{m:02d}")
        m += 1
        if m > 12:
            m = 1
            y += 1

    # 取 ETF 当前收盘价（用于挑选平值 ATM 档），按 etf_code 缓存，整个函数只请求一次。
    _ref_cache = {}
    # 取 ETF 在 [start,end] 内逐日收盘序列（用于「逐日」挑选平值 ATM 档），按 etf_code 缓存。
    _series_cache = {}

    def _etf_ref_close(code):
        if code in _ref_cache:
            return _ref_cache[code]
        price = None
        try:
            from urllib.request import Request, urlopen
            import json as _json
            sina_code = f"sh{code}"
            url = (
                "https://money.finance.sina.com.cn/quotes_service/api/json_v2.php/"
                f"CN_MarketData.getKLineData?symbol={sina_code}&scale=240&ma=no&datalen=5"
            )
            req = Request(url, headers={"User-Agent": "Mozilla/5.0", "Referer": "https://finance.sina.com.cn/"})
            with urlopen(req, timeout=15) as resp:
                raw = _json.loads(resp.read().decode())
            if raw:
                price = float(raw[-1].get("close", 0)) or None
        except Exception:
            price = None
        _ref_cache[code] = price
        return price

    def _etf_close_series(code):
        """取该 ETF 在 [start_date, end_date] 内逐日收盘序列 {date: close}，按 etf_code 缓存。"""
        if code in _series_cache:
            return _series_cache[code]
        import json as _json
        from urllib.request import Request, urlopen
        sina_code = f"sh{code}"
        url = (
            "https://money.finance.sina.com.cn/quotes_service/api/json_v2.php/"
            f"CN_MarketData.getKLineData?symbol={sina_code}&scale=240&ma=no&datalen=600"
        )
        out = {}
        try:
            req = Request(url, headers={"User-Agent": "Mozilla/5.0", "Referer": "https://finance.sina.com.cn/"})
            with urlopen(req, timeout=15) as resp:
                arr = _json.loads(resp.read().decode())
            for it in arr:
                d = it.get("day", "")
                if start_date <= d <= end_date:
                    try:
                        out[d] = float(it.get("close", 0))
                    except (ValueError, TypeError):
                        pass
        except Exception:
            out = {}
        _series_cache[code] = out
        return out

    def _front_month(d_str):
        """给定交易日，返回 months 列表中「到期日 >= 该日」的最近月（与 _compute_synthetic_f 的当月口径一致）。"""
        dd = date.fromisoformat(d_str)
        for m in months:
            if _opt_exp(int(m[:2]), int(m[2:])) >= dd:
                return m
        return months[-1] if months else None

    opt_data = {}
    for etf_code in _OPT_ETF_CODES:
        contracts = {}
        for mm in months:
            for side, pre in [("C", "OP_UP"), ("P", "OP_DOWN")]:
                # 1) 枚举该月该方向合约代码
                codes = []
                for attempt in range(max_retry):
                    try:
                        t = _opt_get(f"https://hq.sinajs.cn/list={pre}_{etf_code}{mm}", SINA_OPT_RT_REF)
                        inner = t.split('"')[1] if '"' in t else ""
                        if inner:
                            codes = [c for c in inner.split(",") if c]
                        if codes:
                            break
                    except Exception:
                        if attempt == max_retry - 1:
                            print(f"  [WARN] 枚举 {etf_code} {mm} {side} 失败")
                    # 限流/空结果：退避后重试（不要直接 break 丢弃整月）
                    if not codes and attempt < max_retry - 1:
                        time.sleep(1.5 * (attempt + 1))
                if not codes:
                    continue

                # 2) 批量拿合约简称(含行权价/类型)，分批 50
                meta = {}
                for i in range(0, len(codes), 50):
                    batch = codes[i:i + 50]
                    ok = False
                    for attempt in range(max_retry):
                        try:
                            t = _opt_get("https://hq.sinajs.cn/list=" + ",".join(batch), SINA_OPT_RT_REF)
                            found_any = False
                            for mt in re.finditer(r'var hq_str_(\w+)="([^"]*)"', t):
                                name_field = None
                                for fv in mt.group(2).split(","):
                                    if "购" in fv or "沽" in fv or "认沽" in fv:
                                        name_field = fv
                                        break
                                if name_field is None:
                                    continue
                                K, typ = _parse_opt_name(name_field)
                                if K is not None:
                                    meta[mt.group(1)] = (K, typ)
                                    found_any = True
                            if found_any:
                                ok = True
                                break
                        except Exception:
                            if attempt == max_retry - 1:
                                print(f"  [WARN] 合约简称 {etf_code} {mm} 批量失败")
                        # 限流/空结果：退避后重试，否则整批合约将丢失 meta 被跳过
                        if not ok and attempt < max_retry - 1:
                            time.sleep(1.5 * (attempt + 1))
                    if not ok:
                        print(f"  [WARN] 合约简称 {etf_code} {mm} 批量(空)")

                # 2.5) 逐日挑选平值(ATM)档并只抓取这些档的日线（仍远少于全月）。
                # ATM 行权价 = 当天「当月合约」行权价列表中、与当天 ETF 收盘最接近者（|K - 当天ETF价| 最小）。
                # 关键改进：改为「逐交易日」选档 —— 现货在范围内漂移时，不同日期会锁定到不同平值档，
                # 而不是用单一参考价（最新收盘）锁定一档套用全部日期。
                # 由 Put-Call Parity，F=K+(C-P) 几乎与行权价无关（不同平值档 F 仅差千分之几），
                # 故只抓 ATM 档即可，请求量远低于全月抓取。
                # 仅对「当月为 mm」的那些日期计算 ATM，避免为远期月在无关日期浪费抓取。
                if meta:
                    ks = sorted({info[0] for info in meta.values()})
                    series = _etf_close_series(etf_code)
                    atm_set = set()
                    if series:
                        for d, price in series.items():
                            if price is not None and _front_month(d) == mm:
                                atm_set.add(min(ks, key=lambda k: abs(k - price)))
                    if not atm_set:
                        ref = _etf_ref_close(etf_code)
                        atm_K = min(ks, key=lambda k: abs(k - ref)) if ref is not None else ks[len(ks) // 2]
                        atm_set.add(atm_K)
                    fetch_codes = [c for c in codes if c in meta and meta[c][0] in atm_set]
                else:
                    fetch_codes = []

                # 3) 逐合约拉日线历史（仅 ATM 一对，由 _compute_synthetic_f 计算 F）
                for ccode in fetch_codes:
                    if ccode not in meta or ccode in contracts:
                        continue
                    K, typ = meta[ccode]
                    df = None
                    for attempt in range(max_retry):
                        try:
                            t = _opt_get(
                                f"https://stock.finance.sina.com.cn/futures/api/jsonp_v2.php//StockOptionDaylineService.getSymbolInfo?symbol={ccode}",
                                SINA_OPT_HIST_REF,
                            )
                            ss = t.find("(")
                            ee = t.rfind(")")
                            arr = json.loads(t[ss + 1:ee]) if ss >= 0 and ee > 0 else json.loads(t)
                            rows = []
                            for it in arr:
                                d = it.get("d", "")
                                if start_date <= d <= end_date:
                                    try:
                                        rows.append({"date": d, "close": float(it["c"])})
                                    except (ValueError, KeyError):
                                        pass
                            if rows:
                                df = pd.DataFrame(rows)
                                break
                            # rows 为空（Sina 限流常返回空数组）：退避后重试，不要丢弃合约
                        except Exception:
                            if attempt == max_retry - 1:
                                print(f"  [WARN] 日线 {ccode} 失败")
                        if df is None and attempt < max_retry - 1:
                            time.sleep(1.5 * (attempt + 1))
                    if df is not None and not df.empty:
                        contracts[ccode] = {"K": K, "type": typ, "df": df, "month": mm}
                    # 礼貌间隔，降低触发限流概率
                    time.sleep(0.1)
        if contracts:
            opt_data[etf_code] = contracts
            print(f"  [OK] ETF期权 {etf_code}: {len(contracts)} 合约")
        else:
            print(f"  [ERR] ETF期权 {etf_code}: 无数据")
    return opt_data


def _compute_synthetic_f(opt_data, etf_sheet_data):
    """计算合成期货 F = K + (C - P)，逐日取最平值(ATM)合约对。返回 {etf_code: {date: F}}。"""
    if not opt_data:
        return {}

    etf_close = {}  # etf_code -> {date: close}
    for code, df in etf_sheet_data.items():
        if df is not None and not df.empty:
            for d, row in df.iterrows():
                try:
                    etf_close.setdefault(code, {})[str(d)] = float(row["close"])
                except Exception:
                    pass

    # etf_code -> date -> month(YYMM) -> K -> {C:, P:}
    # 增加 month 维度，便于强制只用「当月合约」算合成F（排除远月干扰）
    by = {}
    for etf_code, contracts in opt_data.items():
        for ccode, info in contracts.items():
            K, typ, df, mm = info["K"], info["type"], info["df"], info["month"]
            for _, row in df.iterrows():
                d = str(row["date"])
                try:
                    c = float(row["close"])
                except Exception:
                    continue
                by.setdefault(etf_code, {}).setdefault(d, {}).setdefault(mm, {}).setdefault(K, {})[typ] = c

    result = {}
    for etf_code in opt_data:
        result[etf_code] = {}
        for d, mmap in by.get(etf_code, {}).items():
            e = etf_close.get(etf_code, {}).get(d)
            if e is None:
                continue
            # 当月 = 当天有报价合约中月份最小者（与期货近月逻辑一致）
            # 不用自然月 YYMM 硬匹配：期权当月合约在到期后切换，自然月会错位导致缺口
            yymm = min(mmap.keys())
            kmap = mmap[yymm]
            best, bestdiff = None, 1e18
            for K, cp in kmap.items():
                if "C" in cp and "P" in cp:
                    diff = abs(K - e)
                    if diff < bestdiff:
                        bestdiff = diff
                        best = K
            if best is not None:
                cp = kmap[best]
                result[etf_code][d] = round(best + cp["C"] - cp["P"], 4)
    return result


def _third_friday(year, month):
    """返回指定年月的第三个周五（股指期货交割日）。"""
    first = date(year, month, 1)
    # 周五 weekday=4，计算到第一个周五的天数
    days = (4 - first.weekday()) % 7
    return first + timedelta(days=days + 14)


def _get_third_fridays(start_str, end_str):
    """返回 start~end 范围内所有第三个周五的日期字符串集合。"""
    s = date.fromisoformat(start_str)
    e = date.fromisoformat(end_str)
    result = set()
    y, m = s.year, s.month
    while True:
        tf = _third_friday(y, m)
        if tf > e:
            break
        if tf >= s:
            result.add(tf.isoformat())
        m += 1
        if m > 12:
            m = 1
            y += 1
    return result


def _compute_ratio_data(all_data, etf_sheet_data, synthetic_f=None):
    """计算所有品种当月合约的 期货/ETF 与 期货/合成F 比率。
    返回 (ratio_data, dates, ratio_data_sf)。synthetic_f 为 {etf_code:{date:F}}。"""
    etf_map = {
        "IH": "510050",
        "IF": "510300",
        "IC": "510500",
        "IM": "512100",
    }

    # 收集所有期货交易日
    all_fut_dates = set()
    for df in all_data.values():
        all_fut_dates.update(df["date"].tolist())
    all_fut_dates = sorted(all_fut_dates)

    # 预处理 ETF close
    etf_close = {}
    for code, df in etf_sheet_data.items():
        if df is not None and not df.empty:
            for d, row in df.iterrows():
                k = str(d)
                if k not in etf_close:
                    etf_close[k] = {}
                try:
                    etf_close[k][code] = float(row["close"])
                except Exception:
                    pass

    # 合成期货 F：date -> etf_code -> F
    sf_by_date = {}
    if synthetic_f:
        for etf_code, dmap in synthetic_f.items():
            for d, fv in dmap.items():
                sf_by_date.setdefault(d, {})[etf_code] = fv

    # 预处理期货
    fut_by_inst_date = {}
    for code, df in all_data.items():
        inst = code[:2]
        month = code[2:]
        if inst not in fut_by_inst_date:
            fut_by_inst_date[inst] = {}
        for _, row in df.iterrows():
            d = str(row["date"])
            try:
                close = float(row["close"])
            except Exception:
                continue
            if d not in fut_by_inst_date[inst]:
                fut_by_inst_date[inst][d] = []
            fut_by_inst_date[inst][d].append((month, close))

    # 算比率
    ratio_data = {}
    ratio_data_sf = {}
    for d in all_fut_dates:
        ratio_data[d] = {}
        ratio_data_sf[d] = {}
        for inst in INSTRUMENTS:
            contracts = fut_by_inst_date.get(inst, {}).get(d, [])
            if not contracts:
                continue
            contracts.sort(key=lambda x: x[0])
            front_close = contracts[0][1]
            etf_code = etf_map[inst]
            e_close = etf_close.get(d, {}).get(etf_code)
            if e_close and e_close > 0:
                ratio_data[d][inst] = round(front_close / e_close, 1)
            # 期货/合成F
            fv = sf_by_date.get(d, {}).get(etf_code)
            if fv and fv > 0:
                ratio_data_sf[d][inst] = round(front_close / fv, 1)

    dates = sorted(ratio_data.keys())
    return ratio_data, dates, ratio_data_sf


def _build_single_ratio_sheet(wb, ratio_data, dates, inst, start_date, end_date, ratio_data_sf=None):
    """为单个品种构建「期货/ETF比率」sheet，含折线图 + 交割日标注。"""
    instr_names = {"IH": "上证50", "IF": "沪深300", "IC": "中证500", "IM": "中证1000"}
    chart_colors = {
        "IH": "4472C4",
        "IF": "ED7D31",
        "IC": "70AD47",
        "IM": "FFC000",
    }

    ws = wb.create_sheet(title=f"{inst}_期货ETF比率")
    name = instr_names.get(inst, inst)

    # 该品种是否有合成F数据（IM/512100 无期权行情 → 无，跳过期货/合成F列与图）
    inst_has_sf = bool(ratio_data_sf) and any(inst in ratio_data_sf.get(d, {}) for d in dates)

    # ── 计算交割日（含假日顺延） ──
    delivery_dates = _get_third_fridays(start_date, end_date)
    # 若第三个周五为非交易日，顺延到下一个交易日
    effective_delivery = set()
    dates_set = set(dates)
    for dd in sorted(delivery_dates):
        if dd in dates_set:
            effective_delivery.add(dd)
        else:
            # 找到 dates 中紧接其后的第一个交易日
            from datetime import timedelta
            dd_obj = date.fromisoformat(dd)
            for offset in range(1, 10):
                candidate = (dd_obj + timedelta(days=offset)).isoformat()
                if candidate in dates_set:
                    effective_delivery.add(candidate)
                    break

    # ── 标题 ──
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=2)
    title_cell = ws.cell(row=1, column=1, value=f"{inst} ({name}) 当月合约 期货/ETF 比率  |  {start_date} ~ {end_date}")
    title_cell.font = TITLE_FONT
    title_cell.alignment = Alignment(horizontal="center")

    # ── 表头 ──
    headers = [(1, "日期"), (2, f"{inst} 期货/ETF")]
    if inst_has_sf:
        headers.append((3, f"{inst} 期货/合成F"))
    for col, label in headers:
        cell = ws.cell(row=2, column=col, value=label)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.border = THIN_BORDER
        cell.alignment = Alignment(horizontal="center")

    # ── 数据行 ──
    delivery_idxs = []  # 交割日在 dates 中的索引
    for row_idx, d in enumerate(dates, 3):
        ws.cell(row=row_idx, column=1, value=d).font = DATA_FONT
        ws.cell(row=row_idx, column=1).border = THIN_BORDER
        ws.cell(row=row_idx, column=1).alignment = Alignment(horizontal="center")

        cell = ws.cell(row=row_idx, column=2)
        cell.font = DATA_FONT
        cell.border = THIN_BORDER
        cell.alignment = Alignment(horizontal="right")
        cell.number_format = "0.0"
        if inst in ratio_data[d]:
            cell.value = ratio_data[d][inst]

        # 期货/合成F 列（仅该品种有合成F数据时填写）
        if inst_has_sf:
            cell3 = ws.cell(row=row_idx, column=3)
            cell3.font = DATA_FONT
            cell3.border = THIN_BORDER
            cell3.alignment = Alignment(horizontal="right")
            cell3.number_format = "0.0"
            if inst in ratio_data_sf.get(d, {}):
                cell3.value = ratio_data_sf[d][inst]

        if d in effective_delivery and inst in ratio_data[d]:
            delivery_idxs.append(row_idx - 3)

    # ── 列宽 ──
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 16
    if inst_has_sf:
        ws.column_dimensions["C"].width = 16

    # ── 折线图 ──
    chart = LineChart()
    chart.title = f"{inst} ({name}) 期货/ETF 比率"
    chart.style = 10
    chart.width = 26
    chart.height = 15
    chart.y_axis.title = "比率"

    chart.x_axis.tickLblSkip = 1
    chart.x_axis.tickMarkSkip = 1
    chart.x_axis.numFmt = "m/d"
    chart.x_axis.delete = False

    cats = Reference(ws, min_col=1, min_row=3, max_row=2 + len(dates))
    data_ref = Reference(ws, min_col=2, min_row=2, max_row=2 + len(dates))
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats)

    s = chart.series[0]
    s.graphicalProperties.line.width = 28000
    if inst in chart_colors:
        s.graphicalProperties.line.solidFill = chart_colors[inst]

    # 所有点加微小同色圆点（确保每个点都能悬停）
    from openpyxl.chart.marker import Marker
    color = chart_colors.get(inst, "4472C4")
    s.marker.symbol = "circle"
    s.marker.size = 4
    s.marker.graphicalProperties.solidFill = color

    # 交割日点覆盖为大红色圆点
    for idx in delivery_idxs:
        pt = DataPoint(idx=idx)
        m = Marker(symbol="circle", size=9)
        m.graphicalProperties.solidFill = "FF0000"
        pt.marker = m
        s.data_points.append(pt)

    chart.legend.position = "b"
    ws.add_chart(chart, "D3")

    # ── 第二折线图：期货/合成F（仅该品种有合成F数据时）──
    if inst_has_sf:
        chart2 = LineChart()
        chart2.title = f"{inst} ({name}) 期货/合成F 比率"
        chart2.style = 10
        chart2.width = 26
        chart2.height = 15
        chart2.y_axis.title = "比率"

        chart2.x_axis.tickLblSkip = 1
        chart2.x_axis.tickMarkSkip = 1
        chart2.x_axis.numFmt = "m/d"
        chart2.x_axis.delete = False

        cats2 = Reference(ws, min_col=1, min_row=3, max_row=2 + len(dates))
        data_ref2 = Reference(ws, min_col=3, min_row=2, max_row=2 + len(dates))
        chart2.add_data(data_ref2, titles_from_data=True)
        chart2.set_categories(cats2)

        s2 = chart2.series[0]
        s2.graphicalProperties.line.width = 28000
        if inst in chart_colors:
            s2.graphicalProperties.line.solidFill = chart_colors[inst]

        from openpyxl.chart.marker import Marker
        color = chart_colors.get(inst, "4472C4")
        s2.marker.symbol = "circle"
        s2.marker.size = 4
        s2.marker.graphicalProperties.solidFill = color

        for idx in delivery_idxs:
            pt = DataPoint(idx=idx)
            m = Marker(symbol="circle", size=9)
            m.graphicalProperties.solidFill = "FF0000"
            pt.marker = m
            s2.data_points.append(pt)

        chart2.legend.position = "b"
        ws.add_chart(chart2, "D20")

    return ws


def build_ratio_chart_sheets(wb, all_data, etf_sheet_data, start_date, end_date, synthetic_f=None):
    """为四个品种各建一个「期货/ETF比率」sheet，每个含 期货/ETF 与 期货/合成F 折线图。"""
    ratio_data, dates, ratio_data_sf = _compute_ratio_data(all_data, etf_sheet_data, synthetic_f)
    if not ratio_data:
        ws = wb.create_sheet(title="期货ETF比率")
        ws.cell(row=1, column=1, value="无数据")
        return {}

    for inst in INSTRUMENTS:
        _build_single_ratio_sheet(wb, ratio_data, dates, inst, start_date, end_date, ratio_data_sf=ratio_data_sf)
        print(f"  [OK] {inst}_期货ETF比率: {len(dates)} 个交易日, 图表已嵌入")

    return ratio_data


def _rebuild_ratio_charts(ws, inst, has_sf, full_dates, delivery_idxs):
    """清空并重建某比率 sheet 的折线图，引用范围扩展到 full_dates 长度。"""
    from openpyxl.chart.marker import Marker

    chart_colors = {"IH": "4472C4", "IF": "ED7D31", "IC": "70AD47", "IM": "FFC000"}
    color = chart_colors.get(inst, "4472C4")
    max_row = 2 + len(full_dates)

    # 清空旧图（增量追加后行数变多，旧图范围已过期）
    ws._charts = []

    chart = LineChart()
    chart.title = f"{inst} ({INSTRUMENT_NAMES.get(inst, inst)}) 期货/ETF 比率"
    chart.style = 10
    chart.width = 26
    chart.height = 15
    chart.y_axis.title = "比率"
    chart.x_axis.tickLblSkip = 1
    chart.x_axis.tickMarkSkip = 1
    chart.x_axis.numFmt = "m/d"
    chart.x_axis.delete = False
    cats = Reference(ws, min_col=1, min_row=3, max_row=max_row)
    data_ref = Reference(ws, min_col=2, min_row=2, max_row=max_row)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats)
    s = chart.series[0]
    s.graphicalProperties.line.width = 28000
    s.graphicalProperties.line.solidFill = color
    s.marker.symbol = "circle"
    s.marker.size = 4
    s.marker.graphicalProperties.solidFill = color
    for idx in delivery_idxs:
        pt = DataPoint(idx=idx)
        m = Marker(symbol="circle", size=9)
        m.graphicalProperties.solidFill = "FF0000"
        pt.marker = m
        s.data_points.append(pt)
    chart.legend.position = "b"
    ws.add_chart(chart, "D3")

    if has_sf:
        chart2 = LineChart()
        chart2.title = f"{inst} ({INSTRUMENT_NAMES.get(inst, inst)}) 期货/合成F 比率"
        chart2.style = 10
        chart2.width = 26
        chart2.height = 15
        chart2.y_axis.title = "比率"
        chart2.x_axis.tickLblSkip = 1
        chart2.x_axis.tickMarkSkip = 1
        chart2.x_axis.numFmt = "m/d"
        chart2.x_axis.delete = False
        cats2 = Reference(ws, min_col=1, min_row=3, max_row=max_row)
        data_ref2 = Reference(ws, min_col=3, min_row=2, max_row=max_row)
        chart2.add_data(data_ref2, titles_from_data=True)
        chart2.set_categories(cats2)
        s2 = chart2.series[0]
        s2.graphicalProperties.line.width = 28000
        s2.graphicalProperties.line.solidFill = color
        s2.marker.symbol = "circle"
        s2.marker.size = 4
        s2.marker.graphicalProperties.solidFill = color
        for idx in delivery_idxs:
            pt = DataPoint(idx=idx)
            m = Marker(symbol="circle", size=9)
            m.graphicalProperties.solidFill = "FF0000"
            pt.marker = m
            s2.data_points.append(pt)
        chart2.legend.position = "b"
        ws.add_chart(chart2, "D20")


def _incremental_update(output_path, today_str):
    """增量更新：读现有 xlsx，只补抓 last_date+1 ~ today 的新交易日，追加行并重建图表。
    返回输出路径表示成功（含“已是最新”）；返回 None 表示无法增量，调用方应回退全量重算。"""
    from openpyxl import load_workbook

    if not os.path.exists(output_path):
        return None

    wb = load_workbook(output_path)

    ratio_sheets = {
        inst: wb[f"{inst}_期货ETF比率"]
        for inst in INSTRUMENTS
        if f"{inst}_期货ETF比率" in wb.sheetnames
    }
    if not ratio_sheets:
        print("  [!] 现有文件无可增量更新的比率 sheet，转全量重算")
        return None

    # 读现有日期轴（取第一个 sheet 的 A 列，从 row3 起，连续非空）
    first_ws = next(iter(ratio_sheets.values()))
    existing_dates = []
    r = 3
    while True:
        v = first_ws.cell(row=r, column=1).value
        if v is None:
            break
        existing_dates.append(str(v))
        r += 1
    if not existing_dates:
        print("  [!] 现有比率 sheet 无数据行，转全量重算")
        return None

    last_date = date.fromisoformat(existing_dates[-1])
    new_start = last_date + timedelta(days=1)
    if new_start > date.fromisoformat(today_str):
        print(f"  [✓] 已是最新（末日 {existing_dates[-1]}），无需增量")
        return output_path

    print(f"\n📥 增量补抓 {new_start.isoformat()} ~ {today_str} ...")
    contracts = discover_active_contracts(min_last_date=new_start.isoformat())
    new_all = fetch_all_history(contracts, new_start.isoformat(), today_str)
    _, new_etf = build_etf_sheet(Workbook(), new_start.isoformat(), today_str, create_sheet=False)

    new_sf = {}
    try:
        opt = fetch_etf_option_history(new_start.isoformat(), today_str)
        new_sf = _compute_synthetic_f(opt, new_etf)
    except Exception as e:
        print(f"  [!] 合成F增量失败，仅补 期货/ETF: {e}")

    new_ratio, new_dates, new_ratio_sf = _compute_ratio_data(new_all, new_etf, new_sf)
    if not new_dates:
        print("  [✓] 新区间无期货交易日数据，无需增量")
        return output_path

    # 全量日期轴 + 交割日红点索引（覆盖现有+新，保证重建图表时旧红点不丢）
    full_dates = existing_dates + new_dates
    delivery_dates = _get_third_fridays(existing_dates[0], today_str)
    dates_set = set(full_dates)
    effective_delivery = set()
    for dd in sorted(delivery_dates):
        if dd in dates_set:
            effective_delivery.add(dd)
        else:
            for off in range(1, 10):
                cand = (date.fromisoformat(dd) + timedelta(days=off)).isoformat()
                if cand in dates_set:
                    effective_delivery.add(cand)
                    break
    delivery_idxs = [full_dates.index(d) for d in effective_delivery if d in full_dates]

    base_row = 3 + len(existing_dates)
    for inst, ws in ratio_sheets.items():
        has_sf_col = ws.cell(row=2, column=3).value is not None
        for i, d in enumerate(new_dates):
            rr = base_row + i
            c1 = ws.cell(row=rr, column=1, value=d)
            c1.font = DATA_FONT
            c1.border = THIN_BORDER
            c1.alignment = Alignment(horizontal="center")
            c2 = ws.cell(row=rr, column=2)
            c2.font = DATA_FONT
            c2.border = THIN_BORDER
            c2.alignment = Alignment(horizontal="right")
            c2.number_format = "0.0"
            if inst in new_ratio.get(d, {}):
                c2.value = new_ratio[d][inst]
            if has_sf_col:
                c3 = ws.cell(row=rr, column=3)
                c3.font = DATA_FONT
                c3.border = THIN_BORDER
                c3.alignment = Alignment(horizontal="right")
                c3.number_format = "0.0"
                if inst in new_ratio_sf.get(d, {}):
                    c3.value = new_ratio_sf[d][inst]
        _rebuild_ratio_charts(ws, inst, has_sf_col, full_dates, delivery_idxs)

    # 同步刷新 ETF收盘价 辅助 sheet（按新交易日追加）
    if "ETF收盘价" in wb.sheetnames:
        etf_ws = wb["ETF收盘价"]
        ed = []
        er = 3
        while True:
            v = etf_ws.cell(row=er, column=1).value
            if v is None:
                break
            ed.append(str(v))
            er += 1
        ebase = 3 + len(ed)
        etf_order = ["510050", "510300", "510500", "512100"]
        for i, d in enumerate(new_dates):
            rr = ebase + i
            cc = etf_ws.cell(row=rr, column=1, value=d)
            cc.font = DATA_FONT
            cc.border = THIN_BORDER
            cc.alignment = Alignment(horizontal="center")
            for col, code in enumerate(etf_order, 2):
                cell = etf_ws.cell(row=rr, column=col)
                cell.font = DATA_FONT
                cell.border = THIN_BORDER
                cell.alignment = Alignment(horizontal="right")
                if code in new_etf and d in new_etf[code].index:
                    try:
                        cell.value = round(float(new_etf[code].loc[d, "close"]), 3)
                    except Exception:
                        pass

    wb.save(output_path)
    print(f"  [✓] 增量完成：新增 {len(new_dates)} 个交易日（{new_dates[0]} ~ {new_dates[-1]}），已保存 {output_path}")
    return output_path


def main():
    parser = argparse.ArgumentParser(description="导出 CFFEX 股指期货历史数据到 Excel")
    parser.add_argument("--start", default="2026-06-01", help="起始日期 YYYY-MM-DD")
    parser.add_argument("--end", default=date.today().isoformat(), help="结束日期 YYYY-MM-DD")
    parser.add_argument("--output", default=None, help="输出 Excel 文件路径")
    parser.add_argument("--mode", default="auto", choices=["auto", "full", "incremental"],
                        help="auto=文件存在则增量否则全量; full=全量重算; incremental=强制增量(文件须已存在)")
    args = parser.parse_args()

    start_date = args.start
    end_date = args.end

    output_path = args.output or f"cffex_futures_{start_date}_{end_date}.xlsx"

    # 增量分派：文件已存在时默认增量，避免每天把整张表全量重算一遍
    if args.mode in ("auto", "incremental"):
        if args.mode == "incremental" and not os.path.exists(output_path):
            print("  [!] --mode incremental 但输出文件不存在，转全量重算")
        elif os.path.exists(output_path):
            res = _incremental_update(output_path, end_date)
            if res is not None:
                return res
            print("  [!] 增量失败，回退全量重算")

    # ===== 全量重算（首次 / --full / 增量不可用） =====
    print(f"\n📊 发现活跃合约...")
    contracts = discover_active_contracts(min_last_date=start_date)
    print(f"   共 {len(contracts)} 个合约: {', '.join(contracts)}\n")

    print(f"📥 拉取 {start_date} ~ {end_date} 历史数据...")
    all_data = fetch_all_history(contracts, start_date, end_date)

    if not all_data:
        print("❌ 没有任何数据")
        return 1

    print(f"\n📝 生成 Excel: {output_path}")
    wb = Workbook()

    # 获取 ETF 数据（不建 sheet）
    print(f"\n📥 拉取 ETF 数据...")
    _, etf_data = build_etf_sheet(wb, start_date, end_date, create_sheet=False)

    # 合成期货 F = K + (C - P)（新浪原生 ETF 期权接口，逐日取最平值合约对）
    print(f"\n📥 拉取 ETF 期权历史 + 计算合成期货 F...")
    synthetic_f = {}
    try:
        opt_data = fetch_etf_option_history(start_date, end_date)
        synthetic_f = _compute_synthetic_f(opt_data, etf_data)
        n = sum(len(v) for v in synthetic_f.values())
        print(f"   合成F覆盖: {n} 个 (etf_date) 数据点")
    except Exception as e:
        print(f"  [!] 合成F计算失败，仅生成 期货/ETF 列: {e}")

    # 删除默认空白 sheet
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    # 期货/ETF 比率图表（每个品种独立 sheet，含期货/合成F）
    print(f"\n📈 生成期货/ETF比率图表（每品种独立）...")
    build_ratio_chart_sheets(wb, all_data, etf_data, start_date, end_date, synthetic_f=synthetic_f)

    wb.save(output_path)
    print(f"\n✅ 已保存: {output_path}")
    print(f"   共 {len(wb.sheetnames)} 个工作表: {', '.join(wb.sheetnames)}")
    return output_path


if __name__ == "__main__":
    main()
