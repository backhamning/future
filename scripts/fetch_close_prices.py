#!/usr/bin/env python3
"""
CFFEX 股指期货收盘价获取脚本

获取 IH(上证50)/IC(中证500)/IF(沪深300)/IM(中证1000) 所有活跃合约的收盘价。

数据源：
  主：akshare (futures_zh_daily_sina) — 返回结构化 DataFrame，含完整 OHLCV + settle
  备：新浪 HTTP 直连 — 纯标准库，零外部依赖

用法：
    python fetch_close_prices.py              # 获取最新数据，输出表格
    python fetch_close_prices.py --csv        # 保存为 CSV 文件
    python fetch_close_prices.py --json       # JSON 格式输出
    python fetch_close_prices.py --output-dir ./data  # 指定输出目录
"""

import sys
import os
import json
import re
import csv
import warnings
from datetime import datetime, date

import pandas as pd

warnings.filterwarnings("ignore")

# ── 配置 ──────────────────────────────────────────────────────

INSTRUMENTS = ["IH", "IC", "IF", "IM"]
INSTRUMENT_NAMES = {
    "IH": "上证50",
    "IF": "沪深300",
    "IC": "中证500",
    "IM": "中证1000",
}

# 对应 ETF 映射
INSTRUMENT_ETF = {
    "IH": ("510050", "上证50ETF"),
    "IF": ("510300", "沪深300ETF"),
    "IC": ("510500", "中证500ETF"),
    "IM": ("512100", "中证1000ETF"),
}

# 中金所挂牌规则：当月、下月、随后两个季月
# 覆盖未来 12 个月确保不漏季月
CONTRACT_MONTHS_AHEAD = 12

# Sina HTTP 配置（备用方案）
SINA_URL = "https://hq.sinajs.cn/list="
SINA_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Referer": "https://finance.sina.com.cn/",
}

DEFAULT_OUTPUT_DIR = None


# ── 合约代码生成 ──────────────────────────────────────────────


def gen_contract_codes(year, month, months_ahead=CONTRACT_MONTHS_AHEAD):
    """生成候选合约代码列表（纯代码，无前缀）。"""
    codes = []
    for inst in INSTRUMENTS:
        for offset in range(months_ahead):
            m = month + offset
            y = year + (m - 1) // 12
            m = (m - 1) % 12 + 1
            codes.append(f"{inst}{y % 100:02d}{m:02d}")
    return codes


# ── 数据获取：akshare 实时 ──────────────────────────────────────


# akshare 实时接口品种名映射（中文品种名 → 代码前缀）
AKSHARE_REALTIME_SYMBOLS = {
    "IH": "上证50指数期货",
    "IF": "沪深300指数期货",
    "IC": "中证500指数期货",
    "IM": "中证1000股指期货",
}


def _has_akshare():
    """检查 akshare 是否可用。"""
    try:
        import akshare
        return True
    except ImportError:
        return False


def fetch_via_akshare_realtime():
    """
    通过 akshare futures_zh_realtime 获取各品种全部活跃合约的实时行情。

    优点：返回当日数据（收盘后立即可取），一次拉一个品种所有合约。
    限制：settlement 列在收盘后约 16:00 才有值，15:15 时为 0。

    Returns:
        list[dict]
    """
    import akshare as ak

    results = []
    for inst, cn_name in AKSHARE_REALTIME_SYMBOLS.items():
        try:
            df = ak.futures_zh_realtime(symbol=cn_name)
            if df is None or df.empty:
                continue

            for _, row in df.iterrows():
                symbol = str(row.get("symbol", ""))
                # 过滤连续合约（如 IH0, IF0），它是当月合约镜像
                if symbol.endswith("0") or len(symbol) <= 4:
                    continue

                code = symbol  # e.g. "IH2608"
                month = code[2:]  # e.g. "2608"

                close = _safe_float(row.get("close"))
                volume = _safe_int(row.get("volume"))

                # 过滤无效数据：close 为 None 或 0（收盘瞬间 akshare 可能返回 0）
                if close is None or close == 0 or volume is None or volume == 0:
                    continue

                settle = _safe_float(row.get("settlement"))
                # settlement 在收盘后不久为 0，用 presettlement 标记"待更新"
                if settle is not None and settle == 0.0:
                    settle = None

                record = {
                    "code": code,
                    "instrument": inst,
                    "month": month,
                    "name": str(row.get("name", "")),
                    "trade_date": str(row.get("tradedate", "")),
                    "open": _safe_float(row.get("open")),
                    "high": _safe_float(row.get("high")),
                    "low": _safe_float(row.get("low")),
                    "close": close,
                    "volume": volume,
                    "amount": None,
                    "open_interest": _safe_int(row.get("position")),
                    "settle": settle,
                    "source": "akshare_realtime",
                }

                results.append(record)

        except Exception as e:
            print(f"[WARN] akshare realtime 获取 {inst} 失败: {e}", file=sys.stderr)

    return results


def fetch_via_akshare_daily(codes):
    """
    通过 akshare futures_zh_daily_sina 逐合约获取日线数据，取最新一条。

    注意：此接口数据更新有延迟（收盘后 1-2 小时），15:15 时通常只有前一交易日数据。

    Returns:
        list[dict]
    """
    import akshare as ak

    results = []
    for code in codes:
        try:
            df = ak.futures_zh_daily_sina(symbol=code)

            if df is None or (hasattr(df, "empty") and df.empty):
                continue

            row = df.iloc[-1]
            record = {
                "code": code,
                "instrument": code[:2],
                "month": code[2:],
                "name": "",
                "trade_date": str(row.get("date", "")),
                "open": _safe_float(row.get("open")),
                "high": _safe_float(row.get("high")),
                "low": _safe_float(row.get("low")),
                "close": _safe_float(row.get("close")),
                "volume": _safe_int(row.get("volume")),
                "amount": None,
                "open_interest": _safe_int(row.get("hold")),
                "settle": _safe_float(row.get("settle")),
                "source": "akshare_daily",
            }

            if record["close"] is None or record["volume"] is None or record["volume"] == 0:
                continue

            results.append(record)

        except ValueError:
            continue
        except Exception as e:
            print(f"[WARN] akshare daily 获取 {code} 失败: {e}", file=sys.stderr)

    return results


# ── ETF 收盘价 ────────────────────────────────────────────────


def fetch_etf_close():
    """
    获取 IH/IF/IC/IM 对应 ETF 的最新收盘价。

    仅用 Sina HTTP（东方财富已弃用，不做回退）。

    Returns:
        dict: {instrument: {"code": "510050", "name": "上证50ETF", "close": 3.050}, ...}
    """
    # Sina 代码映射
    sina_map = {
        "IH": "sh510050",
        "IF": "sh510300",
        "IC": "sh510500",
        "IM": "sh512100",
    }

    from urllib.request import Request, urlopen
    import time

    url = SINA_URL + ",".join(sina_map.values())
    rev = {v: k for k, v in sina_map.items()}
    etf_data = {}

    # GitHub Actions 上 Sina 实时接口常因限流返回空（导致 ETF 抓到 0 只），
    # 重试最多 3 次（间隔 2 秒），集齐 4 只才收手，避免偶发限流污染数据。
    for attempt in range(3):
        try:
            req = Request(url, headers=SINA_HEADERS)
            with urlopen(req, timeout=15) as resp:
                text = resp.read().decode("gbk", errors="replace")

            batch = {}
            for m in re.finditer(r'var hq_str_(\w+)="([^"]*)"', text):
                sina_code = m.group(1)
                vals = m.group(2).split(",")
                if sina_code not in rev or len(vals) < 4:
                    continue
                inst = rev[sina_code]
                code, name = INSTRUMENT_ETF[inst]
                batch[inst] = {
                    "code": code,
                    "name": name,
                    "close": float(vals[3]) if vals[3] and vals[3] != "" else None,
                    "date": vals[30].strip() if len(vals) > 30 else "",
                }
            if batch:
                etf_data = batch
            # 全部 4 只都拿到才视为成功
            if len(etf_data) >= len(sina_map):
                break
        except Exception as e:
            print(f"[WARN] ETF 实时获取第 {attempt + 1} 次失败: {e}", file=sys.stderr)
        if attempt < 2:
            time.sleep(2)

    return etf_data


# ── 数据获取：新浪 HTTP（备） ────────────────────────────────────


def fetch_via_sina(codes, batch_size=40):
    """
    从新浪财经 HTTP 批量获取期货行情。

    字段映射（已验证）：
      [0] 今开  [1] 最高  [2] 最低  [3] 收盘  [4] 成交量
      [5] 成交额 [6] 持仓量
      [36] 日期 [37] 时间  [48] 结算价  [49] 名称
    """
    # Sina 需要 CFF_RE_ 前缀
    sina_codes = [f"CFF_RE_{c}" for c in codes]
    results = []

    for i in range(0, len(sina_codes), batch_size):
        batch = sina_codes[i : i + batch_size]
        url = SINA_URL + ",".join(batch)

        try:
            from urllib.request import Request, urlopen
            from urllib.error import URLError

            req = Request(url, headers=SINA_HEADERS)
            with urlopen(req, timeout=15) as resp:
                text = resp.read().decode("gbk", errors="replace")
        except URLError as e:
            print(f"[WARN] 网络请求失败: {e}", file=sys.stderr)
            continue
        except Exception as e:
            print(f"[WARN] 请求异常: {e}", file=sys.stderr)
            continue

        for m in re.finditer(r'var hq_str_(\w+)="([^"]*)"', text):
            raw_code = m.group(1)
            vals = m.group(2).split(",")

            if len(vals) < 10:
                continue

            code = raw_code.replace("CFF_RE_", "")

            try:
                close = _safe_float(vals[3])
                volume = _safe_int(vals[4])

                if close is None or volume is None or volume == 0:
                    continue

                record = {
                    "code": code,
                    "instrument": code[:2],
                    "month": code[2:],
                    "name": vals[49].strip() if len(vals) > 49 else "",
                    "trade_date": vals[36].strip() if len(vals) > 36 else "",
                    "open": _safe_float(vals[0]),
                    "high": _safe_float(vals[1]),
                    "low": _safe_float(vals[2]),
                    "close": close,
                    "volume": volume,
                    "amount": _safe_float(vals[5]),
                    "open_interest": _safe_int(vals[6]),
                    "settle": _safe_float(vals[48]) if len(vals) > 48 else None,
                    "source": "sina_http",
                }
                results.append(record)
            except (IndexError, ValueError):
                continue

    return results


# ── 统一入口 ──────────────────────────────────────────────────


def fetch_all():
    """获取所有活跃合约数据。优先级：akshare realtime → Sina HTTP → akshare daily。"""
    # 第一优先：akshare realtime（实时行情，收盘后立即可取，一次拉整个品种）
    results = []
    if _has_akshare():
        results = fetch_via_akshare_realtime()

    # 检查是否有品种缺失（可能是 close=0 被过滤或数据源延迟）
    present_insts = {r["instrument"] for r in results}
    missing_insts = [inst for inst in INSTRUMENTS if inst not in present_insts]

    # 第二优先：Sina HTTP 补充（无数据或有品种缺失时触发）
    if missing_insts or not results:
        today = date.today()
        codes = gen_contract_codes(today.year, today.month)
        pm = today.month - 1
        py = today.year
        if pm <= 0:
            pm = 12
            py -= 1
        codes.extend(gen_contract_codes(py, pm, 1))
        codes = list(dict.fromkeys(codes))

        if not results:
            print("[INFO] akshare realtime 无数据，回退到 Sina HTTP", file=sys.stderr)
        else:
            print(f"[INFO] akshare realtime 缺少品种 {missing_insts}，用 Sina HTTP 补充", file=sys.stderr)

        sina_results = fetch_via_sina(codes)
        # 合并：只添加 results 中不存在的合约
        existing_codes = {r["code"] for r in results}
        for r in sina_results:
            if r["code"] not in existing_codes:
                results.append(r)

    # 第三优先：akshare daily（延迟数据，收盘后 1-2 小时才有）
    if not results and _has_akshare():
        today = date.today()
        codes = gen_contract_codes(today.year, today.month)
        pm = today.month - 1
        py = today.year
        if pm <= 0:
            pm = 12
            py -= 1
        codes.extend(gen_contract_codes(py, pm, 1))
        codes = list(dict.fromkeys(codes))

        print("[INFO] akshare realtime 和 Sina HTTP 均无数据，回退到 akshare daily", file=sys.stderr)
        results = fetch_via_akshare_daily(codes)

    # 去重
    seen = set()
    unique = []
    for r in results:
        if r["code"] not in seen:
            seen.add(r["code"])
            unique.append(r)

    # 过滤：只保留最新交易日数据
    all_dates = [r.get("trade_date", "") for r in unique if r.get("trade_date")]
    if all_dates:
        latest_date = max(all_dates)
        unique = [r for r in unique if r.get("trade_date") == latest_date]

    # 排序：按品种 → 月份
    instr_order = {"IH": 0, "IF": 1, "IC": 2, "IM": 3}
    unique.sort(key=lambda x: (instr_order.get(x["instrument"], 99), x["month"]))

    return unique


# ── 辅助函数 ──────────────────────────────────────────────────


def _safe_float(val):
    if val is None or val == "":
        return None
    try:
        return round(float(val), 1)
    except (TypeError, ValueError):
        return None


def _safe_int(val):
    if val is None or val == "":
        return None
    try:
        return int(float(val))
    except (TypeError, ValueError):
        return None


# ── 输出格式化 ────────────────────────────────────────────────


def format_table(results, etf_data=None):
    """生成美化表格。"""
    if not results:
        return "\n  (无数据 — 可能非交易日或网络异常)\n"

    if etf_data is None:
        etf_data = {}

    dates = [r.get("trade_date", "") for r in results if r.get("trade_date")]
    trade_date = max(dates) if dates else "—"

    # 数据源标记
    src = results[0].get("source", "") if results else ""
    src_label = "akshare 实时" if src == "akshare_realtime" else ("akshare 日线" if src == "akshare_daily" else "新浪 HTTP")

    grouped = {}
    for r in results:
        grouped.setdefault(r["instrument"], []).append(r)

    lines = [f"\n  CFFEX 股指期货收盘价  |  交易日: {trade_date}  |  数据源: {src_label}\n"]
    lines.append("  " + "─" * 80)

    instr_order = ["IH", "IF", "IC", "IM"]
    for inst in instr_order:
        if inst not in grouped:
            continue
        name = INSTRUMENT_NAMES.get(inst, inst)

        # ETF 信息
        etf_line = ""
        if inst in etf_data:
            e = etf_data[inst]
            etf_c = f"{e['close']:.3f}" if e["close"] is not None else "—"
            etf_line = f"  |  {e['code']} {e['name']}: {etf_c}"

        lines.append(f"\n  {inst} — {name}{etf_line}")
        lines.append(
            f"  {'合约':>6s}  {'今开':>8s}  {'最高':>8s}  {'最低':>8s}  "
            f"{'收盘':>8s}  {'结算':>8s}  {'价差':>8s}  {'期货/ETF':>8s}  {'成交量':>10s}  {'持仓量':>10s}"
        )
        lines.append(
            f"  {'─'*6}  {'─'*8}  {'─'*8}  {'─'*8}  "
            f"{'─'*8}  {'─'*8}  {'─'*8}  {'─'*8}  {'─'*10}  {'─'*10}"
        )

        # ETF 收盘价（用于计算比率）
        etf_close = etf_data.get(inst, {}).get("close") if etf_data else None

        # 取当月合约收盘价作为基准
        front_close = grouped[inst][0]["close"] if grouped[inst] else None

        for r in grouped[inst]:
            month = r["month"]
            o = f"{r['open']:>8.1f}" if r["open"] is not None else "       —"
            h = f"{r['high']:>8.1f}" if r["high"] is not None else "       —"
            l = f"{r['low']:>8.1f}" if r["low"] is not None else "       —"
            c = f"{r['close']:>8.1f}" if r["close"] is not None else "       —"
            s = f"{r['settle']:>8.1f}" if r["settle"] is not None else "       —"
            v = f"{r['volume']:>10,d}" if r["volume"] is not None else "        —"
            oi = f"{r['open_interest']:>10,d}" if r["open_interest"] is not None else "        —"

            # 价差：当月=—，其他=收盘-当月收盘
            if r is grouped[inst][0]:
                spread = "       —"
            elif front_close is not None and r["close"] is not None:
                diff = r["close"] - front_close
                sign = "+" if diff >= 0 else ""
                spread = f"{sign}{diff:>7.1f}"
            else:
                spread = "       —"

            # 期货/ETF 比率
            if etf_close and etf_close > 0 and r["close"] is not None:
                ratio = r["close"] / etf_close
                ratio_str = f"{ratio:>8.1f}"
            else:
                ratio_str = "       —"

            lines.append(f"  {month:>6s}  {o}  {h}  {l}  {c}  {s}  {spread}  {ratio_str}  {v}  {oi}")

    lines.append("\n  " + "─" * 80)
    lines.append(f"  共 {len(results)} 个合约")
    return "\n".join(lines)


def save_excel(results, etf_data, filepath):
    """
    追加当日数据到一个固定的 Excel 文件中。
    每个交易日一个 sheet，命名为 trade_date（如 2026-07-27）。
    同时维护一个"汇总" sheet，所有交易日的当月合约收盘价 + ETF 收盘价横向展开。
    """
    if not results:
        return None

    dates = [r.get("trade_date", "") for r in results if r.get("trade_date")]
    if not dates:
        return None
    trade_date = max(dates)

    # 确保目录存在
    parent = os.path.dirname(os.path.abspath(filepath))
    if parent:
        os.makedirs(parent, exist_ok=True)

    # ── 准备当日 sheet 数据（含 ETF 收盘、价差、期货/ETF 比率） ──
    inst_order = {"IH": 0, "IF": 1, "IC": 2, "IM": 3}
    sorted_results = sorted(results, key=lambda x: (inst_order.get(x["instrument"], 99), x["month"]))

    # ETF 收盘价
    etf_close_map = {}
    for inst, e in etf_data.items():
        if e.get("close") is not None:
            etf_close_map[inst] = e["close"]

    # 当月合约收盘价
    front_close_map = {}
    for r in sorted_results:
        inst = r["instrument"]
        if inst not in front_close_map:
            front_close_map[inst] = r["close"]

    rows_daily = []
    for r in sorted_results:
        inst = r["instrument"]
        etf_c = etf_close_map.get(inst)
        front_c = front_close_map.get(inst)
        close = r["close"]

        # 价差：相对于当月合约
        if front_c is not None and close is not None:
            diff = close - front_c
            spread = round(diff, 1) if diff != 0 else 0.0
        else:
            spread = None

        # 期货/ETF 比率
        ratio = round(close / etf_c, 1) if etf_c and etf_c > 0 and close is not None else None

        rows_daily.append({
            "品种": inst,
            "合约": r["month"],
            "合约名": r.get("name", ""),
            "ETF代码": INSTRUMENT_ETF.get(inst, ("", ""))[0],
            "ETF收盘": etf_c,
            "今开": r["open"],
            "最高": r["high"],
            "最低": r["low"],
            "收盘": close,
            "结算": r["settle"],
            "价差": spread,
            "期货/ETF": ratio,
            "成交量": r["volume"],
            "持仓量": r["open_interest"],
        })

    df_daily = pd.DataFrame(rows_daily)

    # ── 准备汇总 sheet 数据 ──
    summary_row = {"日期": trade_date}
    for inst in ["IH", "IF", "IC", "IM"]:
        inst_rows = [r for r in sorted_results if r["instrument"] == inst]
        if inst_rows:
            front_c = inst_rows[0]["close"]
            last_c = inst_rows[-1]["close"]
            summary_row[f"{inst}_当月收盘"] = front_c
            summary_row[f"{inst}_最远收盘"] = last_c
            if front_c is not None and last_c is not None:
                summary_row[f"{inst}_贴水"] = round(last_c - front_c, 1)
            else:
                summary_row[f"{inst}_贴水"] = None
        if inst in etf_close_map:
            etf_c = etf_close_map[inst]
            summary_row[f"{inst}_ETF"] = etf_c
            if inst_rows and inst_rows[0]["close"] is not None:
                summary_row[f"{inst}_期货/ETF"] = round(inst_rows[0]["close"] / etf_c, 1)

    df_summary_new = pd.DataFrame([summary_row])

    # ── 读写 Excel（如果已存在则追加） ──
    if os.path.exists(filepath):
        # 读取已有数据
        existing_sheets = {}
        with pd.ExcelFile(filepath) as xls:
            for sheet in xls.sheet_names:
                existing_sheets[sheet] = pd.read_excel(xls, sheet_name=sheet)

        # 检查当日 sheet 是否已存在，存在则覆盖
        if trade_date in existing_sheets:
            del existing_sheets[trade_date]

        # 合并汇总数据
        if "汇总" in existing_sheets:
            df_existing_summary = existing_sheets["汇总"]
            # 去重：移除同日旧数据
            df_existing_summary = df_existing_summary[df_existing_summary["日期"] != trade_date]
            df_summary = pd.concat([df_existing_summary, df_summary_new], ignore_index=True)
        else:
            df_summary = df_summary_new

        # 按日期排序
        df_summary = df_summary.sort_values("日期").reset_index(drop=True)

        # 写入
        with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
            # 先写汇总
            df_summary.to_excel(writer, sheet_name="汇总", index=False)
            # 再写当日
            df_daily.to_excel(writer, sheet_name=trade_date, index=False)
            # 保留其他 sheet（历史交易日）
            for sheet_name, df_old in existing_sheets.items():
                if sheet_name not in ("汇总", trade_date):
                    df_old.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        # 新文件
        with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
            df_summary_new.to_excel(writer, sheet_name="汇总", index=False)
            df_daily.to_excel(writer, sheet_name=trade_date, index=False)

    return filepath


def save_csv(results, output_dir=None):
    """保存为 CSV 文件。"""
    if not results:
        return None

    if output_dir is None:
        output_dir = os.path.join(os.path.dirname(__file__), "..", "data")
    os.makedirs(output_dir, exist_ok=True)

    dates = [r.get("trade_date", "") for r in results if r.get("trade_date")]
    trade_date = max(dates) if dates else datetime.now().strftime("%Y-%m-%d")
    filename = f"cffex_futures_{trade_date}.csv"
    filepath = os.path.join(output_dir, filename)

    fields = [
        "code", "instrument", "month", "name", "trade_date",
        "open", "high", "low", "close", "settle",
        "volume", "amount", "open_interest",
    ]

    with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(results)

    return filepath


# ── 主入口 ────────────────────────────────────────────────────


def main():
    json_mode = "--json" in sys.argv
    csv_mode = "--csv" in sys.argv
    excel_mode = "--excel" in sys.argv

    output_dir = None
    excel_file = None
    for i, arg in enumerate(sys.argv):
        if arg == "--output-dir" and i + 1 < len(sys.argv):
            output_dir = sys.argv[i + 1]
        if arg == "--excel-file" and i + 1 < len(sys.argv):
            excel_file = sys.argv[i + 1]

    results = fetch_all()
    etf_data = fetch_etf_close()

    if json_mode:
        output = {"futures": results, "etf": etf_data}
        print(json.dumps(output, ensure_ascii=False, indent=2, default=str))
    else:
        print(format_table(results, etf_data))

        if csv_mode:
            fp = save_csv(results, output_dir)
            if fp:
                print(f"\n  已保存 CSV: {fp}")

        if excel_mode and excel_file:
            fp = save_excel(results, etf_data, excel_file)
            if fp:
                print(f"\n  已保存 Excel: {fp}")

        if not results:
            print("\n  未获取到数据，可能原因：")
            print("     • 当前为非交易日")
            print("     • 网络连接异常")
            print("     • 请稍后重试或检查网络")


if __name__ == "__main__":
    main()
