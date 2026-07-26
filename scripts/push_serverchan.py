#!/usr/bin/env python3
"""
Server酱 (ServerChan) 微信推送脚本

获取 CFFEX 股指期货收盘价数据，格式化为 Markdown，通过 Server酱 API 推送到个人微信。

环境变量：
    SERVERCHAN_SENDKEY  — Server酱 SendKey（必填）
    EXCEL_FILE          — Excel 文件路径（可选，用于保存当日数据）

用法：
    python push_serverchan.py
    python push_serverchan.py --excel --excel-file /path/to/cffex_daily.xlsx

注册 SendKey：https://sct.ftqq.com/
"""

import os
import sys
import json
import urllib.request
import urllib.parse

# 添加脚本目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fetch_close_prices import (
    fetch_all,
    fetch_etf_close,
    save_excel,
    INSTRUMENT_NAMES,
    INSTRUMENT_ETF,
)

SERVERCHAN_URL = "https://sctapi.ftqq.com/{key}.send"


# ── 格式化 ──────────────────────────────────────────────────


def format_markdown(results, etf_data):
    """
    将收盘价数据格式化为 Server酱 支持的 Markdown。

    Server酱 Markdown 限制 32KB，表格需简洁。
    """
    if not results:
        return None

    dates = [r.get("trade_date", "") for r in results if r.get("trade_date")]
    trade_date = max(dates) if dates else "—"

    # 按品种分组
    grouped = {}
    for r in results:
        grouped.setdefault(r["instrument"], []).append(r)

    md = f"# CFFEX 股指期货收盘价\n\n"
    md += f"> 交易日: **{trade_date}**\n\n"

    for inst in ["IH", "IF", "IC", "IM"]:
        if inst not in grouped:
            continue

        name = INSTRUMENT_NAMES.get(inst, inst)

        # ETF 信息
        etf_close = None
        etf_info = ""
        if inst in etf_data:
            e = etf_data[inst]
            etf_close = e.get("close")
            if etf_close is not None:
                etf_info = f"  |  {e['code']} {e['name']}: **{etf_close:.3f}**"

        md += f"## {inst} {name}{etf_info}\n\n"
        md += "| 合约 | 收盘 | 价差 | 期货/ETF | 成交量 | 持仓量 |\n"
        md += "|:---:|---:|---:|---:|---:|---:|\n"

        front_close = grouped[inst][0]["close"]

        for r in grouped[inst]:
            close = r["close"]

            # 价差
            if r is grouped[inst][0]:
                spread = "—"
            elif front_close is not None and close is not None:
                spread = f"{close - front_close:+.1f}"
            else:
                spread = "—"

            # 期货/ETF 比率
            if etf_close and etf_close > 0 and close is not None:
                ratio = f"{close / etf_close:.1f}"
            else:
                ratio = "—"

            vol = f"{r['volume']:,}" if r["volume"] else "—"
            oi = f"{r['open_interest']:,}" if r["open_interest"] else "—"

            md += f"| {r['month']} | {close:.1f} | {spread} | {ratio} | {vol} | {oi} |\n"

        md += "\n"

    # 贴水概况
    md += "---\n\n**贴水概况（远月 vs 当月）**\n\n"
    md += "| 品种 | 当月收盘 | 远月收盘 | 贴水 | 贴水率 |\n"
    md += "|:---:|---:|---:|---:|---:|\n"

    for inst in ["IH", "IF", "IC", "IM"]:
        if inst not in grouped:
            continue
        rows = grouped[inst]
        front = rows[0]["close"]
        last = rows[-1]["close"]
        if front is not None and last is not None:
            diff = last - front
            pct = diff / front * 100 if front else 0
            md += f"| {inst} | {front:.1f} | {last:.1f} | {diff:+.1f} | {pct:+.1f}% |\n"

    md += "\n"

    return md


# ── 推送 ──────────────────────────────────────────────────


def push_serverchan(sendkey, title, content):
    """
    POST 到 Server酱 API。

    Returns:
        dict: API 返回的 JSON
    """
    url = SERVERCHAN_URL.format(key=sendkey)

    # Server酱支持 Markdown（desp 字段）
    data = urllib.parse.urlencode(
        {
            "title": title,
            "desp": content,
        }
    ).encode("utf-8")

    req = urllib.request.Request(url, data=data, method="POST")
    req.add_header("Content-Type", "application/x-www-form-urlencoded")

    with urllib.request.urlopen(req, timeout=30) as resp:
        return json.loads(resp.read().decode("utf-8"))


# ── 主入口 ──────────────────────────────────────────────────


def main():
    sendkey = os.environ.get("SERVERCHAN_SENDKEY", "")
    if not sendkey:
        print("[ERROR] 未设置环境变量 SERVERCHAN_SENDKEY", file=sys.stderr)
        print("        请在 Server酱 (https://sct.ftqq.com/) 注册获取 SendKey", file=sys.stderr)
        sys.exit(1)

    # 可选：保存 Excel
    excel_mode = "--excel" in sys.argv
    excel_file = None
    for i, arg in enumerate(sys.argv):
        if arg == "--excel-file" and i + 1 < len(sys.argv):
            excel_file = sys.argv[i + 1]
        # 也支持环境变量
        if not excel_file:
            excel_file = os.environ.get("EXCEL_FILE", "")

    # 获取数据
    print("[1/3] 获取期货数据...")
    results = fetch_all()
    if not results:
        print("[INFO] 未获取到数据，可能非交易日，跳过推送")
        return

    print(f"      获取到 {len(results)} 个合约")

    print("[2/3] 获取 ETF 数据...")
    etf_data = fetch_etf_close()
    etf_count = len(etf_data) if etf_data else 0
    print(f"      获取到 {etf_count} 只 ETF")

    # 可选：保存 Excel
    if excel_mode and excel_file:
        print(f"[2.5] 保存 Excel: {excel_file}")
        fp = save_excel(results, etf_data, excel_file)
        if fp:
            print(f"      Excel 已保存")

    # 格式化
    dates = [r.get("trade_date", "") for r in results if r.get("trade_date")]
    trade_date = max(dates) if dates else ""
    title = f"CFFEX收盘 {trade_date}"

    content = format_markdown(results, etf_data)
    if not content:
        print("[INFO] 无数据可推送")
        return

    # 推送
    print(f"[3/3] 推送到微信: {title}")
    result = push_serverchan(sendkey, title, content)

    if result.get("code") == 0:
        print(f"[OK] 推送成功!")
        # 如果有 pushlink，打印出来
        if result.get("data", {}).get("pushid"):
            print(f"     pushid: {result['data']['pushid']}")
    else:
        print(f"[ERROR] 推送失败: {json.dumps(result, ensure_ascii=False)}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
