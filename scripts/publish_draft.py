#!/usr/bin/env python3
"""
微信公众号草稿箱发布脚本

获取 CFFEX 股指期货收盘价数据，格式化为微信兼容 HTML，
直接通过微信公众号 API 推送到草稿箱。

环境变量：
    WECHAT_APP_ID      — 公众号 AppID（必填）
    WECHAT_APP_SECRET  — 公众号 AppSecret（必填）

前置条件：
    1. 个人订阅号在 mp.weixin.qq.com → 开发 → 基本配置 获取 AppID/AppSecret
    2. 在同一页面添加本机公网 IP 到白名单（curl ifconfig.me 查看）

用法：
    python publish_draft.py                # 仅推草稿
    python publish_draft.py --excel --excel-file /path/to/cffex_daily.xlsx  # 推草稿 + 存 Excel
"""

import os
import sys
import json
import urllib.request
import urllib.parse

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fetch_close_prices import (
    fetch_all,
    fetch_etf_close,
    save_excel,
    INSTRUMENT_NAMES,
    INSTRUMENT_ETF,
)

# ── 微信 API ────────────────────────────────────────────────

TOKEN_URL = "https://api.weixin.qq.com/cgi-bin/token"
DRAFT_URL = "https://api.weixin.qq.com/cgi-bin/draft/add"


def get_access_token(appid, secret):
    """获取 access_token，有效期 7200 秒。"""
    params = urllib.parse.urlencode({
        "grant_type": "client_credential",
        "appid": appid,
        "secret": secret,
    })
    url = f"{TOKEN_URL}?{params}"
    with urllib.request.urlopen(url, timeout=15) as resp:
        result = json.loads(resp.read().decode("utf-8"))

    if "access_token" in result:
        return result["access_token"]

    errcode = result.get("errcode", -1)
    errmsg = result.get("errmsg", "未知错误")

    if errcode == 40164:
        msg = (
            "IP 不在白名单中。请将本机公网 IP 添加到微信公众号后台：\n"
            "  mp.weixin.qq.com → 开发 → 基本配置 → IP白名单"
        )
    elif errcode == 41002:
        msg = "AppID 无效，请检查 WECHAT_APP_ID 是否正确"
    elif errcode == 40001:
        msg = "AppSecret 无效，请检查 WECHAT_APP_SECRET 是否正确"
    else:
        msg = f"获取 access_token 失败 [{errcode}]: {errmsg}"

    raise RuntimeError(msg)


def add_draft(token, title, html_content):
    """向草稿箱添加一篇草稿。返回 media_id。"""
    data = json.dumps({
        "articles": [{
            "title": title,
            "content": html_content,
        }]
    }, ensure_ascii=False).encode("utf-8")

    url = f"{DRAFT_URL}?access_token={token}"
    req = urllib.request.Request(url, data=data, method="POST")
    req.add_header("Content-Type", "application/json; charset=utf-8")

    with urllib.request.urlopen(req, timeout=30) as resp:
        result = json.loads(resp.read().decode("utf-8"))

    if result.get("media_id"):
        return result["media_id"]

    errcode = result.get("errcode", -1)
    errmsg = result.get("errmsg", "未知错误")
    raise RuntimeError(f"添加草稿失败 [{errcode}]: {errmsg}")


# ── HTML 格式化 ──────────────────────────────────────────────

# 微信支持的 HTML 样式（兼容 MP 编辑器）
CSS_RESET = """
<style>
  body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif; color: #333; font-size: 15px; line-height: 1.75; padding: 0 16px; }
  h1 { font-size: 22px; text-align: center; margin: 24px 0 8px; color: #1a1a1a; }
  h2 { font-size: 18px; margin: 28px 0 8px; padding-left: 10px; border-left: 4px solid #e74c3c; color: #1a1a1a; }
  h3 { font-size: 15px; margin: 16px 0 4px; color: #555; }
  table { width: 100%; border-collapse: collapse; margin: 12px 0; font-size: 13px; }
  th { background: #f5f5f5; font-weight: 600; padding: 6px 8px; border: 1px solid #ddd; text-align: center; }
  td { padding: 5px 8px; border: 1px solid #ddd; text-align: center; }
  tr:nth-child(even) td { background: #fafafa; }
  .red { color: #e74c3c; font-weight: 600; }
  .green { color: #27ae60; font-weight: 600; }
  .note { color: #999; font-size: 13px; text-align: center; margin-top: 24px; }
  hr { border: none; border-top: 1px solid #eee; margin: 20px 0; }
</style>
"""


def color_num(val, with_sign=False):
    """涨红跌绿，中国配色惯例。"""
    if val is None:
        return "—"
    s = f"{val:+.1f}" if with_sign else f"{val:.1f}"
    if isinstance(val, (int, float)):
        if val > 0:
            return f'<span class="red">{s}</span>'
        elif val < 0:
            return f'<span class="green">{s}</span>'
    return s


def format_html(results, etf_data):
    """格式化为微信兼容 HTML。"""
    if not results:
        return None, None

    dates = [r.get("trade_date", "") for r in results if r.get("trade_date")]
    trade_date = max(dates) if dates else "—"

    grouped = {}
    for r in results:
        grouped.setdefault(r["instrument"], []).append(r)

    html = f"{CSS_RESET}\n"
    html += f"<h1>CFFEX 股指期货收盘价</h1>\n"
    html += f"<h3 style=\"text-align:center;color:#999;\">交易日: {trade_date}</h3>\n"

    # 品种详情表
    for inst in ["IH", "IF", "IC", "IM"]:
        if inst not in grouped:
            continue

        name = INSTRUMENT_NAMES.get(inst, inst)

        etf_close = None
        etf_sub = ""
        if inst in etf_data:
            e = etf_data[inst]
            etf_close = e.get("close")
            if etf_close is not None:
                etf_sub = f'  {e["code"]} {e["name"]}: <b>{etf_close:.3f}</b>'

        html += f"<h2>{inst} {name}{etf_sub}</h2>\n"

        html += """<table>
<tr><th>合约</th><th>收盘</th><th>价差</th><th>期货/ETF</th><th>成交量</th><th>持仓量</th></tr>
"""
        front_close = grouped[inst][0]["close"]

        for r in grouped[inst]:
            close = r["close"]

            # 价差
            if r is grouped[inst][0]:
                spread = "—"
            elif front_close is not None and close is not None:
                spread = color_num(close - front_close, with_sign=True)
            else:
                spread = "—"

            # 期货/ETF 比率
            if etf_close and etf_close > 0 and close is not None:
                ratio = f"{close / etf_close:.1f}"
            else:
                ratio = "—"

            vol = f"{r['volume']:,}" if r["volume"] else "—"
            oi = f"{r['open_interest']:,}" if r["open_interest"] else "—"

            html += f"<tr><td>{r['month']}</td><td>{close:.1f}</td><td>{spread}</td><td>{ratio}</td><td>{vol}</td><td>{oi}</td></tr>\n"

        html += "</table>\n"

    # 贴水概况
    html += "<hr>\n<h2>贴水概况（远月 vs 当月）</h2>\n"
    html += """<table>
<tr><th>品种</th><th>当月收盘</th><th>远月收盘</th><th>贴水</th><th>贴水率</th></tr>
"""
    for inst in ["IH", "IF", "IC", "IM"]:
        if inst not in grouped:
            continue
        rows = grouped[inst]
        front = rows[0]["close"]
        last = rows[-1]["close"]
        if front is not None and last is not None:
            diff = last - front
            pct = diff / front * 100 if front else 0
            html += f"<tr><td><b>{inst}</b></td><td>{front:.1f}</td><td>{last:.1f}</td><td>{color_num(diff, with_sign=True)}</td><td>{color_num(pct, with_sign=True)}%</td></tr>\n"

    html += "</table>\n"
    html += '<p class="note">数据来源: akshare · 每日自动生成 · 公众号草稿箱</p>\n'

    return html, trade_date


# ── 主入口 ──────────────────────────────────────────────────


def main():
    appid = os.environ.get("WECHAT_APP_ID", "")
    secret = os.environ.get("WECHAT_APP_SECRET", "")

    if not appid or not secret:
        print("[ERROR] 请设置环境变量:", file=sys.stderr)
        print("  WECHAT_APP_ID      公众号 AppID", file=sys.stderr)
        print("  WECHAT_APP_SECRET  公众号 AppSecret", file=sys.stderr)
        print("", file=sys.stderr)
        print("获取方式: mp.weixin.qq.com → 开发 → 基本配置", file=sys.stderr)
        sys.exit(1)

    # 可选：保存 Excel
    excel_mode = "--excel" in sys.argv
    excel_file = None
    for i, arg in enumerate(sys.argv):
        if arg == "--excel-file" and i + 1 < len(sys.argv):
            excel_file = sys.argv[i + 1]

    # 获取数据
    print("[1/4] 获取期货数据...")
    results = fetch_all()
    if not results:
        print("[INFO] 未获取到数据，可能非交易日")
        return

    print(f"      获取到 {len(results)} 个合约")

    print("[2/4] 获取 ETF 数据...")
    etf_data = fetch_etf_close()
    etf_count = len(etf_data) if etf_data else 0
    print(f"      获取到 {etf_count} 只 ETF")

    # 可选 Excel
    if excel_mode and excel_file:
        print(f"[2.5] 保存 Excel: {excel_file}")
        save_excel(results, etf_data, excel_file)

    # 格式化
    print("[3/4] 格式化为微信 HTML...")
    html, trade_date = format_html(results, etf_data)
    title = f"CFFEX 股指期货收盘 {trade_date}"

    # 获取 token 并推送
    print("[4/4] 推送到公众号草稿箱...")
    try:
        token = get_access_token(appid, secret)
        media_id = add_draft(token, title, html)
        print(f"\n[OK] 草稿已创建!")
        print(f"     media_id: {media_id}")
        print(f"     标题: {title}")
        print(f"\n前往 mp.weixin.qq.com → 草稿箱 → 预览并发布")
    except RuntimeError as e:
        print(f"\n[ERROR] {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
