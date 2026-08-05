#!/usr/bin/env python3
"""
CFFEX 股指期货每日采集 + 邮件发送脚本

采集期货收盘价 + ETF 数据 → 保存 Excel → 格式化邮件 → SMTP 发送

环境变量：
    QQ_EMAIL_ACCOUNT   — QQ 邮箱账号
    QQ_EMAIL_AUTH_CODE — QQ 邮箱授权码
    EMAIL_TO           — 收件人（可选，默认同账号）

用法：
    python daily_collect_email.py --excel-file /path/to/cffex_daily.xlsx
"""

import os
import sys
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fetch_close_prices import (
    fetch_all,
    fetch_etf_close,
    save_excel,
    INSTRUMENTS,
    INSTRUMENT_NAMES,
    INSTRUMENT_ETF,
)


# ── 邮件样式 ──────────────────────────────────────────────────

EMAIL_CSS = """
<style>
  body { font-family: -apple-system, 'Microsoft YaHei', 'PingFang SC', sans-serif;
         background: #f5f6fa; margin: 0; padding: 20px; color: #2d3436; }
  .container { max-width: 680px; margin: 0 auto; }
  .card { background: #fff; border-radius: 10px; padding: 20px 24px;
          margin-bottom: 16px; box-shadow: 0 1px 4px rgba(0,0,0,.06); }
  h1 { font-size: 20px; margin: 0 0 4px; color: #1a1a2e; }
  .subtitle { font-size: 13px; color: #999; margin-bottom: 16px; }

  table { width: 100%; border-collapse: collapse; font-size: 13px; }
  th { background: #f0f2f5; color: #555; font-weight: 600; text-align: right;
       padding: 8px 10px; border-bottom: 2px solid #e0e0e0; white-space: nowrap; }
  th:first-child, th.col-left { text-align: left; }
  td { padding: 7px 10px; text-align: right; border-bottom: 1px solid #f0f0f0; }
  td:first-child, td.col-left { text-align: left; font-weight: 500; }

  .up   { color: #e74c3c; font-weight: 600; }
  .down { color: #27ae60; font-weight: 600; }
  .muted { color: #bbb; }
  .highlight { background: #fff9e6; font-weight: 700; }
  .highlight td { border-bottom-color: #f5e6b0; }

  .summary-table td { font-size: 15px; padding: 10px 10px; }
  .summary-table .inst-name { font-size: 18px; font-weight: 700; }
  .summary-table .ratio-val { font-size: 16px; font-weight: 700; }

  .tag { display: inline-block; padding: 1px 6px; border-radius: 3px;
         font-size: 11px; margin-left: 4px; }
  .tag-front { background: #e8f4fd; color: #2980b9; }

  .section-title { font-size: 15px; font-weight: 700; color: #333;
                   margin: 0 0 10px; padding-bottom: 6px;
                   border-bottom: 2px solid #eee; }
  .section-title .etf-info { font-size: 12px; color: #888; font-weight: 400; margin-left: 8px; }

  .footer { text-align: center; font-size: 11px; color: #bbb; padding: 12px; }
</style>
"""


def get_prev_close(code_map):
    """
    获取各品种当月合约的「昨收」（前一交易日收盘价），用于计算涨跌/涨跌幅。

    来源：akshare futures_zh_daily_sina（新浪日线，仅读取历史收盘，符合只用 Sina 的约定）。
    取 date < 今日 的最新一行收盘作为昨收。
    返回 {instrument: 昨收(float)}；取不到时该品种不出现于字典。
    """
    prev = {}
    try:
        import akshare as ak
        from datetime import date as _d
        today_str = _d.today().strftime("%Y-%m-%d")
        for inst, code in code_map.items():
            try:
                df = ak.futures_zh_daily_sina(symbol=code)
                if df is None or getattr(df, "empty", True):
                    continue
                df = df.copy()
                df["date"] = df["date"].astype(str)
                prior = df[df["date"] < today_str]
                if prior.empty:
                    # 兜底：取倒数第二行（假设最后一行是今日）
                    if len(df) >= 2:
                        prior = df.iloc[[-2]]
                    else:
                        continue
                row = prior.iloc[-1]
                pc = row.get("close")
                if pc is not None:
                    try:
                        prev[inst] = float(pc)
                    except (TypeError, ValueError):
                        pass
            except Exception:
                continue
    except Exception:
        pass
    return prev


def _third_friday(year, month):
    """返回某年某月的第三个周五（CFFEX 股指期货交割日）。"""
    import datetime as _dt
    first = _dt.date(year, month, 1)
    # weekday(): Monday=0 ... Sunday=6；第一个周五 = 1 + (4 - 周一偏移) % 7
    first_friday = 1 + (4 - first.weekday()) % 7
    return _dt.date(year, month, first_friday + 14)


def _delivery_date(month_code):
    """合约月份代码（如 '2608'）对应的交割日 date 对象。"""
    yy = 2000 + int(month_code[:2])
    mm = int(month_code[2:])
    return _third_friday(yy, mm)


def days_between_deliveries(front_month, far_month):
    """两个合约月份代码之间的交割日天数差（绝对值），用于年化跨期价差。"""
    try:
        d1 = _delivery_date(front_month)
        d2 = _delivery_date(far_month)
        return abs((d2 - d1).days)
    except (ValueError, IndexError):
        return None


def build_html_body(results, etf_data):
    """生成 HTML 邮件正文。"""
    if not results:
        return "<p style='color:#999;text-align:center;padding:40px'>今日无数据（可能为非交易日或网络异常）</p>"

    dates = [r.get("trade_date", "") for r in results if r.get("trade_date")]
    trade_date = max(dates) if dates else datetime.now().strftime("%Y-%m-%d")
    weekday = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
    try:
        dt = datetime.strptime(trade_date, "%Y-%m-%d")
        wd = weekday[dt.weekday()]
        date_display = f"{trade_date} {wd}"
    except ValueError:
        date_display = trade_date

    grouped = {}
    for r in results:
        grouped.setdefault(r["instrument"], []).append(r)

    # 取各品种当月合约代码，用于查询「昨收」
    code_map = {inst: grouped[inst][0]["code"] for inst in grouped if grouped.get(inst)}
    prev_close = get_prev_close(code_map)

    parts = [EMAIL_CSS]
    parts.append('<div class="container">')

    # ── 标题卡片 ──
    parts.append(f'<div class="card">'
                 f'<h1>CFFEX 股指期货收盘日报</h1>'
                 f'<div class="subtitle">{date_display}  &middot;  共 {len(results)} 个合约</div>')

    # ── 概览表 ──
    parts.append('<table class="summary-table">'
                 '<tr><th class="col-left">品种</th><th>当月合约</th><th>收盘</th>'
                 '<th>涨跌 / 涨跌幅</th><th>期货/ETF</th><th>价差结构</th>'
                 '<th>贴水率</th><th>年化贴水率</th></tr>')

    for inst in ["IH", "IF", "IC", "IM"]:
        if inst not in grouped:
            continue
        contracts = grouped[inst]
        front = contracts[0]
        name = INSTRUMENT_NAMES.get(inst, inst)
        month_label = f"{front['month'][:2]}年{front['month'][2:]}月"

        cls = f"{front['close']:.0f}" if front["close"] is not None else "--"

        # 涨跌 / 涨跌幅：今收盘 - 昨收
        chg_sign = ""
        chg_cls = ""
        prev_c = prev_close.get(inst)
        if front.get("close") is not None and prev_c:
            chg = front["close"] - prev_c
            pct = chg / prev_c * 100
            if chg > 0:
                chg_sign = f'<span class="up">+{chg:.1f} (+{pct:.2f}%)</span>'
                chg_cls = "up"
            elif chg < 0:
                chg_sign = f'<span class="down">{chg:.1f} ({pct:.2f}%)</span>'
                chg_cls = "down"
            else:
                chg_sign = '<span class="muted">0.0 (0.00%)</span>'
        elif front.get("close") is not None:
            chg_sign = '<span class="muted">--</span>'

        # 期货/ETF 比率
        etf_close = etf_data.get(inst, {}).get("close")
        if etf_close and etf_close > 0 and front["close"] is not None:
            ratio = front["close"] / etf_close
            ratio_str = f"{ratio:.1f}"
        else:
            ratio_str = "--"

        # 价差结构：最远月 - 当月
        last = contracts[-1]
        structure_str = "--"
        structure_cls = ""
        disc_str = "--"        # 贴水率（简单 %）
        ann_str = "--"         # 年化贴水率（%）
        rate_cls = ""
        if front["close"] is not None and last["close"] is not None and len(contracts) > 1 and front["close"] > 0:
            diff = last["close"] - front["close"]
            if diff < 0:
                structure_str = f"贴水 {abs(diff):.0f}"
                structure_cls = "down"
            elif diff > 0:
                structure_str = f"升水 {diff:.0f}"
                structure_cls = "up"
            else:
                structure_str = "平水"

            # 贴水率：以「当月 - 远月」为正（远月低于当月 = 贴水）
            discount = front["close"] - last["close"]
            disc_pct = discount / front["close"] * 100
            if disc_pct > 0.005:
                disc_str = f"贴水 {disc_pct:.2f}%"
                rate_cls = "down"
            elif disc_pct < -0.005:
                disc_str = f"升水 {abs(disc_pct):.2f}%"
                rate_cls = "up"
            else:
                disc_str = "平水"
                rate_cls = "muted"

            # 年化贴水率：跨期价差（远月-近月）须按「近月→远月」交割间隔折算，
            # 不能用近月自身的剩余天数（那样会把 4 个月价差错误按 ~18 天摊到全年）。
            dspan = days_between_deliveries(front["month"], last["month"])
            if dspan and dspan > 0:
                ann_pct = disc_pct * 365.0 / dspan
                if ann_pct > 0.005:
                    ann_str = f"贴水 {ann_pct:.2f}%"
                elif ann_pct < -0.005:
                    ann_str = f"升水 {abs(ann_pct):.2f}%"
                else:
                    ann_str = "平水"

        parts.append(
            f'<tr>'
            f'<td class="col-left inst-name">{inst}</td>'
            f'<td class="col-left">{name} <span class="tag tag-front">{month_label}</span></td>'
            f'<td>{cls}</td>'
            f'<td class="{chg_cls}">{chg_sign}</td>'
            f'<td class="ratio-val">{ratio_str}</td>'
            f'<td class="{structure_cls}">{structure_str}</td>'
            f'<td class="{rate_cls}">{disc_str}</td>'
            f'<td class="{rate_cls}">{ann_str}</td>'
            f'</tr>'
        )

    parts.append('</table>')
    parts.append('</div>')  # end 标题卡片

    # ── 每个品种的明细卡片 ──

    def fmt_f(v, prec=1):
        if v is None:
            return '<span class="muted">--</span>'
        return f"{v:.{prec}f}"

    def fmt_i(v):
        if v is None:
            return '<span class="muted">--</span>'
        if abs(v) >= 10000:
            return f"{v/10000:.1f}万"
        return f"{v:,}"

    for inst in ["IH", "IF", "IC", "IM"]:
        if inst not in grouped:
            continue
        contracts = grouped[inst]
        name = INSTRUMENT_NAMES.get(inst, inst)
        etf_close = etf_data.get(inst, {}).get("close")
        front_close = contracts[0]["close"] if contracts else None

        etf_info = ""
        if inst in etf_data:
            e = etf_data[inst]
            ec = f"{e['close']:.3f}" if e.get("close") is not None else "--"
            etf_info = f'<span class="etf-info">{INSTRUMENT_ETF[inst][0]} {e["name"]} = {ec}</span>'

        parts.append(f'<div class="card">'
                     f'<div class="section-title">{inst} — {name}{etf_info}</div>'
                     f'<table>'
                     f'<tr><th class="col-left">合约</th><th>今开</th><th>最高</th><th>最低</th>'
                     f'<th>收盘</th><th>结算</th><th>价差</th><th>期货/ETF</th>'
                     f'<th>成交量</th><th>持仓量</th></tr>')

        for i, r in enumerate(contracts):
            month = r["month"]

            o = fmt_f(r["open"])
            h = fmt_f(r["high"])
            l = fmt_f(r["low"])
            c = fmt_f(r["close"])
            s = fmt_f(r["settle"])
            v = fmt_i(r["volume"])
            oi = fmt_i(r["open_interest"])

            # 价差（相对于当月）
            if i == 0 or front_close is None or r["close"] is None:
                spread = '<span class="muted">--</span>'
            else:
                diff = r["close"] - front_close
                sign = "+" if diff >= 0 else ""
                cls = "up" if diff > 0 else ("down" if diff < 0 else "")
                spread = f'<span class="{cls}">{sign}{diff:.1f}</span>'

            # 期货/ETF 比率
            if etf_close and etf_close > 0 and r["close"] is not None:
                ratio = r["close"] / etf_close
                ratio_str = f"{ratio:.1f}"
            else:
                ratio_str = '<span class="muted">--</span>'

            row_class = "highlight" if i == 0 else ""

            parts.append(
                f'<tr class="{row_class}">'
                f'<td class="col-left">{month}</td>'
                f'<td>{o}</td><td>{h}</td><td>{l}</td>'
                f'<td>{c}</td><td>{s}</td>'
                f'<td>{spread}</td>'
                f'<td><b>{ratio_str}</b></td>'
                f'<td>{v}</td><td>{oi}</td>'
                f'</tr>'
            )

        parts.append('</table></div>')

    # ── Footer ──
    parts.append('<div class="footer">数据来源：中金所 / 新浪财经  &middot;  '
                 '自动生成于 ' + datetime.now().strftime("%H:%M") + '</div>')
    parts.append('</div>')  # end container

    return "\n".join(parts)


def build_email_body(results, etf_data):
    """生成邮件正文（HTML 格式）。"""
    return build_html_body(results, etf_data)


def send_email(subject, html_body):
    """通过 QQ 邮箱 SMTP 发送邮件（HTML + 纯文本备选）。"""
    account = os.environ.get("QQ_EMAIL_ACCOUNT", "")
    auth_code = os.environ.get("QQ_EMAIL_AUTH_CODE", "")

    if not account or not auth_code:
        print("[ERR] 缺少环境变量 QQ_EMAIL_ACCOUNT 或 QQ_EMAIL_AUTH_CODE", file=sys.stderr)
        return False

    to_addr = os.environ.get("EMAIL_TO", account)

    # 构建纯文本备选
    import re
    text_body = re.sub(r"<[^>]+>", "", html_body)
    text_body = re.sub(r"\n\s*\n+", "\n\n", text_body).strip()

    msg = MIMEMultipart("alternative")
    msg["From"] = account
    msg["To"] = to_addr
    msg["Subject"] = Header(subject, "utf-8")
    msg.attach(MIMEText(text_body, "plain", "utf-8"))
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    try:
        server = smtplib.SMTP_SSL("smtp.qq.com", 465, timeout=30)
        server.login(account, auth_code)
        server.sendmail(account, [to_addr], msg.as_string())
        server.quit()
        print(f"[OK] 邮件已发送 → {to_addr}")
        return True
    except Exception as e:
        print(f"[ERR] 邮件发送失败: {e}", file=sys.stderr)
        return False


def main():
    excel_file = None
    no_email = "--no-email" in sys.argv
    for i, arg in enumerate(sys.argv):
        if arg == "--excel-file" and i + 1 < len(sys.argv):
            excel_file = sys.argv[i + 1]

    # 1. 采集数据
    print("[1/4] 获取期货数据...")
    results = fetch_all()
    print(f"      获取到 {len(results)} 个合约")

    print("[2/4] 获取 ETF 数据...")
    etf_data = fetch_etf_close()
    print(f"      获取到 {len(etf_data)} 只 ETF")

    # 保存 Excel
    if excel_file:
        fp = save_excel(results, etf_data, excel_file)
        if fp:
            print(f"      Excel 已保存: {fp}")

    # 2. 格式化邮件
    print("\n[3/4] 格式化邮件...")
    body = build_email_body(results, etf_data)

    # 控制台摘要（不打印完整 HTML）
    for inst in ["IH", "IF", "IC", "IM"]:
        inst_rows = [r for r in results if r["instrument"] == inst]
        if not inst_rows:
            continue
        front = inst_rows[0]
        name = INSTRUMENT_NAMES.get(inst, inst)
        c = f"{front['close']:.1f}" if front["close"] is not None else "--"
        ratio_str = ""
        etf_c = etf_data.get(inst, {}).get("close")
        if etf_c and front["close"]:
            ratio_str = f"  比率={front['close']/etf_c:.1f}"
        print(f"      {inst} {name}: {c}{ratio_str}  ({len(inst_rows)}个合约)")

    # 3. 发送邮件（--no-email 时跳过，供本地自动化等不希望发信的场景使用）
    dates = [r.get("trade_date", "") for r in results if r.get("trade_date")]
    trade_date = max(dates) if dates else datetime.now().strftime("%Y-%m-%d")
    subject = f"CFFEX 股指期货收盘 {trade_date}"

    if no_email:
        print(f"\n[4/4] 跳过邮件发送（--no-email）")
        return

    print(f"\n[4/4] 发送邮件...")
    send_email(subject, body)


if __name__ == "__main__":
    main()
