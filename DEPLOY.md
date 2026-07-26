# 云端部署指南

## 方案：GitHub Actions + Server酱

每天 15:15（北京时间）自动在 GitHub 云端运行，采集 CFFEX 四大股指期货收盘价，推送到个人微信。

---

## 第一步：注册 Server酱 获取 SendKey

1. 打开 https://sct.ftqq.com/
2. 用微信扫码登录
3. 进入「SendKey」页面，复制你的 SendKey（格式类似 `SCT1234abcd...`）

## 第二步：创建 GitHub 仓库

1. 在 GitHub 创建一个新仓库（public 或 private 均可）
2. 将以下文件推送到仓库：

```
your-repo/
├── scripts/
│   ├── fetch_close_prices.py    # 核心数据采集
│   └── push_serverchan.py       # 微信推送
├── .github/
│   └── workflows/
│       └── cffex-daily.yml       # GitHub Actions 定时任务
└── requirements.txt              # Python 依赖
```

文件位置：`~/.workbuddy/skills/cffex-futures/`

## 第三步：设置 GitHub Secret

1. 进入仓库 **Settings → Secrets and variables → Actions**
2. 点击 **New repository secret**
3. Name: `SERVERCHAN_SENDKEY`
4. Secret: 粘贴你的 Server酱 SendKey

## 第四步：测试运行

1. 进入仓库 **Actions** 页面
2. 选择「CFFEX 股指期货每日收盘价」workflow
3. 点击 **Run workflow** 手动触发一次
4. 检查运行日志，确认推送成功
5. 查看微信是否收到消息

## 自动运行

配置完成后，GitHub Actions 会在每个交易日（周一至周五）15:15 自动运行。
非交易日数据为空时会自动跳过，不会推送空消息。

## 本地运行（测试）

```bash
# 设置环境变量后本地测试推送
export SERVERCHAN_SENDKEY="你的SendKey"
python scripts/push_serverchan.py

# 同时保存 Excel
export SERVERCHAN_SENDKEY="你的SendKey"
python scripts/push_serverchan.py --excel --excel-file ./cffex_daily.xlsx
```

## 注意事项

- GitHub Actions cron 使用 UTC 时间，15:15 CST = 07:15 UTC
- GitHub Actions 定时任务可能有 5-15 分钟延迟，属正常现象
- Server酱免费版每天限 5 条消息，足够日常使用
- Excel 文件不会在云端累积，每次运行生成单日文件作为 artifact（保留 90 天）
- 本地 WorkBuddy 自动化继续运行，负责 Excel 累积
