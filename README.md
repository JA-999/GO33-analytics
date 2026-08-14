# GO33 运营归因分析系统 · Web 版

多用户共享的运营数据实时归因分析 Web 应用，基于 Streamlit 构建，可一键部署到 **Streamlit Cloud**。

## 功能概览

- **侧边栏**：拖拽上传 Excel（`.xlsx` / `.xls`）、选择分析基准日期（默认报表最新一日）、导出 Markdown 报告、下载对比 CSV
- **顶部 KPI Card**：`st.metric` 动态展示存款金额 / 公司输赢 / 毛利 / 有效投注 / 盈利率 / 红利金额 及 DoD 环比涨跌
- **Tab 1 · 🤖 AI 智能诊断与归因**：核心结论 + 涨跌深度归因（场馆 / VIP / 大客 / 红利）+ 三条落地运营建议
- **Tab 2 · 📈 趋势与对比看板**：Plotly 渲染日度毛利趋势、有效投注 vs 公司输赢、场馆杀率分布、MTD vs 上月日均
- **Tab 3 · 📑 基础数据明细**：规范化日表 / 场馆表 / VIP 每日分布

## 项目结构

```
.
├── app.py              # Streamlit 核心界面与可视化
├── core_analytics.py   # 数据解析 / 日周月对比 / 多维归因算法
├── requirements.txt    # 依赖声明
└── README.md           # 本文件
```

## 本地运行

```bash
pip install -r requirements.txt
streamlit run app.py
# 浏览器打开 http://localhost:8501
```

## 部署到 Streamlit Cloud

1. 将本项目推送到 GitHub 仓库（确保 `app.py` 在仓库根目录）
2. 打开 https://streamlit.io/cloud → **New app**
3. 选择仓库与分支，Main file path 填 `app.py`
4. 高级设置中 Python version 选 **3.11**
5. 点击 **Deploy** —— 系统会自动按 `requirements.txt` 安装依赖

> 无需 secrets / 外部数据库：每个用户上传的报表仅在会话内解析，结果存于 `st.session_state`，天然支持多用户隔离共享。

## 报表要求

| 工作表 | 用途 | 缺失时 |
|---|---|---|
| `ribao1` | 月度累计 + 每日明细（必需） | 报错提示 |
| `场馆` | 各游戏大类输赢/有效投注/杀率日环比 | 跳过场馆归因 |
| `VIP` | VIP 各等级存款/投注/投注额对比 | 跳过 VIP 归因 |
| `红利明细` | 各活动每日红利 | 跳过红利成本归因 |
| `大客输赢` | 大客赢前十 / 亏后十 | 跳过大客预警 |
| `留存数据-60月` | 首充留存率趋势 | 跳过留存建议 |
| `每日投注区间` | 小额/大额客群结构 | 跳过客群表 |

**容错**：自动适配日期格式（Excel datetime / `8-13` / `2026-08-13`）、NaN、缺列、列名变体（如「红利（不含返水）」），日明细表头自动定位无需固定行号。

## 波动归因逻辑

- 日度环比：最新日 vs 前一日 vs 上周同日
- **±10% 波动阀值**：毛利 / 公司输赢 / 存款金额任一环比波动超阈值即触发深度归因
- 归因维度：流量转化、游戏场馆、VIP/大客、红利成本
- 趋势比对：MTD 日均 vs 上月日均增幅、客群结构、留存趋势

## 核心 API（供二次开发）

```python
from core_analytics import analyze_workbook, render_markdown

result = analyze_workbook("报表.xlsx", base_date=None)  # 结构化 dict
md     = render_markdown(result)                         # Markdown 文本
# result 含: data / attribution / kpi / daily_df / venue_df / vip_df / trend_df
```
