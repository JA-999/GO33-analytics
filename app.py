# -*- coding: utf-8 -*-
"""
app.py — GO33 越南站 多用户共享 Web 数据归因分析系统 (Streamlit)

部署: Streamlit Cloud
依赖: streamlit, pandas, openpyxl, plotly  (见 requirements.txt)

运行:
    streamlit run app.py
"""
from __future__ import annotations

import io
from datetime import datetime

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from core_analytics import (
    analyze_workbook,
    render_markdown,
    _fmt,
    _arrow,
    _chg,
    THRESHOLD,
)

# ------------------------- 页面配置 -------------------------
st.set_page_config(
    page_title="GO33 运营归因分析引擎",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ------------------------- 样式 -------------------------
st.markdown(
    """
    <style>
    .block-container { padding-top: 1.2rem; }
    .kpi-card {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        border: 1px solid #334155; border-radius: 12px; padding: 1rem 1.1rem;
        color: #e2e8f0;
    }
    .kpi-label { font-size: 0.8rem; color: #94a3b8; margin-bottom: 0.3rem; }
    .kpi-value { font-size: 1.6rem; font-weight: 700; color: #f8fafc; }
    .kpi-delta { font-size: 0.85rem; margin-top: 0.25rem; }
    .up { color: #4ade80; } .down { color: #f87171; } .flat { color: #94a3b8; }
    .conclusion-box {
        background: #0f172a; border-left: 4px solid #3b82f6; border-radius: 8px;
        padding: 0.8rem 1rem; margin: 0.6rem 0; color: #e2e8f0;
    }
    .tag-up { background:#14532d; color:#bbf7d0; padding:1px 8px; border-radius:6px; font-size:0.8rem;}
    .tag-down { background:#7f1d1d; color:#fecaca; padding:1px 8px; border-radius:6px; font-size:0.8rem;}
    </style>
    """,
    unsafe_allow_html=True,
)

STATE_KEY = "analysis_result"


# ========================= 侧边栏 =========================
def sidebar():
    st.sidebar.title("📊 GO33 归因分析")
    st.sidebar.caption("多用户共享 · 实时运营诊断")

    uploaded = st.sidebar.file_uploader(
        "上传运营报表 (.xlsx / .xls)",
        type=["xlsx", "xls"],
        help="支持拖拽上传。需包含 ribao1 / 场馆 / VIP 等核心工作表。",
    )

    base_date = None
    if uploaded is not None:
        try:
            import tempfile, os
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
                tf.write(uploaded.getvalue())
                tmp_path = tf.name
            res0 = analyze_workbook(tmp_path)
            os.unlink(tmp_path)
            dates = list(res0["daily_df"]["日期"])
            if dates:
                sel = st.sidebar.selectbox(
                    "分析基准日期",
                    options=dates,
                    index=len(dates) - 1,
                    help="默认使用报表中最新一日。",
                )
                base_date = datetime.strptime(sel, "%Y-%m-%d")
        except Exception as e:
            st.sidebar.error(f"解析失败: {e}")

    return uploaded, base_date


def run_analysis(uploaded, base_date):
    """执行解析并返回结果 dict (写入临时文件后调用 core_analytics)。"""
    import tempfile, os
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
        tf.write(uploaded.getvalue())
        tmp_path = tf.name
    try:
        return analyze_workbook(tmp_path, base_date=base_date)
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass


# ========================= 顶部 KPI Card =========================
def kpi_cards(kpi: list):
    cols = st.columns(len(kpi))
    for col, k in zip(cols, kpi):
        with col:
            chg = k["chg"]
            val = _fmt(k["value"], k["kind"])
            if chg is None:
                delta = None
                help_txt = "无前一日数据可对比"
            else:
                # st.metric 的 delta 用百分比点表示涨跌
                delta = round(chg * 100, 1)
                help_txt = f"环比前一日 (DoD): {_arrow(chg)}"
            st.metric(
                label=k["label"],
                value=val,
                delta=delta,
                delta_color="normal",  # 涨跌均用系统配色(绿涨红跌)
                help=help_txt,
            )


# ========================= Tab 1: AI 诊断 =========================
def tab_diagnosis(result: dict):
    d = result["data"]
    att = result["attribution"]
    st.subheader("🤖 AI 智能诊断与归因")

    if not d.latest:
        st.warning("无有效日度数据。")
        return

    # 核心结论
    st.markdown("### 🎯 核心结论")
    if att["triggered"]:
        trig = " ｜ ".join(att["triggered"])
        st.markdown(
            f'<div class="conclusion-box">⚡ <b>波动阀值触发 (±10%)</b>：{trig}。'
            f'系统已自动激活深度归因。</div>',
            unsafe_allow_html=True,
        )
    else:
        st.markdown('<div class="conclusion-box">✅ 本期各核心指标波动均在 ±10% 正常区间内，运营状态稳健。</div>',
                    unsafe_allow_html=True)

    # 涨跌深度归因
    st.markdown("### 🔍 涨跌深度归因")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<span class="tag-up">✅ 核心驱动力（上涨）</span>', unsafe_allow_html=True)
        for x in att["drivers"][:3] or ["本期无显著正向驱动因素。"]:
            st.markdown(f"- {x}")
    with c2:
        st.markdown('<span class="tag-down">⚠️ 主要拖累项（下降）</span>', unsafe_allow_html=True)
        for x in att["drags"][:3] or ["本期无显著负向拖累。"]:
            st.markdown(f"- {x}")

    # 场馆 / VIP / 红利 拆解
    st.markdown("---")
    st.markdown("#### 🎰 游戏场馆归因")
    vdf = result["venue_df"]
    if not vdf.empty:
        show = vdf[["场馆", "公司输赢(今)", "输赢环比", "有效投注环比", "杀率(今)", "杀率环比"]].copy()
        show["输赢环比"] = show["输赢环比"].apply(lambda x: _arrow(x))
        show["有效投注环比"] = show["有效投注环比"].apply(lambda x: _arrow(x))
        show["杀率环比"] = show["杀率环比"].apply(lambda x: _arrow(x))
        show["公司输赢(今)"] = show["公司输赢(今)"].apply(lambda x: _fmt(x, "money"))
        show["杀率(今)"] = show["杀率(今)"].apply(lambda x: _fmt(x, "pct"))
        st.dataframe(show, width='stretch', hide_index=True)

    st.markdown("#### 👑 VIP / 大客归因")
    vip = att["details"]["vip"]
    if vip["high_vip_bet_chg"] is not None:
        st.markdown(f"- 高净值客群(>=VIP6)投注额月度环比: **{_arrow(vip['high_vip_bet_chg'])}**")
    if d.vip_segment:
        seg = " ｜ ".join(f"{s['seg']}({s['trend']},{_fmt(s['chg'], 'pct')})" for s in d.vip_segment)
        st.markdown(f"- 有效投注分段涨跌(本月 vs 上月): {seg}")
    if d.big_loss:
        w = min(d.big_loss, key=lambda x: x["win"])
        st.markdown(f"- 大客亏损 TOP1: `{w['uid']}`({w['vip']}) 当日公司输赢 **{_fmt(w['win'], 'money')}**, "
                    f"杀率 {_fmt(w['rate'], 'pct')} —— <span class='tag-down'>爆赔/离场风险</span>",
                    unsafe_allow_html=True)
    if d.big_win:
        b = max(d.big_win, key=lambda x: x["win"])
        st.markdown(f"- 大客盈利 TOP1: `{b['uid']}`({b['vip']}) 当日贡献 {_fmt(b['win'], 'money')}")

    st.markdown("#### 🎁 红利成本归因")
    bo = att["details"]["bonus"]
    t = att["details"]["traffic"]
    if bo["bonus_chg"] is not None:
        st.markdown(f"- 红利金额环比: **{_arrow(bo['bonus_chg'])}**")
    if bo["bonus_rate_latest"] is not None:
        st.markdown(f"- 红利/有效投注率: {_fmt(bo['bonus_rate_latest'], 'pct')} vs 前一日 {_fmt(bo['bonus_rate_prev'], 'pct')}")
    st.markdown(f"- 流量/转化: 首存环比 {_arrow(t['first_dep_chg'])} ｜ 转化率环比 {_arrow(t['conv_chg'])}")

    # 三条落地运营建议
    st.markdown("---")
    st.markdown("### 💡 三条落地运营建议")
    worst = att["details"]["venue"]["worst"]
    with st.expander("1️⃣ 风控 / 止损建议", expanded=True):
        if worst and getattr(worst, "rate_chg") is not None and getattr(worst, "rate_chg") < -0.05:
            st.write(f"【{worst.name}】杀率骤降至 {_fmt(worst.rate_today, 'pct')} ({_arrow(worst.rate_chg)}), "
                     f"公司输赢 {_fmt(worst.win_today, 'money')}。建议临时下调单笔下注限额 / 提高赔率审核，"
                     f"并对当日大赢家启动赢钱监控。")
        else:
            st.write("各场馆杀率处于合理区间, 保持异常波动(>±5%)实时告警即可。")
        if d.big_loss:
            st.write(f"大客亏损榜 `{d.big_loss[0]['uid']}` 等高额赔付, 建议核实投注行为是否异常, 必要时人工复核。")
    with st.expander("2️⃣ 营销 / 活动建议", expanded=True):
        fd = t["first_dep_chg"]
        if fd is not None and fd < -0.05:
            st.write(f"首存人数环比 {_arrow(fd)} 缩水, 建议针对首存渠道加投笔笔存送 / 首充礼包, "
                     f"将红利杠杆向新客首充倾斜。")
        else:
            st.write(f"首存转化平稳, 维持现有红利结构; 关注红利/有效投注率 {_fmt(bo['bonus_rate_latest'], 'pct')}, "
                     f"若持续 >2.5% 应收缩无效活动。")
    with st.expander("3️⃣ VIP / 留存维护建议", expanded=True):
        drop = None
        for s in d.vip_segment:
            if s["trend"] == "下降" and (drop is None or (s["chg"] or 0) < (drop["chg"] or 0)):
                drop = s
        if drop:
            st.write(f"VIP 分段 **{drop['seg']}** 本月人均下跌 {_arrow(drop['chg'])}, "
                     f"建议定向推送 VIP 俸禄加码 / 二次存送召回, 配置专属客服 1v1 唤醒。")
        if d.retention:
            recent = sorted(d.retention, key=lambda x: x["date"])[-7:]
            avg_d2 = __import__("numpy").nanmean([r["d2"] for r in recent if not pd.isna(r["d2"])])
            st.write(f"近 7 日首充 2 日留存率均值约 **{_fmt(avg_d2, 'pct')}**, "
                     f"建议对首充后未续存用户投放关怀援助金类召回券, 提升 2→7 日留存。")


# ========================= Tab 2: 趋势看板 =========================
def tab_trend(result: dict):
    st.subheader("📈 趋势与对比看板")
    trend = result["trend_df"].copy()
    trend["日期"] = pd.to_datetime(trend["日期"])

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**日度毛利趋势**")
        fig = px.line(trend, x="日期", y="毛利", markers=True,
                      title="每日毛利 (VND)", template="plotly_dark")
        fig.update_traces(line=dict(color="#3b82f6", width=3))
        fig.add_hline(y=0, line_dash="dash", line_color="#64748b")
        st.plotly_chart(fig, width='stretch')

        st.markdown("**日度公司输赢 vs 有效投注**")
        fig2 = go.Figure()
        fig2.add_trace(go.Bar(x=trend["日期"], y=trend["有效投注"], name="有效投注", marker_color="#1e40af"))
        fig2.add_trace(go.Scatter(x=trend["日期"], y=trend["公司输赢"], name="公司输赢",
                                  mode="lines+markers", line=dict(color="#f59e0b", width=3)))
        fig2.update_layout(template="plotly_dark", title="有效投注(柱) vs 公司输赢(线)")
        st.plotly_chart(fig2, width='stretch')

    with c2:
        st.markdown("**场馆杀率分布 (最新日)**")
        vdf = result["venue_df"].copy()
        if not vdf.empty:
            vdf = vdf[vdf["场馆"] != "平台"] if "平台" in vdf["场馆"].values else vdf
            fig3 = px.bar(vdf, x="场馆", y="杀率(今)",
                          color="杀率(今)", color_continuous_scale="RdYlGn",
                          title="各场馆杀率", template="plotly_dark",
                          text=vdf["杀率(今)"].apply(lambda x: f"{x * 100:.2f}%"))
            fig3.update_layout(coloraxis_showscale=False)
            st.plotly_chart(fig3, width='stretch')

        st.markdown("**场馆公司输赢贡献 (最新日)**")
        if not vdf.empty:
            vdf2 = vdf.copy()
            vdf2["输赢(今)"] = vdf2["公司输赢(今)"].apply(lambda x: x / 1e6)
            fig4 = px.bar(vdf2, x="场馆", y="输赢(今)",
                          color="输赢(今)", color_continuous_scale="RdYlGn",
                          title="公司输赢 (百万 VND)", template="plotly_dark")
            st.plotly_chart(fig4, width='stretch')

    # MTD vs 上月
    st.markdown("---")
    st.markdown("**📆 MTD vs 上月日均对比**")
    cum = {m["month"]: m for m in result["data"].month_cum}
    keys = list(cum.keys())
    if len(keys) >= 2:
        import calendar, re
        prev_m, cur_m = cum[keys[-2]], cum[keys[-1]]

        def md_days(mrec):
            mm = re.search(r"(\d{1,2})月", str(mrec["month"]))
            return calendar.monthrange(2026, int(mm.group(1)))[1] if mm else 30

        pdays = md_days(prev_m)
        cdays = int(cur_m.get("days") or md_days(cur_m))
        rows = []
        for label, attr in [("存款金额", "dep_amt"), ("有效投注", "valid_bet"),
                            ("公司输赢", "win"), ("毛利", "gross_profit"), ("红利", "bonus")]:
            pv, cv = prev_m.get(attr), cur_m.get(attr)
            p_avg = pv / pdays if not pd.isna(pv) else float("nan")
            c_avg = cv / cdays if not pd.isna(cv) else float("nan")
            rows.append({
                "指标": label,
                "上月日均": p_avg, "本月日均": c_avg,
                "增幅": _chg(c_avg, p_avg),
            })
        mdf = pd.DataFrame(rows)
        mdf["上月日均"] = mdf["上月日均"].apply(lambda x: _fmt(x, "money"))
        mdf["本月日均"] = mdf["本月日均"].apply(lambda x: _fmt(x, "money"))
        mdf["增幅"] = mdf["增幅"].apply(lambda x: _arrow(x))
        st.dataframe(mdf, width='stretch', hide_index=True)
    else:
        st.info("月度累计数据不足，无法计算 MTD 对比。")


# ========================= Tab 3: 基础明细 =========================
def tab_detail(result: dict):
    st.subheader("📑 基础数据明细")
    tab_a, tab_b, tab_c = st.tabs(["📅 日度明细", "🎰 场馆明细", "👑 VIP 每日分布"])

    with tab_a:
        st.dataframe(result["daily_df"], width='stretch', hide_index=True)
    with tab_b:
        vd = result["venue_df"].copy()
        for col in ["公司输赢(本月)", "公司输赢(今)", "公司输赢(昨)", "有效投注(今)"]:
            vd[col] = vd[col].apply(lambda x: _fmt(x, "money"))
        for col in ["杀率(今)", "杀率(昨)"]:
            vd[col] = vd[col].apply(lambda x: _fmt(x, "pct"))
        for col in ["输赢环比", "有效投注环比", "杀率环比"]:
            vd[col] = vd[col].apply(lambda x: _arrow(x))
        st.dataframe(vd, width='stretch', hide_index=True)
    with tab_c:
        st.dataframe(result["vip_df"], width='stretch', hide_index=True)


# ========================= 主流程 =========================
def main():
    uploaded, base_date = sidebar()

    if uploaded is None:
        st.title("📈 GO33 运营归因分析引擎")
        st.info("👈 请在左侧侧边栏上传运营报表 (.xlsx)，系统将自动生成实时归因诊断。")
        st.markdown(
            """
            **功能概览**
            - 自动解析 ribao1 / 场馆 / VIP / 红利 / 大客 / 留存 等多工作表
            - 日度环比 (DoD) + ±10% 波动阀值自动触发深度归因
            - 场馆 / VIP / 大客 / 红利 / 流量 五维归因拆解
            - 三条可落地运营建议 + 周月趋势比对
            - 支持导出 Markdown 报告与对比 CSV
            """
        )
        return

    # 解析 (带缓存键避免重复跑)
    cache_key = (uploaded.name, uploaded.size, base_date.strftime("%Y-%m-%d") if base_date else "auto")
    if st.session_state.get("cache_key") != cache_key:
        with st.spinner("正在解析并归因…"):
            result = run_analysis(uploaded, base_date)
        st.session_state[STATE_KEY] = result
        st.session_state["cache_key"] = cache_key
    else:
        result = st.session_state[STATE_KEY]

    d = result["data"]
    st.title(f"📈 {d.platform} · {d.country} 实时运营归因")
    st.caption(f"基准日期 **{result['base_date'].strftime('%Y-%m-%d')}** ｜ "
               f"波动触发阈值 ±{THRESHOLD * 100:.0f}% ｜ 生成于 {datetime.now().strftime('%H:%M')}")

    # 顶部 KPI Card
    kpi_cards(result["kpi"])

    # 三个 Tab
    t1, t2, t3 = st.tabs([
        "🤖 AI 智能诊断与归因",
        "📈 趋势与对比看板",
        "📑 基础数据明细",
    ])
    with t1:
        tab_diagnosis(result)
    with t2:
        tab_trend(result)
    with t3:
        tab_detail(result)

    # 导出区
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 📤 导出")
    md_text = render_markdown(result)
    st.sidebar.download_button(
        "📝 导出 Markdown 报告",
        data=md_text,
        file_name=f"GO33_归因报告_{result['base_date'].strftime('%Y-%m-%d')}.md",
        mime="text/markdown",
    )
    # 对比 CSV: 日度 + 场馆 + MTD
    csv_buf = io.StringIO()
    result["daily_df"].to_csv(csv_buf, index=False)
    csv_buf.seek(0)
    st.sidebar.download_button(
        "📊 下载对比 CSV (日度明细)",
        data=csv_buf.getvalue(),
        file_name=f"GO33_日度对比_{result['base_date'].strftime('%Y-%m-%d')}.csv",
        mime="text/csv",
    )


if __name__ == "__main__":
    main()
