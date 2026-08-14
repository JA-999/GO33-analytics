# -*- coding: utf-8 -*-
"""
core_analytics.py
=================
GO33 越南站 · 数据解析 / 日周月对比 / 多维归因算法核心模块

该模块不依赖 Streamlit, 可独立被 app.py 调用, 也可在纯 Python 环境中使用:
    from core_analytics import analyze_workbook
    result = analyze_workbook("报表.xlsx")          # 返回结构化 dict
    md     = render_markdown(result)                # 转 Markdown 文本

设计要点
--------
* 容错: 处理 NaN / 缺失 Sheet / 日期格式混用 (Excel datetime, "8-13", "2026-08-13")
* 自动定位: 日明细表头自动匹配 (含"日期"+"有效投注"的行), 不硬编码行号
* 产出: ReportData 数据容器 + 结构化归因结果, 供 UI 渲染 / CSV 导出 / Markdown 导出
"""
from __future__ import annotations

import calendar
import re
import warnings
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from typing import Any, Optional

import numpy as np
import pandas as pd

warnings.simplefilter("ignore")

THRESHOLD = 0.10  # ±10% 波动触发阈值


# =============================================================================
# 工具函数
# =============================================================================

def _to_num(v: Any) -> float:
    if v is None:
        return np.nan
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        s = v.replace(",", "").replace("%", "").strip()
        if s in ("", "-", "—", "·", "None", "nan", "NaN"):
            return np.nan
        try:
            return float(s)
        except ValueError:
            return np.nan
    return np.nan


def _to_date(v: Any) -> Optional[datetime]:
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return None
    if isinstance(v, datetime):
        return v
    if isinstance(v, str):
        s = v.strip()
        if not s:
            return None
        m = re.match(r"^(\d{1,2})[-/](\d{1,2})(?:[-/](\d{2,4}))?$", s)
        if m:
            mo, da, yr = m.groups()
            year = int(yr) if yr else datetime.now().year
            if year < 100:
                year += 2000
            try:
                return datetime(year, int(mo), int(da))
            except ValueError:
                return None
        for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S"):
            try:
                return datetime.strptime(s, fmt)
            except ValueError:
                continue
        return None
    return None


def _fmt(v: float, kind: str = "num", nd: int = 2) -> str:
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "—"
    if kind == "pct":
        return f"{v * 100:.2f}%"
    if kind == "money":
        if abs(v) >= 1e9:
            return f"{v / 1e9:.2f}B"
        if abs(v) >= 1e6:
            return f"{v / 1e6:.2f}M"
        if abs(v) >= 1e3:
            return f"{v / 1e3:.1f}K"
        return f"{v:.0f}"
    return f"{v:,.{nd}f}"


def _chg(cur: float, prev: float) -> Optional[float]:
    if prev is None or (isinstance(prev, float) and np.isnan(prev)) or prev == 0:
        return None
    if cur is None or (isinstance(cur, float) and np.isnan(cur)):
        return None
    return (cur - prev) / prev


def _arrow(chg: Optional[float]) -> str:
    if chg is None:
        return "—"
    if abs(chg) < 0.005:
        return "→ 持平"
    return ("▲ +" if chg > 0 else "▼ ") + f"{chg * 100:.1f}%"


def _getattr(obj, attr):
    return getattr(obj, attr, np.nan)


# =============================================================================
# 数据容器
# =============================================================================

@dataclass
class DailyRow:
    date: datetime
    reg: float = np.nan
    first_dep: float = np.nan
    conv: float = np.nan
    active: float = np.nan
    dep_people: float = np.nan
    dep_amt: float = np.nan
    wd_amt: float = np.nan
    net_dep: float = np.nan
    bet_people: float = np.nan
    valid_bet: float = np.nan
    win: float = np.nan
    profit_rate: float = np.nan
    bonus: float = np.nan
    rebate: float = np.nan
    gross_profit: float = np.nan


@dataclass
class VenueBlock:
    name: str
    win_mtd: float = np.nan
    win_today: float = np.nan
    win_yest: float = np.nan
    win_diff: float = np.nan
    win_chg: float = np.nan
    vb_mtd: float = np.nan
    vb_today: float = np.nan
    vb_yest: float = np.nan
    vb_diff: float = np.nan
    vb_chg: float = np.nan
    rate_mtd: float = np.nan
    rate_today: float = np.nan
    rate_yest: float = np.nan
    rate_diff: float = np.nan
    rate_chg: float = np.nan


@dataclass
class ReportData:
    platform: str = "GO33"
    country: str = "越南站"
    latest: Optional[DailyRow] = None
    prev: Optional[DailyRow] = None
    lastweek: Optional[DailyRow] = None
    mtd_avg: Optional[DailyRow] = None
    month_cum: list = field(default_factory=list)
    venues: list = field(default_factory=list)
    vip_month: dict = field(default_factory=dict)
    vip_daily: list = field(default_factory=list)
    vip_segment: list = field(default_factory=list)
    bonus_daily: dict = field(default_factory=dict)
    big_win: list = field(default_factory=list)
    big_loss: list = field(default_factory=list)
    retention: list = field(default_factory=list)
    bet_range: dict = field(default_factory=dict)
    _daily_rows: list = field(default_factory=list)


# =============================================================================
# 解析器
# =============================================================================

class ReportParser:
    def __init__(self, file_path: str):
        self.file = file_path
        self.xl = pd.ExcelFile(file_path)
        self.sheets = self.xl.sheet_names
        self.data = ReportData()

    # ---- ribao1 ----------------------------------------------------------
    def parse_ribao1(self):
        if "ribao1" not in self.sheets:
            return
        raw = self.xl.parse("ribao1", header=None, dtype=object)

        # 月度累计区: 用正则匹配 "YYYY年M月"
        for i in range(2, len(raw)):
            row = raw.iloc[i]
            mt = row[2] if isinstance(row[2], str) else None
            if not mt or not re.search(r"\d{4}年\d{1,2}月", str(mt)):
                continue
            self.data.month_cum.append({
                "month": row[2],
                "days": _to_num(row[4]),
                "reg": _to_num(row[5]), "first_dep": _to_num(row[6]),
                "conv": _to_num(row[7]), "active": _to_num(row[8]),
                "dep_people": _to_num(row[9]), "bet_people": _to_num(row[10]),
                "dep_amt": _to_num(row[11]), "wd_amt": _to_num(row[12]),
                "net_dep": _to_num(row[13]), "valid_bet": _to_num(row[15]),
                "win": _to_num(row[16]), "bonus": _to_num(row[17]),
                "rebate": _to_num(row[18]), "gross_profit": _to_num(row[19]),
                "profit_rate": _to_num(row[20]),
            })

        # 日明细: 自动定位表头行
        hdr_idx = None
        for i in range(len(raw)):
            vals = [str(x) for x in raw.iloc[i].tolist() if x is not None]
            if any("日期" in v for v in vals) and any("有效投注" in v for v in vals):
                hdr_idx = i
                break
        if hdr_idx is None:
            return
        header = raw.iloc[hdr_idx]
        col_map = {}
        alias = {
            "reg": ["注册"], "first_dep": ["首存"], "conv": ["转化率"],
            "active": ["活跃人数"], "dep_people": ["存款人数"], "dep_amt": ["存款金额"],
            "wd_amt": ["取款金额"], "net_dep": ["存提差"], "bet_people": ["投注人数"],
            "valid_bet": ["有效投注"], "win": ["公司输赢"], "profit_rate": ["盈利率"],
            "bonus": ["红利"], "rebate": ["返水"], "gross_profit": ["毛利"],
        }
        for ci, v in enumerate(header):
            if v is None:
                continue
            s = str(v).replace("（不含返水）", "").replace("（", "(").replace("）", ")")
            matched = False
            for key, aliases in alias.items():
                if any(a == s for a in aliases):
                    col_map[key] = ci
                    matched = True
                    break
            if matched:
                continue
            for key, aliases in alias.items():
                if key in col_map:
                    continue
                if key == "gross_profit" and s.endswith("率"):
                    continue
                if any(a in s for a in aliases):
                    col_map[key] = ci
                    break

        rows = []
        for i in range(hdr_idx + 1, len(raw)):
            d = _to_date(raw.iloc[i, 0])
            if d is None:
                continue
            r = raw.iloc[i]
            rows.append(DailyRow(
                date=d,
                reg=_to_num(r[col_map.get("reg", 1)]),
                first_dep=_to_num(r[col_map.get("first_dep", 2)]),
                conv=_to_num(r[col_map.get("conv", 3)]),
                active=_to_num(r[col_map.get("active", 5)]),
                dep_people=_to_num(r[col_map.get("dep_people", 6)]),
                dep_amt=_to_num(r[col_map.get("dep_amt", 7)]),
                wd_amt=_to_num(r[col_map.get("wd_amt", 9)]),
                net_dep=_to_num(r[col_map.get("net_dep", 10)]),
                bet_people=_to_num(r[col_map.get("bet_people", 11)]),
                valid_bet=_to_num(r[col_map.get("valid_bet", 13)]),
                win=_to_num(r[col_map.get("win", 14)]),
                profit_rate=_to_num(r[col_map.get("profit_rate", 15)]),
                bonus=_to_num(r[col_map.get("bonus", 16)]),
                rebate=_to_num(r[col_map.get("rebate", 17)]),
                gross_profit=_to_num(r[col_map.get("gross_profit", 18)]),
            ))
        if not rows:
            return
        rows.sort(key=lambda x: x.date)
        self.data._daily_rows = rows
        self.data.latest = rows[-1]
        self.data.prev = rows[-2] if len(rows) >= 2 else None
        if self.data.latest:
            lw = self.data.latest.date - timedelta(days=7)
            for r in rows:
                if r.date.date() == lw.date():
                    self.data.lastweek = r
                    break
        self._calc_mtd_avg(rows)

    def _calc_mtd_avg(self, rows):
        cur_month = self.data.latest.date.replace(day=1)
        mrows = [r for r in rows if r.date >= cur_month]
        if not mrows:
            return

        def avg(attr):
            vals = [getattr(r, attr) for r in mrows]
            vals = [v for v in vals if not (isinstance(v, float) and np.isnan(v))]
            return sum(vals) / len(vals) if vals else np.nan

        self.data.mtd_avg = DailyRow(
            date=cur_month,
            reg=avg("reg"), first_dep=avg("first_dep"), conv=avg("conv"),
            active=avg("active"), dep_people=avg("dep_people"),
            dep_amt=avg("dep_amt"), wd_amt=avg("wd_amt"), net_dep=avg("net_dep"),
            bet_people=avg("bet_people"), valid_bet=avg("valid_bet"), win=avg("win"),
            profit_rate=avg("profit_rate"), bonus=avg("bonus"), rebate=avg("rebate"),
            gross_profit=avg("gross_profit"),
        )

    # ---- 场馆 ------------------------------------------------------------
    def parse_venue(self):
        if "场馆" not in self.sheets:
            return
        raw = self.xl.parse("场馆", header=None, dtype=object)
        i = 4
        while i < min(len(raw), 12):
            name = raw.iloc[i, 1]
            if name in (None, "平台", "") or not isinstance(name, str):
                i += 1
                continue
            name = str(name).strip()
            if name == "平台":
                i += 1
                continue
            r = raw.iloc[i]
            vb = VenueBlock(
                name=name,
                win_mtd=_to_num(r[2]), win_today=_to_num(r[3]), win_yest=_to_num(r[4]),
                win_diff=_to_num(r[5]), win_chg=_to_num(r[6]),
                vb_mtd=_to_num(r[10]), vb_today=_to_num(r[11]), vb_yest=_to_num(r[12]),
                vb_diff=_to_num(r[13]), vb_chg=_to_num(r[14]),
                rate_mtd=_to_num(r[17]), rate_today=_to_num(r[18]), rate_yest=_to_num(r[19]),
                rate_diff=_to_num(r[20]), rate_chg=_to_num(r[21]),
            )
            self.data.venues.append(vb)
            i += 1

    # ---- VIP -------------------------------------------------------------
    def parse_vip(self):
        if "VIP" not in self.sheets:
            return
        raw = self.xl.parse("VIP", header=None, dtype=object)
        vip_levels = [">=6", "5", "4", "3", "2", "1", "0"]

        def grab_block(start_col, hr_row):
            ri = 1 + hr_row
            out = {}
            for j, lv in enumerate(vip_levels):
                c = start_col + 2 + j
                out[lv] = _to_num(raw.iloc[ri, c])
            out["合计"] = _to_num(raw.iloc[ri, start_col + 1])
            return out

        self.data.vip_month = {
            "dep_people": {"jul": grab_block(1, 0), "mtd": grab_block(1, 1), "chg": grab_block(1, 2)},
            "bet_people": {"jul": grab_block(14, 0), "mtd": grab_block(14, 1), "chg": grab_block(14, 2)},
            "bet_amt": {"jul": grab_block(25, 0), "mtd": grab_block(25, 1), "chg": grab_block(25, 2)},
        }
        for i in range(4, len(raw)):
            d = _to_date(raw.iloc[i, 0])
            if d is None:
                continue
            self.data.vip_daily.append({
                "date": d,
                "dep_people_total": _to_num(raw.iloc[i, 1]),
                "bet_people_total": _to_num(raw.iloc[i, 14]),
                "bet_amt_total": _to_num(raw.iloc[i, 25]),
            })
        seen = set()
        for i in range(len(raw) - 1, 0, -1):
            lab = raw.iloc[i, 0]
            if isinstance(lab, str) and re.match(r"^V\d|^>=V", lab) and lab not in seen:
                seen.add(lab)
                self.data.vip_segment.append({
                    "seg": lab, "trend": raw.iloc[i, 1],
                    "amt": _to_num(raw.iloc[i, 2]), "chg": _to_num(raw.iloc[i, 4]),
                })

    # ---- 红利明细 --------------------------------------------------------
    def parse_bonus(self):
        if "红利明细" not in self.sheets:
            return
        raw = self.xl.parse("红利明细", header=None, dtype=object)
        date_cols = {}
        for c in range(3, len(raw.columns)):
            d = _to_date(raw.iloc[1, c])
            if d:
                date_cols[c] = d
        for i in range(2, len(raw)):
            act = raw.iloc[i, 1]
            if not isinstance(act, str) or act.strip() == "":
                continue
            for c, d in date_cols.items():
                v = _to_num(raw.iloc[i, c])
                if not np.isnan(v):
                    self.data.bonus_daily[d] = self.data.bonus_daily.get(d, 0.0) + v

    # ---- 大客输赢 --------------------------------------------------------
    def parse_bigclient(self):
        if "大客输赢" not in self.sheets:
            return
        raw = self.xl.parse("大客输赢", header=None, dtype=object)
        mode = None
        for i in range(len(raw)):
            lab = raw.iloc[i, 1]
            if isinstance(lab, str):
                if "前十" in lab:
                    mode = "win"
                    continue
                if "后十" in lab:
                    mode = "loss"
                    continue
            val = _to_num(raw.iloc[i, 6])
            if val is None or np.isnan(val) or val == 0:
                continue
            rec = {
                "rank": _to_num(raw.iloc[i, 0]), "uid": raw.iloc[i, 1],
                "vip": raw.iloc[i, 2], "dep": _to_num(raw.iloc[i, 4]),
                "bet": _to_num(raw.iloc[i, 5]), "win": val,
                "rate": _to_num(raw.iloc[i, 7]),
            }
            if mode == "win":
                self.data.big_win.append(rec)
            elif mode == "loss":
                self.data.big_loss.append(rec)

    # ---- 留存 ------------------------------------------------------------
    def parse_retention(self):
        if "留存数据-60月" not in self.sheets:
            return
        raw = self.xl.parse("留存数据-60月", header=None, dtype=object)
        for i in range(1, len(raw)):
            d = _to_date(raw.iloc[i, 0])
            if d is None:
                continue
            self.data.retention.append({
                "date": d, "first_dep": _to_num(raw.iloc[i, 1]),
                "first_dep_amt": _to_num(raw.iloc[i, 2]), "conv": _to_num(raw.iloc[i, 3]),
                "d2": _to_num(raw.iloc[i, 16]), "d7": _to_num(raw.iloc[i, 21]),
                "d30": _to_num(raw.iloc[i, 24]), "d60": _to_num(raw.iloc[i, 26]),
            })

    # ---- 每日投注区间 ----------------------------------------------------
    def parse_bet_range(self):
        if "每日投注区间" not in self.sheets:
            return
        raw = self.xl.parse("每日投注区间", header=None, dtype=object)
        for i in range(len(raw)):
            if isinstance(raw.iloc[i, 0], str) and "当月日均" in raw.iloc[i, 0]:
                hr = 1
                self.data.bet_range["bet_seg_avg"] = {
                    raw.iloc[hr, c]: _to_num(raw.iloc[i, c]) for c in range(1, 12)
                }
                self.data.bet_range["bet_seg_chg_prev"] = {
                    raw.iloc[hr, c]: _to_num(raw.iloc[i + 1, c]) for c in range(1, 12)
                }
                break

    def parse(self) -> ReportData:
        self.parse_ribao1()
        self.parse_venue()
        self.parse_vip()
        self.parse_bonus()
        self.parse_bigclient()
        self.parse_retention()
        self.parse_bet_range()
        return self.data


# =============================================================================
# 归因引擎
# =============================================================================

def build_attribution(d: ReportData) -> dict:
    res = {"triggered": [], "drivers": [], "drags": [], "details": {}}
    L, P = d.latest, d.prev
    if not L or not P:
        return res

    for key, label in [("gross_profit", "毛利"), ("win", "公司输赢"), ("dep_amt", "存款金额")]:
        chg = _chg(getattr(L, key), getattr(P, key))
        if chg is not None and abs(chg) > THRESHOLD:
            res["triggered"].append(f"{label} 环比 {_arrow(chg)} (阈值 ±10%)")

    venue_rank = sorted(d.venues, key=lambda v: getattr(v, "win_diff") or 0)
    worst = venue_rank[0] if venue_rank else None
    best = venue_rank[-1] if venue_rank else None
    res["details"]["venue"] = {
        "worst": worst, "best": best,
        "win_contrib": [(v.name, getattr(v, "win_diff")) for v in venue_rank
                        if not np.isnan(getattr(v, "win_diff"))],
    }

    vip_bet = d.vip_month.get("bet_amt", {})
    high_vip_chg = vip_bet.get("chg", {}).get(">=6") if "chg" in vip_bet else None
    res["details"]["vip"] = {"high_vip_bet_chg": high_vip_chg, "segment": d.vip_segment}

    bonus_chg = _chg(L.bonus, P.bonus)
    bonus_rate_l = L.bonus / L.valid_bet if (not np.isnan(L.valid_bet) and L.valid_bet) else np.nan
    bonus_rate_p = P.bonus / P.valid_bet if (not np.isnan(P.valid_bet) and P.valid_bet) else np.nan
    res["details"]["bonus"] = {
        "bonus_chg": bonus_chg, "bonus_rate_latest": bonus_rate_l, "bonus_rate_prev": bonus_rate_p,
    }

    conv_chg = _chg(L.conv, P.conv)
    first_dep_chg = _chg(L.first_dep, P.first_dep)
    res["details"]["traffic"] = {"conv_chg": conv_chg, "first_dep_chg": first_dep_chg}

    if best and getattr(best, "win_diff") and getattr(best, "win_diff") > 0:
        res["drivers"].append(
            f"【{best.name}】场馆公司输赢环比 +{_fmt(getattr(best, 'win_diff'), 'money')} "
            f"({_arrow(getattr(best, 'win_chg'))}), 有效投注 {'上升' if (getattr(best, 'vb_diff') or 0) > 0 else '下降'}"
        )
    if high_vip_chg is not None and high_vip_chg > 0:
        res["drivers"].append(f"高净值客群(>=VIP6)投注额月度环比 +{high_vip_chg * 100:.1f}%, 大客留存回升")

    if worst and getattr(worst, "win_diff") and getattr(worst, "win_diff") < 0:
        res["drags"].append(
            f"【{worst.name}】场馆公司输赢环比 {_fmt(getattr(worst, 'win_diff'), 'money')} "
            f"({_arrow(getattr(worst, 'win_chg'))}), 杀率 {_arrow(getattr(worst, 'rate_chg'))} 为主要拖累"
        )
    if bonus_rate_l is not None and bonus_rate_p is not None and (bonus_rate_l - bonus_rate_p) > 0.002:
        res["drags"].append(
            f"红利/有效投注率由 {_fmt(bonus_rate_p, 'pct')} 升至 {_fmt(bonus_rate_l, 'pct')}, "
            f"红利支出侵蚀毛利 {_fmt((bonus_rate_l - bonus_rate_p) * (L.valid_bet or 0), 'money')}"
        )
    if first_dep_chg is not None and first_dep_chg < -0.05:
        res["drags"].append(
            f"首存人数环比 {_arrow(first_dep_chg)}, 转化率 {_arrow(conv_chg)}, 新增流量转化缩水拖累存款"
        )

    return res


# =============================================================================
# 统一分析入口 (供 app.py / 脚本调用)
# =============================================================================

def analyze_workbook(file_path: str, base_date: Optional[datetime] = None) -> dict:
    """
    解析报表并返回结构化结果 dict。

    返回字段
    --------
    {
      "data": ReportData,
      "attribution": dict,
      "base_date": datetime,
      "kpi": [...],          # 供 st.metric 使用的列表
      "daily_df": DataFrame, # 规范化日表
      "venue_df": DataFrame, # 场馆表
      "vip_df": DataFrame,   # VIP 日表
      "trend_df": DataFrame, # 毛利/输赢日趋势
    }
    """
    parser = ReportParser(file_path)
    d = parser.parse()
    if d.latest is None:
        raise ValueError("未能解析到任何日度数据，请检查报表结构（需要含 ribao1 工作表）。")

    # 基准日期选择
    if base_date is not None:
        # 找到最接近 base_date 的日行
        target = None
        for r in d._daily_rows:
            if r.date.date() == base_date.date():
                target = r
                break
        if target is not None:
            d.latest = target
            idx = d._daily_rows.index(target)
            d.prev = d._daily_rows[idx - 1] if idx >= 1 else None
            lw = target.date - timedelta(days=7)
            d.lastweek = next((r for r in d._daily_rows if r.date.date() == lw.date()), None)

    att = build_attribution(d)

    # KPI 卡片
    L, P = d.latest, d.prev
    kpi_spec = [
        ("存款金额", "dep_amt", "money"),
        ("公司输赢", "win", "money"),
        ("毛利", "gross_profit", "money"),
        ("有效投注", "valid_bet", "money"),
        ("盈利率", "profit_rate", "pct"),
        ("红利金额", "bonus", "money"),
    ]
    kpi = []
    for label, attr, kind in kpi_spec:
        cur = getattr(L, attr)
        prev = getattr(P, attr) if P else np.nan
        chg = _chg(cur, prev)
        kpi.append({
            "label": label, "value": cur, "prev": prev,
            "chg": chg, "kind": kind,
            "delta": (cur - prev) if (not np.isnan(cur) and not np.isnan(prev)) else np.nan,
        })

    # 规范化 DataFrame
    daily_df = pd.DataFrame([{
        "日期": r.date.strftime("%Y-%m-%d"),
        "注册": r.reg, "首存": r.first_dep, "转化率": r.conv,
        "活跃人数": r.active, "存款人数": r.dep_people, "存款金额": r.dep_amt,
        "取款金额": r.wd_amt, "存提差": r.net_dep, "投注人数": r.bet_people,
        "有效投注": r.valid_bet, "公司输赢": r.win, "盈利率": r.profit_rate,
        "红利": r.bonus, "返水": r.rebate, "毛利": r.gross_profit,
    } for r in d._daily_rows])

    venue_df = pd.DataFrame([{
        "场馆": v.name, "公司输赢(本月)": v.win_mtd, "公司输赢(今)": v.win_today,
        "公司输赢(昨)": v.win_yest, "输赢环比": v.win_chg,
        "有效投注(今)": v.vb_today, "有效投注环比": v.vb_chg,
        "杀率(今)": v.rate_today, "杀率(昨)": v.rate_yest, "杀率环比": v.rate_chg,
    } for v in d.venues])

    vip_df = pd.DataFrame([{
        "日期": v["date"].strftime("%Y-%m-%d"),
        "存款人数": v["dep_people_total"], "投注人数": v["bet_people_total"],
        "投注额": v["bet_amt_total"],
    } for v in d.vip_daily])

    trend_df = daily_df[["日期", "毛利", "公司输赢", "有效投注", "存款金额"]].copy()

    return {
        "data": d, "attribution": att,
        "base_date": d.latest.date,
        "kpi": kpi, "daily_df": daily_df, "venue_df": venue_df,
        "vip_df": vip_df, "trend_df": trend_df,
    }


# =============================================================================
# Markdown 渲染 (供导出)
# =============================================================================

def render_markdown(result: dict) -> str:
    d = result["data"]
    att = result["attribution"]
    L, P = d.latest, d.prev
    if not L:
        return "⚠️ 无数据"
    date_str = f"{L.date.month}-{L.date.day}"
    prev_str = f"{P.date.month}-{P.date.date().day}" if P else "—"
    lw_str = f"{d.lastweek.date.month}-{d.lastweek.date.day}" if d.lastweek else "—"

    md = [f"# 📈 {d.platform} · {d.country} 实时运营归因报告",
          f"> **报告日期**: 最新日 **{date_str}** ｜ 前一日 **{prev_str}** ｜ 上周同日 **{lw_str}**",
          f"> **生成时间**: {datetime.now().strftime('%Y-%m-%d %H:%M')} ｜ 波动触发阈值: ±10%", ""]
    md.append("## 📊 1. 核心指标监控看板")
    md.append("")
    md.append("| 指标 | 最新日 | 前一日 | 环比变化 | 本月日均 |")
    md.append("|---|---|---|---|---|")
    spec = [("dep_amt", "存款金额", "money"), ("net_dep", "存提差", "money"),
            ("valid_bet", "有效投注", "money"), ("win", "公司输赢", "money"),
            ("profit_rate", "盈利率", "pct"), ("bonus", "红利金额", "money"),
            ("gross_profit", "毛利", "money")]
    for attr, label, kind in spec:
        cur = getattr(L, attr); prev = getattr(P, attr) if P else np.nan
        avg = getattr(d.mtd_avg, attr) if d.mtd_avg else np.nan
        md.append(f"| {label} | {_fmt(cur, kind)} | {_fmt(prev, kind)} | {_arrow(_chg(cur, prev))} | {_fmt(avg, kind)} |")
    md.append("")
    md.append("## 🔍 2. 为什么上涨 / 为什么下降？")
    md.append("")
    if att["triggered"]:
        md.append("**⚡ 波动阀值触发（深度归因已激活）**")
        for t in att["triggered"]:
            md.append(f"- {t}")
        md.append("")
    md.append("### ✅ 核心驱动力")
    for x in att["drivers"][:3] or ["本期无显著正向驱动因素。"]:
        md.append(f"- {x}")
    md.append("")
    md.append("### ⚠️ 主要拖累项")
    for x in att["drags"][:3] or ["本期无显著负向拖累。"]:
        md.append(f"- {x}")
    md.append("")
    md.append("**🎰 游戏场馆归因**")
    md.append("")
    md.append("| 场馆 | 公司输赢(今) | 输赢环比 | 有效投注环比 | 杀率(今) | 杀率环比 |")
    md.append("|---|---|---|---|---|---|")
    for v in sorted(d.venues, key=lambda x: getattr(x, "win_today") or 0, reverse=True):
        md.append(f"| {v.name} | {_fmt(v.win_today, 'money')} | {_arrow(v.win_chg)} | "
                  f"{_arrow(v.vb_chg)} | {_fmt(v.rate_today, 'pct')} | {_arrow(v.rate_chg)} |")
    md.append("")
    md.append("**👑 VIP / 大客归因**")
    md.append("")
    vip = att["details"]["vip"]
    if vip["high_vip_bet_chg"] is not None:
        md.append(f"- 高净值客群(>=VIP6)投注额月度环比: **{_arrow(vip['high_vip_bet_chg'])}**")
    if d.vip_segment:
        md.append("- 有效投注分段: " + " ｜ ".join(
            f"{s['seg']}({s['trend']},{_fmt(s['chg'], 'pct')})" for s in d.vip_segment))
    if d.big_loss:
        w = min(d.big_loss, key=lambda x: x["win"])
        md.append(f"- 大客亏损 TOP1: `{w['uid']}`({w['vip']}) 当日公司输赢 **{_fmt(w['win'], 'money')}**")
    if d.big_win:
        b = max(d.big_win, key=lambda x: x["win"])
        md.append(f"- 大客盈利 TOP1: `{b['uid']}`({b['vip']}) 当日贡献 {_fmt(b['win'], 'money')}")
    md.append("")
    md.append("**🎁 红利成本归因**")
    md.append("")
    bo = att["details"]["bonus"]
    if bo["bonus_chg"] is not None:
        md.append(f"- 红利金额环比: **{_arrow(bo['bonus_chg'])}**")
    if bo["bonus_rate_latest"] is not None:
        md.append(f"- 红利/有效投注率: {_fmt(bo['bonus_rate_latest'], 'pct')} vs {_fmt(bo['bonus_rate_prev'], 'pct')}")
    t = att["details"]["traffic"]
    md.append(f"- 流量/转化: 首存环比 {_arrow(t['first_dep_chg'])} ｜ 转化率环比 {_arrow(t['conv_chg'])}")
    md.append("")
    md.append("## 💡 3. 运营行动建议")
    md.append("")
    worst = att["details"]["venue"]["worst"]
    md.append("**1. 风控 / 止损建议**")
    if worst and getattr(worst, "rate_chg") is not None and getattr(worst, "rate_chg") < -0.05:
        md.append(f"- 【{worst.name}】杀率骤降至 {_fmt(worst.rate_today, 'pct')} ({_arrow(worst.rate_chg)}), "
                  f"公司输赢 {_fmt(worst.win_today, 'money')}。建议临时下调单笔下注限额 / 提高赔率审核。")
    else:
        md.append("- 各场馆杀率处于合理区间, 保持异常波动实时告警。")
    if d.big_loss:
        md.append(f"- 大客亏损榜 `{d.big_loss[0]['uid']}` 等高额赔付, 建议核实投注行为是否异常。")
    md.append("")
    md.append("**2. 营销 / 活动建议**")
    fd = t["first_dep_chg"]
    if fd is not None and fd < -0.05:
        md.append(f"- 首存人数环比 {_arrow(fd)} 缩水, 建议加投首充礼包 / 笔笔存送。")
    else:
        md.append(f"- 首存转化平稳, 关注红利率 {_fmt(bo['bonus_rate_latest'], 'pct')}, 持续>2.5% 应收缩活动。")
    md.append("")
    md.append("**3. VIP / 留存维护建议**")
    drop = None
    for s in d.vip_segment:
        if s["trend"] == "下降" and (drop is None or (s["chg"] or 0) < (drop["chg"] or 0)):
            drop = s
    if drop:
        md.append(f"- VIP 分段 **{drop['seg']}** 本月人均下跌 {_arrow(drop['chg'])}, 建议定向推送俸禄加码 / 二次存送召回。")
    if d.retention:
        recent = sorted(d.retention, key=lambda x: x["date"])[-7:]
        avg_d2 = np.nanmean([r["d2"] for r in recent if not np.isnan(r["d2"])])
        md.append(f"- 近 7 日首充 2 日留存率均值约 **{_fmt(avg_d2, 'pct')}**, 建议投关怀券提升留存。")
    md.append("")
    md.append("## 📅 4. 周度 / 月度趋势比对")
    md.append("")
    cum = {m["month"]: m for m in d.month_cum}
    keys = list(cum.keys())
    if len(keys) >= 2:
        prev_m, cur_m = cum[keys[-2]], cum[keys[-1]]

        def md_days(mrec):
            mm = re.search(r"(\d{1,2})月", str(mrec["month"]))
            return calendar.monthrange(2026, int(mm.group(1)))[1] if mm else 30

        pdays, cdays = md_days(prev_m), int(cur_m.get("days") or md_days(cur_m))
        md.append(f"**📆 MTD vs 上月日均 (本月截至{cdays}天)**")
        md.append("")
        md.append("| 指标 | 上月日均 | 本月日均 | 增幅 |")
        md.append("|---|---|---|---|")
        for k_label, attr in [("存款金额", "dep_amt"), ("有效投注", "valid_bet"),
                              ("公司输赢", "win"), ("毛利", "gross_profit"), ("红利", "bonus")]:
            pv, cv = prev_m.get(attr), cur_m.get(attr)
            p_avg = pv / pdays if not np.isnan(pv) else np.nan
            c_avg = cv / cdays if not np.isnan(cv) else np.nan
            md.append(f"| {k_label} | {_fmt(p_avg, 'money')} | {_fmt(c_avg, 'money')} | {_arrow(_chg(c_avg, p_avg))} |")
        md.append("")
    if d.bet_range.get("bet_seg_avg"):
        md.append("**🎯 玩家客群结构（投注额区间·当月日均人数）**")
        md.append("")
        md.append("| 区间 | 日均人数 | 较前天涨跌 |")
        md.append("|---|---|---|")
        for seg, val in d.bet_range["bet_seg_avg"].items():
            chg_v = d.bet_range.get("bet_seg_chg_prev", {}).get(seg)
            md.append(f"| {seg} | {_fmt(val, 'num', 0)} | {_fmt(chg_v, 'num', 0) if chg_v == chg_v else '—'} |")
        md.append("")
    md.append("---")
    md.append(f"*本报告由 GO33 自动化归因引擎生成 · 共解析 {len(d.venues)} 个场馆 / "
              f"{len(d.vip_segment)} 个 VIP 分段 / {len(d.big_win) + len(d.big_loss)} 条大客记录*")
    return "\n".join(md)


# 兼容旧入口
def analyze_daily_report(file_path: str) -> str:
    return render_markdown(analyze_workbook(file_path))
