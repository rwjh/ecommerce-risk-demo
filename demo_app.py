# -*- coding: utf-8 -*-
"""
课题二 步骤13—15：电商市场监管风险筛查与核查辅助 Demo

本文件同时承担三项职责：
1. 业务核心（可被测试脚本导入）：数据读取、契约校验、组合筛选、证据聚合、导出、复核记录；
2. 页面与交互（步骤13）：风险总览、风险商品、商品详情、复核工作台、规则中心、数据质量；
3. 复核闭环（步骤15）：复核状态独立落盘，不回写原始风险结果。

运行方式（需在已安装依赖的虚拟环境中）：
    .venv\\Scripts\\python.exe -m streamlit run demo_app.py

说明：
- 页面只消费已生成的风险成果，不在页面内重算风险规则；
- 风险分数表示核查优先级，不代表违法认定。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

BASE_DIR = Path(__file__).resolve().parent
PATHS = {
    "cleaned": BASE_DIR / "数据" / "清洗后商品数据.csv",
    "products": BASE_DIR / "结果" / "商品风险汇总.csv",
    "signals": BASE_DIR / "结果" / "风险信号明细.csv",
    "workbook": BASE_DIR / "结果" / "电商监管风险识别结果.xlsx",
}
REVIEW_PATH = BASE_DIR / "设计结果" / "人工复核记录.csv"

PRODUCT_COUNT = 4000
REVIEW_STATES = ["待核查", "已排除", "需补充材料", "确认问题线索"]
LEVEL_ORDER = ["高", "中", "低", "未命中"]
ID_COLUMNS = ["商品唯一键", "商品id"]
COLOR_MAIN = "#1F4E78"
COLOR_MID = "#2F75B5"
COLOR_LIGHT = "#9DC3E6"


# ============================================================
# 一、数据读取与契约校验
# ============================================================
@dataclass
class DataBundle:
    products: pd.DataFrame
    signals: pd.DataFrame
    cleaned: pd.DataFrame
    rules: pd.DataFrame
    platform_stats: pd.DataFrame
    category_stats: pd.DataFrame
    field_quality: pd.DataFrame
    overview: pd.DataFrame


@dataclass
class ContractReport:
    checks: list = field(default_factory=list)

    def add(self, name: str, passed: bool, evidence: str = "") -> None:
        self.checks.append({
            "检查项": name,
            "结果": "通过" if passed else "未通过",
            "依据": str(evidence),
        })

    @property
    def passed(self) -> bool:
        return all(c["结果"] == "通过" for c in self.checks)

    @property
    def failed(self) -> list:
        return [c for c in self.checks if c["结果"] != "通过"]


def load_bundle(base_dir: Path | str = BASE_DIR) -> DataBundle:
    """读取四份风险成果，统一按文本读取商品标识，避免长ID精度丢失。"""
    base = Path(base_dir)
    paths = {
        "cleaned": base / "数据" / "清洗后商品数据.csv",
        "products": base / "结果" / "商品风险汇总.csv",
        "signals": base / "结果" / "风险信号明细.csv",
        "workbook": base / "结果" / "电商监管风险识别结果.xlsx",
    }
    missing = [str(p) for p in paths.values() if not p.exists()]
    if missing:
        raise FileNotFoundError("缺少风险成果文件：\n" + "\n".join(missing))

    read_csv = lambda p: pd.read_csv(p, dtype={c: str for c in ID_COLUMNS})
    products = read_csv(paths["products"])
    signals = read_csv(paths["signals"])
    cleaned = read_csv(paths["cleaned"])
    book = paths["workbook"]
    return DataBundle(
        products=products,
        signals=signals,
        cleaned=cleaned,
        rules=pd.read_excel(book, sheet_name="风险规则说明"),
        platform_stats=pd.read_excel(book, sheet_name="平台风险统计"),
        category_stats=pd.read_excel(book, sheet_name="类别风险统计"),
        field_quality=pd.read_excel(book, sheet_name="字段质量报告"),
        overview=pd.read_excel(book, sheet_name="总体概览"),
    )


def validate_contract(bundle: DataBundle) -> ContractReport:
    """检查后续页面依赖的数据完整性、关联关系与统计口径。"""
    report = ContractReport()
    p, s, c, r = bundle.products, bundle.signals, bundle.cleaned, bundle.rules

    required = {
        "商品风险汇总": (p, {"商品唯一键", "平台", "商品id", "商品标题", "店铺名称", "风险总分", "风险等级", "命中信号数", "风险代码", "命中证据摘要"}),
        "风险信号明细": (s, {"商品唯一键", "风险代码", "风险类别", "风险信号", "基础分", "命中证据", "建议核查动作"}),
        "清洗后商品数据": (c, {"商品唯一键", "平台", "商品id", "商品标题", "到手价", "现价", "挂牌价"}),
        "风险规则说明": (r, {"风险代码", "风险类别", "风险信号", "基础分", "判定规则/阈值", "建议核查动作"}),
    }
    for name, (frame, cols) in required.items():
        lack = sorted(cols - set(frame.columns))
        report.add(f"{name}必要字段", not lack, "无缺失" if not lack else f"缺少：{lack}")

    report.add("商品汇总行数", len(p) == PRODUCT_COUNT, f"{len(p)} / 约定 {PRODUCT_COUNT}")
    report.add("清洗数据行数", len(c) == PRODUCT_COUNT, f"{len(c)} / 约定 {PRODUCT_COUNT}")
    report.add("商品唯一键唯一", not p["商品唯一键"].duplicated().any(), f"重复 {int(p['商品唯一键'].duplicated().sum())} 条")
    orphan = set(s["商品唯一键"]) - set(p["商品唯一键"])
    report.add("明细可关联商品", not orphan, f"孤立商品键 {len(orphan)} 个")
    report.add(
        "明细数与信号数一致",
        len(s) == int(p["命中信号数"].sum()),
        f"明细 {len(s)} 条 / 汇总 {int(p['命中信号数'].sum())} 条",
    )
    report.add("风险总分范围", p["风险总分"].between(0, 100).all(), f"{p['风险总分'].min()}—{p['风险总分'].max()}")
    report.add("风险等级取值", set(p["风险等级"]).issubset(set(LEVEL_ORDER)), sorted(set(p["风险等级"])))
    level_ok = (
        (p["风险等级"].eq("未命中") & p["命中信号数"].eq(0))
        | (~p["风险等级"].eq("未命中") & p["命中信号数"].gt(0))
    ).all()
    report.add("等级与命中数一致", bool(level_ok), "未命中=0，其余>0")
    report.add("规则数量", len(r) == 15 and r["风险代码"].nunique() == 15, f"{len(r)} 项 / 唯一 {r['风险代码'].nunique()}")
    return report


# ============================================================
# 二、复核记录（独立存储，不回写原始结果）
# ============================================================
REVIEW_COLUMNS = ["商品唯一键", "复核状态", "复核原因", "复核人", "更新时间"]


class ReviewStore:
    """人工复核记录的读写。原始风险成果保持只读。"""

    def __init__(self, path: Path | str = REVIEW_PATH):
        self.path = Path(path)

    def load(self) -> pd.DataFrame:
        if not self.path.exists() or self.path.stat().st_size == 0:
            return pd.DataFrame(columns=REVIEW_COLUMNS)
        df = pd.read_csv(self.path, dtype=str).fillna("")
        for col in REVIEW_COLUMNS:
            if col not in df.columns:
                df[col] = ""
        df = df[REVIEW_COLUMNS]
        return df.drop_duplicates("商品唯一键", keep="last").reset_index(drop=True)

    def status_map(self) -> dict:
        df = self.load()
        return dict(zip(df["商品唯一键"], df["复核状态"]))

    def upsert(self, key: str, status: str, reason: str = "", reviewer: str = "") -> pd.DataFrame:
        if status not in REVIEW_STATES:
            raise ValueError(f"复核状态必须属于 {REVIEW_STATES}")
        if status != "待核查" and not str(reason).strip():
            raise ValueError("非“待核查”状态必须填写复核原因")
        current = self.load()
        record = pd.DataFrame([{
            "商品唯一键": str(key),
            "复核状态": status,
            "复核原因": str(reason).strip(),
            "复核人": str(reviewer).strip() or "未填写",
            "更新时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }])
        merged = pd.concat([current[current["商品唯一键"] != str(key)], record], ignore_index=True)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        merged.to_csv(self.path, index=False, encoding="utf-8-sig")
        return merged

    def summary(self) -> dict:
        df = self.load()
        counts = {state: int(df["复核状态"].eq(state).sum()) for state in REVIEW_STATES}
        return counts


def attach_review_status(products: pd.DataFrame, status_map: dict) -> pd.DataFrame:
    out = products.copy()
    out["复核状态"] = out["商品唯一键"].map(lambda k: status_map.get(k, "待核查"))
    return out


# ============================================================
# 三、筛选、排序与证据聚合
# ============================================================
SORT_OPTIONS = {
    "风险总分（高→低）": ("风险总分", False),
    "命中信号数（多→少）": ("命中信号数", False),
    "月销量（高→低）": ("月销量", False),
    "到手价（高→低）": ("到手价", False),
    "商品标题（升序）": ("商品标题", True),
}


@dataclass
class FilterState:
    platforms: list = field(default_factory=list)
    levels: list = field(default_factory=list)
    categories: list = field(default_factory=list)
    codes: list = field(default_factory=list)
    review_states: list = field(default_factory=list)
    price_min: float = 0.0
    price_max: float = 1e9
    min_sales: float = 0.0
    live: str = "全部"
    cross: str = "全部"
    keyword: str = ""
    sort_label: str = "风险总分（高→低）"

    def describe(self) -> str:
        parts = []
        if self.platforms:
            parts.append("平台=" + "/".join(self.platforms))
        if self.levels:
            parts.append("等级=" + "/".join(self.levels))
        if self.categories:
            parts.append("类别=" + "/".join(self.categories))
        if self.codes:
            parts.append("代码=" + "/".join(self.codes))
        if self.review_states:
            parts.append("复核=" + "/".join(self.review_states))
        if self.price_min > 0 or self.price_max < 1e8:
            parts.append(f"到手价={self.price_min:g}—{self.price_max:g}元")
        if self.min_sales > 0:
            parts.append(f"月销量≥{self.min_sales:g}")
        if self.live != "全部":
            parts.append("直播=" + self.live)
        if self.cross != "全部":
            parts.append("跨境=" + self.cross)
        if self.keyword.strip():
            parts.append("关键词=" + self.keyword.strip())
        parts.append("排序=" + self.sort_label)
        return "；".join(parts)


def _tag_filter(series: pd.Series, selected: list) -> pd.Series:
    """多值字段（以顿号分隔）按“任一命中”筛选。"""
    if not selected:
        return pd.Series(True, index=series.index)
    wanted = set(selected)
    return series.fillna("").astype(str).map(
        lambda text: bool(wanted & {t for t in text.split("、") if t})
    )


def apply_filters(products: pd.DataFrame, state: FilterState) -> pd.DataFrame:
    """按筛选状态返回商品集合，并完成排序。"""
    df = products
    if state.platforms:
        df = df[df["平台"].isin(state.platforms)]
    if state.levels:
        df = df[df["风险等级"].isin(state.levels)]
    df = df[_tag_filter(df["风险类别"], state.categories)]
    df = df[_tag_filter(df["风险代码"], state.codes)]
    if state.review_states:
        df = df[df["复核状态"].isin(state.review_states)]
    df = df[df["到手价"].between(state.price_min, state.price_max, inclusive="both")]
    df = df[df["月销量"].fillna(0) >= state.min_sales]
    if state.live != "全部":
        df = df[df["是否是直播商品"] == ("t" if state.live == "是" else "f")]
    if state.cross != "全部":
        df = df[df["是否跨境"] == ("t" if state.cross == "是" else "f")]
    keyword = state.keyword.strip()
    if keyword:
        haystack = (
            df["商品标题"].fillna("") + " "
            + df["店铺名称"].fillna("") + " "
            + df["品牌"].fillna("") + " "
            + df["商品id"].astype(str)
        )
        df = df[haystack.str.contains(keyword, case=False, regex=False)]
    column, ascending = SORT_OPTIONS.get(state.sort_label, ("风险总分", False))
    df = df.sort_values([column, "商品唯一键"], ascending=[ascending, True])
    return df.reset_index(drop=True)


def get_product_evidence(bundle: DataBundle, key: str) -> dict:
    """聚合单个商品的原始信息、全部风险信号和规则阈值。"""
    row = bundle.products[bundle.products["商品唯一键"] == key]
    if row.empty:
        raise KeyError(f"商品不存在：{key}")
    signals = bundle.signals[bundle.signals["商品唯一键"] == key].copy()
    if not signals.empty:
        signals = signals.merge(
            bundle.rules[["风险代码", "判定规则/阈值"]], on="风险代码", how="left"
        ).sort_values("基础分", ascending=False).reset_index(drop=True)
    return {"product": row.iloc[0], "signals": signals}


def compute_overview(products: pd.DataFrame) -> dict:
    return {
        "商品总数": int(len(products)),
        "命中风险商品": int(products["命中信号数"].gt(0).sum()),
        "高风险": int(products["风险等级"].eq("高").sum()),
        "中风险": int(products["风险等级"].eq("中").sum()),
        "低风险": int(products["风险等级"].eq("低").sum()),
        "未命中": int(products["风险等级"].eq("未命中").sum()),
        "风险信号总数": int(products["命中信号数"].sum()),
    }


def platform_summary(products: pd.DataFrame) -> pd.DataFrame:
    stats = products.groupby("平台", sort=False).agg(
        商品数=("商品唯一键", "size"),
        命中商品数=("命中信号数", lambda x: int(x.gt(0).sum())),
        高风险数=("风险等级", lambda x: int(x.eq("高").sum())),
    ).reset_index()
    stats["命中率"] = (stats["命中商品数"] / stats["商品数"] * 100).round(2)
    return stats


def category_summary(products: pd.DataFrame) -> pd.DataFrame:
    hit = products[products["命中信号数"].gt(0)].copy()
    hit["类别列表"] = hit["风险类别"].fillna("").astype(str).map(
        lambda t: [x for x in t.split("、") if x]
    )
    exploded = hit.explode("类别列表")
    exploded = exploded[exploded["类别列表"].notna() & exploded["类别列表"].ne("")]
    stats = exploded.groupby("类别列表").agg(
        涉及商品数=("商品唯一键", "nunique"),
        高风险商品数=("风险等级", lambda x: int(x.eq("高").sum())),
    ).reset_index().rename(columns={"类别列表": "风险类别"})
    return stats.sort_values("涉及商品数", ascending=False).reset_index(drop=True)


def build_product_export(products: pd.DataFrame, filter_desc: str, exported_at: str) -> pd.DataFrame:
    columns = [
        "商品唯一键", "平台", "商品id", "商品标题", "店铺名称", "品牌", "标准品牌",
        "平台一级分类名称", "最后一级分类名称", "到手价", "现价", "挂牌价", "月销量",
        "商品评论数", "风险总分", "风险等级", "命中信号数", "风险代码", "风险信号",
        "命中证据摘要", "建议核查动作", "复核状态", "商品Url",
    ]
    out = products[[c for c in columns if c in products.columns]].copy()
    out["筛选口径"] = filter_desc
    out["导出时间"] = exported_at
    return out


def build_signal_export(signals: pd.DataFrame, keys: set, filter_desc: str, exported_at: str) -> pd.DataFrame:
    columns = [
        "商品唯一键", "平台", "商品id", "商品标题", "店铺名称", "品牌",
        "风险代码", "风险类别", "风险信号", "基础分", "命中证据", "建议核查动作",
    ]
    out = signals[signals["商品唯一键"].isin(keys)][[c for c in columns if c in signals.columns]].copy()
    out["筛选口径"] = filter_desc
    out["导出时间"] = exported_at
    return out


def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")


def truncate(text: str, limit: int) -> str:
    text = "" if text is None or pd.isna(text) else str(text)
    return text if len(text) <= limit else text[: limit - 1] + "…"


def yes_no(value) -> str:
    return {"t": "是", "f": "否"}.get(str(value), "未知")


def fmt_money(value) -> str:
    return "缺失" if pd.isna(value) else f"{float(value):,.2f} 元"


def text_or_dash(value) -> str:
    """空值、NaN 和纯空白统一显示为“缺失”，避免把 nan 当作文本展示。"""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return "缺失"
    text = str(value).strip()
    return text if text and text.lower() != "nan" else "缺失"


# ============================================================
# 四、页面与交互（步骤13）
# ============================================================
PAGES = ["风险总览", "风险商品", "选中商品", "商品详情", "复核工作台", "规则中心", "数据质量", "关于本产品"]


def inject_style() -> None:
    st.markdown(
        """
        <style>
        .block-container {padding-top: 2.1rem; padding-bottom: 2rem; max-width: 1500px;}
        h1, h2, h3, h4 {font-family: "Microsoft YaHei", "PingFang SC", sans-serif; color: #5B21B6;}
        [data-testid="stMetricValue"] {font-size: 1.55rem; color: #4C1D95;}
        [data-testid="stMetricLabel"] {font-size: 0.85rem; color: #555;}
        .app-header {background: linear-gradient(135deg, #1F4E78 0%, #5B21B6 100%);
                     color: #fff; padding: 1.1rem 1.4rem;
                     border-radius: 10px; margin-bottom: 1rem;}
        .app-header .t {font-size: 1.5rem; font-weight: 700; letter-spacing: .5px;}
        .app-header .s {font-size: .88rem; opacity: .9; margin-top: .35rem; line-height: 1.5;}
        .note {background: #F3F6F9; border-left: 4px solid #7C3AED; padding: .7rem .9rem;
               font-size: .85rem; color: #33475B; border-radius: 0 6px 6px 0;}
        /* 主要操作按钮：紫色 */
        .stButton > button[kind="primary"],
        .stFormSubmitButton > button[kind="primary"],
        button[data-testid="stBaseButton-primary"],
        button[data-testid="stBaseButton-primaryFormSubmit"] {
            background-color: #7C3AED !important;
            border-color: #7C3AED !important;
            color: #FFFFFF !important;
            font-weight: 600;
        }
        .stButton > button[kind="primary"]:hover,
        .stFormSubmitButton > button[kind="primary"]:hover,
        button[data-testid="stBaseButton-primary"]:hover,
        button[data-testid="stBaseButton-primaryFormSubmit"]:hover {
            background-color: #6D28D9 !important;
            border-color: #6D28D9 !important;
        }
        /* 下载按钮：紫色描边 */
        .stDownloadButton > button, button[data-testid="stBaseButton-secondary"] {
            border-color: #7C3AED !important;
            color: #5B21B6 !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def goto(page: str, key: str | None = None) -> None:
    """作为按钮的 on_click 回调使用；回调执行后 Streamlit 会自动重跑。"""
    st.session_state["page"] = page
    if key:
        st.session_state["selected_key"] = key


def clear_selected() -> None:
    """清空批量选择并返回列表页，同时重置勾选状态。"""
    st.session_state["selected_keys"] = []
    st.session_state["table_token"] = st.session_state.get("table_token", 0) + 1
    st.session_state["page"] = "风险商品"


def header(title: str, subtitle: str) -> None:
    st.markdown(
        f'<div class="app-header"><div class="t">{title}</div><div class="s">{subtitle}</div></div>',
        unsafe_allow_html=True,
    )


def note(text: str) -> None:
    st.markdown(f'<div class="note">{text}</div>', unsafe_allow_html=True)


def bar_chart(df: pd.DataFrame, x: str, y: str, title: str, horizontal: bool = False) -> go.Figure:
    if horizontal:
        fig = go.Figure(go.Bar(x=df[y], y=df[x], orientation="h", marker_color=COLOR_MID))
        fig.update_layout(yaxis=dict(autorange="reversed"))
    else:
        fig = go.Figure(go.Bar(x=df[x], y=df[y], marker_color=COLOR_MID))
    fig.update_layout(
        title=dict(text=title, font=dict(size=14)),
        height=310,
        margin=dict(l=10, r=10, t=45, b=10),
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Microsoft YaHei"),
    )
    return fig


def page_overview(bundle: DataBundle) -> None:
    products = attach_review_status(bundle.products, bundle_review_map())
    stats = compute_overview(products)
    header(
        "风险总览",
        "先看整体规模与分布，再进入具体商品核查。所有结果均为待核查线索，分数代表核查优先级。",
    )
    cols = st.columns(5)
    cols[0].metric("商品总数", f"{stats['商品总数']:,}")
    cols[1].metric("命中风险商品", f"{stats['命中风险商品']:,}", f"{stats['命中风险商品']/stats['商品总数']*100:.2f}%")
    cols[2].metric("高风险", stats["高风险"])
    cols[3].metric("中风险", stats["中风险"])
    cols[4].metric("低风险", stats["低风险"])

    left, right = st.columns(2)
    pstats = platform_summary(products)
    with left:
        st.plotly_chart(bar_chart(pstats, "平台", "命中率", "各平台风险商品命中率（%）"), width="stretch")
    with right:
        cstats = category_summary(products)
        st.plotly_chart(
            bar_chart(cstats, "风险类别", "涉及商品数", "各类风险涉及商品数", horizontal=True),
            width="stretch",
        )

    st.subheader("高风险商品（优先核查）")
    high = products[products["风险等级"].eq("高")].sort_values("风险总分", ascending=False)
    if high.empty:
        st.info("当前没有高风险商品。")
    else:
        head = st.columns([2.4, 3.4, 1.6, 1.2, 1.0])
        for col, text in zip(head, ["平台", "商品标题", "命中信号", "总分", "操作"]):
            col.caption(text)
        for _, row in high.iterrows():
            c = st.columns([2.4, 3.4, 1.6, 1.2, 1.0])
            c[0].write(row["平台"])
            c[1].write(truncate(row["商品标题"], 34))
            c[2].write(truncate(row["风险代码"], 16))
            c[3].write(int(row["风险总分"]))
            c[4].button("查看", key=f"hr_{row['商品唯一键']}", on_click=goto, args=("商品详情", row["商品唯一键"]))

    note(
        "解读边界：未命中仅表示当前15项规则未发现异常，不等于商品合规；高风险表示核查优先级较高，"
        "不构成违法认定。命中率差异受各平台类目结构和字段完整度影响。"
    )

    st.subheader("平台风险统计")
    st.dataframe(
        pstats.rename(columns={"商品数": "商品数", "命中商品数": "命中商品数", "高风险数": "高风险商品数"}),
        hide_index=True,
        width="stretch",
        column_config={
            "命中率": st.column_config.ProgressColumn("命中率（%）", min_value=0, max_value=100, format="%.2f"),
        },
    )


def _fkey(name: str) -> str:
    """筛选控件键。重置条件时更换令牌，使控件回到默认值。"""
    return f"f_{name}_{st.session_state.get('filter_token', 0)}"


def build_filter_panel(bundle: DataBundle) -> FilterState:
    """筛选面板：表单内修改不会立即刷新，点击“应用筛选”后才更新结果。"""
    products = bundle.products
    cat_options = sorted({x for t in products["风险类别"].fillna("") for x in str(t).split("、") if x})
    code_options = sorted({x for t in products["风险代码"].fillna("") for x in str(t).split("、") if x})
    max_price = float(products["到手价"].max())

    with st.expander("筛选条件", expanded=True):
        with st.form("filter_form", border=False):
            r1 = st.columns(4)
            platforms = r1[0].multiselect(
                "平台", sorted(products["平台"].unique()),
                default=sorted(products["平台"].unique()), key=_fkey("platform"),
            )
            levels = r1[1].multiselect("风险等级", LEVEL_ORDER, default=["高", "中", "低"], key=_fkey("level"))
            categories = r1[2].multiselect("风险类别", cat_options, default=[], key=_fkey("category"))
            codes = r1[3].multiselect("风险代码", code_options, default=[], key=_fkey("code"))

            r2 = st.columns(5)
            price_min, price_max_sel = r2[0].slider(
                "到手价区间（元）", 0.0, max_price, (0.0, max_price), key=_fkey("price"),
            )
            min_sales = r2[1].number_input("最低月销量", min_value=0, value=0, step=10, key=_fkey("sales"))
            live = r2[2].selectbox("是否直播商品", ["全部", "是", "否"], key=_fkey("live"))
            cross = r2[3].selectbox("是否跨境商品", ["全部", "是", "否"], key=_fkey("cross"))
            review_states = r2[4].multiselect("复核状态", REVIEW_STATES, default=REVIEW_STATES, key=_fkey("review"))

            r3 = st.columns([2.2, 1.6, 1.8])
            keyword = r3[0].text_input("关键词", placeholder="商品标题 / 店铺 / 品牌 / 商品ID", key=_fkey("keyword"))
            sort_label = r3[1].selectbox("排序", list(SORT_OPTIONS), key=_fkey("sort"))
            r3[2].caption("修改条件后不会立即刷新，请点击下方“应用筛选”。")

            b1, b2, _ = st.columns([1, 1, 3.6])
            applied = b1.form_submit_button("应用筛选", type="primary")
            reset = b2.form_submit_button("重置条件")

        if reset:
            st.session_state["filter_token"] = st.session_state.get("filter_token", 0) + 1
            st.rerun()
        if applied:
            # 应用新条件时更换表格键，使旧的行选择自动失效，避免索引错位
            st.session_state["table_token"] = st.session_state.get("table_token", 0) + 1
            st.caption("已按当前条件更新结果。")

    return FilterState(
        platforms=platforms, levels=levels, categories=categories, codes=codes,
        review_states=review_states, price_min=price_min, price_max=price_max_sel,
        min_sales=float(min_sales), live=live, cross=cross, keyword=keyword,
        sort_label=sort_label,
    )


def page_products(bundle: DataBundle) -> None:
    status_map = bundle_review_map()
    products = attach_review_status(bundle.products, status_map)
    header(
        "风险商品",
        "按平台、等级、类别、价格、销量和复核状态组合筛选，形成当前核查清单；勾选多个商品后点“确定选择”。",
    )
    state = build_filter_panel(bundle)
    filtered = apply_filters(products, state)

    m = st.columns(4)
    m[0].metric("当前结果", f"{len(filtered):,}")
    m[1].metric("其中高风险", int(filtered["风险等级"].eq("高").sum()))
    m[2].metric("涉及风险信号", int(filtered["命中信号数"].sum()))
    m[3].metric("涉及平台", filtered["平台"].nunique())
    st.caption("当前筛选口径：" + state.describe())

    if filtered.empty:
        st.info("没有满足当前条件的商品。请放宽价格区间、等级或关键词条件。")
        return

    display = pd.DataFrame({
        "平台": filtered["平台"],
        "商品标题": filtered["商品标题"].map(lambda x: truncate(x, 30)),
        "店铺名称": filtered["店铺名称"].map(lambda x: truncate(x, 16)),
        "到手价": filtered["到手价"],
        "月销量": filtered["月销量"],
        "风险总分": filtered["风险总分"],
        "风险等级": filtered["风险等级"],
        "信号数": filtered["命中信号数"],
        "风险信号": filtered["风险信号"].map(lambda x: truncate(x, 40)),
        "复核状态": filtered["复核状态"],
        "商品唯一键": filtered["商品唯一键"],
    })
    warning = st.session_state.pop("selection_warning", None)
    if warning:
        st.warning(warning)

    editor_key = f"select_editor_{st.session_state.get('table_token', 0)}"
    editor_df = display.drop(columns=["商品唯一键"]).reset_index(drop=True).copy()
    editor_df.insert(0, "选择", False)

    with st.form("select_form", border=False):
        edited = st.data_editor(
            editor_df,
            hide_index=True,
            width="stretch",
            height=430,
            num_rows="fixed",
            key=editor_key,
            disabled=[c for c in editor_df.columns if c != "选择"],
            column_config={
                "选择": st.column_config.CheckboxColumn("选择", default=False),
                "到手价": st.column_config.NumberColumn(format="%.2f"),
                "月销量": st.column_config.NumberColumn(format="%d"),
            },
        )
        bf1, bf2, bf3, _ = st.columns([1.3, 1.3, 1.3, 3.2])
        confirmed = bf1.form_submit_button("✓ 确定选择", type="primary")
        view_detail = bf2.form_submit_button("查看选中商品详情")
        clear_sel = bf3.form_submit_button("清空勾选")

    if confirmed:
        picked = [i for i, v in edited["选择"].items() if bool(v)]
        keys = [filtered.iloc[i]["商品唯一键"] for i in picked if 0 <= i < len(filtered)]
        if keys:
            st.session_state["selected_keys"] = keys
            st.session_state["selection_warning"] = None
        else:
            st.session_state["selection_warning"] = "请先勾选至少一个商品，再点击“确定选择”。"
        st.rerun()

    if view_detail:
        picked = [i for i, v in edited["选择"].items() if bool(v)]
        keys = [filtered.iloc[i]["商品唯一键"] for i in picked if 0 <= i < len(filtered)]
        if keys:
            st.session_state["selected_keys"] = keys
            st.session_state["page"] = "选中商品"
        else:
            st.session_state["selection_warning"] = "请先勾选至少一个商品，再点击“查看选中商品详情”。"
        st.rerun()

    if clear_sel:
        st.session_state["selected_keys"] = []
        st.session_state["table_token"] = st.session_state.get("table_token", 0) + 1
        st.rerun()

    st.caption("操作：勾选过程不会刷新页面；勾选完多个商品后点击“确定选择”，下方会展示已选商品并可批量导出。")

    # ---------- 已选商品面板 ----------
    selected_keys = [k for k in st.session_state.get("selected_keys", [])
                     if k in set(filtered["商品唯一键"])]
    if selected_keys:
        st.divider()
        st.markdown("""<div style="background:#FAF5FF;border-left:4px solid #7C3AED;padding:.7rem 1rem;border-radius:6px;margin:.5rem 0;">
            <b style="color:#5B21B6;">📌 已选商品面板</b>
            <span style="margin-left:.8rem;color:#666;font-size:.9rem;">已确认选择 <b style="color:#7C3AED;">{}个</b> 商品</span>
        </div>""".format(len(selected_keys)), unsafe_allow_html=True)

        sel_df = filtered[filtered["商品唯一键"].isin(selected_keys)].copy()
        order = {k: i for i, k in enumerate(selected_keys)}
        sel_df["_order"] = sel_df["商品唯一键"].map(order)
        sel_df = sel_df.sort_values("_order").drop(columns=["_order"]).reset_index(drop=True)

        m2 = st.columns(4)
        m2[0].metric("已选商品", len(sel_df))
        m2[1].metric("其中高风险", int(sel_df["风险等级"].eq("高").sum()))
        m2[2].metric("涉及风险信号", int(sel_df["命中信号数"].sum()))
        m2[3].metric("涉及平台", sel_df["平台"].nunique())

        st.dataframe(
            pd.DataFrame({
                "平台": sel_df["平台"],
                "商品标题": sel_df["商品标题"].map(lambda x: truncate(x, 28)),
                "店铺名称": sel_df["店铺名称"].map(lambda x: truncate(x, 14)),
                "到手价": sel_df["到手价"],
                "月销量": sel_df["月销量"],
                "风险总分": sel_df["风险总分"],
                "风险等级": sel_df["风险等级"],
                "信号数": sel_df["命中信号数"],
                "复核状态": sel_df["复核状态"],
            }),
            hide_index=True, width="stretch",
            column_config={
                "到手价": st.column_config.NumberColumn(format="%.2f"),
                "月销量": st.column_config.NumberColumn(format="%d"),
            },
        )

        stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        scope = f"勾选商品 {len(selected_keys)} 个"
        signal_count = int(bundle.signals["商品唯一键"].isin(selected_keys).sum())

        st.markdown("""<style>
        .btn-purple-export button {background-color:#7C3AED !important;border-color:#7C3AED !important;color:#fff !important;font-weight:600 !important;}
        .btn-purple-export button:hover {background-color:#6D28D9 !important;border-color:#6D28D9 !important;}
        </style>""", unsafe_allow_html=True)

        ec1, ec2, ec3, _ = st.columns([2.2, 2.2, 1.4, 2.2])
        with ec1:
            st.markdown('<div class="btn-purple-export">', unsafe_allow_html=True)
            st.download_button(
                f"📋 导出选中商品详细信息（{len(selected_keys)}）",
                data=to_csv_bytes(build_product_export(sel_df, scope, stamp)),
                file_name=f"选中商品详细信息_{datetime.now():%Y%m%d_%H%M}.csv",
                mime="text/csv",
                key="dl_sel_detail",
                use_container_width=True,
            )
            st.markdown('</div>', unsafe_allow_html=True)
        with ec2:
            st.markdown('<div class="btn-purple-export">', unsafe_allow_html=True)
            st.download_button(
                f"⚠️ 导出选中风险明细（{signal_count}条）",
                data=to_csv_bytes(build_signal_export(bundle.signals, set(selected_keys), scope, stamp)),
                file_name=f"选中商品风险明细_{datetime.now():%Y%m%d_%H%M}.csv",
                mime="text/csv",
                key="dl_sel_detail_signals",
                use_container_width=True,
            )
            st.markdown('</div>', unsafe_allow_html=True)
        ec3.button("查看全部证据 →", type="primary", use_container_width=True,
                   key="goto_selected_page", on_click=goto, args=("选中商品",))

    # ---------- 筛选结果导出按钮（底部） ----------
    st.divider()
    exported_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    col_b, col_c, col_d = st.columns([1.6, 1.6, 3.6])
    col_b.download_button(
        f"导出筛选清单（{len(filtered):,}）",
        data=to_csv_bytes(build_product_export(filtered, state.describe(), exported_at)),
        file_name=f"风险商品清单_{datetime.now():%Y%m%d_%H%M}.csv",
        mime="text/csv",
        key="dl_products",
    )
    col_c.download_button(
        "导出筛选明细",
        data=to_csv_bytes(build_signal_export(bundle.signals, set(filtered["商品唯一键"]), state.describe(), exported_at)),
        file_name=f"风险信号明细_{datetime.now():%Y%m%d_%H%M}.csv",
        mime="text/csv",
        key="dl_signals",
    )
    col_d.caption("此处导出当前筛选全部结果；若只想导出勾选的几个商品，请在上方“已选商品面板”操作。")
    note("导出文件包含当前筛选口径与导出时间，并保留命中证据，便于后续人工核查与任务分派。")


def page_selected(bundle: DataBundle) -> None:
    """批量查看已勾选商品：对比信息、逐商品证据、批量导出。"""
    header("选中商品", "同时查看已勾选商品的对比信息与全部风险证据，并把多个商品一起导出。")
    all_keys = set(bundle.products["商品唯一键"])
    keys = [k for k in st.session_state.get("selected_keys", []) if k in all_keys]
    st.session_state["selected_keys"] = keys

    if not keys:
        st.info("尚未选择商品。请先到“风险商品”页勾选商品，再点击“查看选中商品信息”。")
        st.button("返回风险商品", type="primary", key="btn_back_products", on_click=goto, args=("风险商品",))
        return

    status_map = bundle_review_map()
    selected = bundle.products[bundle.products["商品唯一键"].isin(keys)].copy()
    selected["复核状态"] = selected["商品唯一键"].map(lambda k: status_map.get(k, "待核查"))
    order = {k: i for i, k in enumerate(keys)}
    selected["_order"] = selected["商品唯一键"].map(order)
    selected = selected.sort_values("_order").drop(columns=["_order"]).reset_index(drop=True)

    m = st.columns(4)
    m[0].metric("选中商品", len(selected))
    m[1].metric("其中高风险", int(selected["风险等级"].eq("高").sum()))
    m[2].metric("涉及风险信号", int(selected["命中信号数"].sum()))
    m[3].metric("涉及平台", selected["平台"].nunique())

    st.subheader("商品对比")
    st.dataframe(
        pd.DataFrame({
            "平台": selected["平台"],
            "商品标题": selected["商品标题"].map(lambda x: truncate(x, 26)),
            "店铺名称": selected["店铺名称"].map(lambda x: truncate(x, 14)),
            "到手价": selected["到手价"],
            "月销量": selected["月销量"],
            "风险总分": selected["风险总分"],
            "风险等级": selected["风险等级"],
            "信号数": selected["命中信号数"],
            "复核状态": selected["复核状态"],
        }),
        hide_index=True,
        width="stretch",
        column_config={
            "到手价": st.column_config.NumberColumn(format="%.2f"),
            "月销量": st.column_config.NumberColumn(format="%d"),
        },
    )

    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    scope = f"勾选商品 {len(keys)} 个"
    signal_count = int(bundle.signals["商品唯一键"].isin(keys).sum())
    d1, d2, d3 = st.columns([1.9, 1.9, 2.2])
    d1.download_button(
        f"导出选中清单（{len(keys)}）",
        data=to_csv_bytes(build_product_export(selected, scope, stamp)),
        file_name=f"选中商品清单_{datetime.now():%Y%m%d_%H%M}.csv",
        mime="text/csv",
        key="dl_sel_products",
    )
    d2.download_button(
        f"导出选中明细（{signal_count}条）",
        data=to_csv_bytes(build_signal_export(bundle.signals, set(keys), scope, stamp)),
        file_name=f"选中商品风险明细_{datetime.now():%Y%m%d_%H%M}.csv",
        mime="text/csv",
        key="dl_sel_signals",
    )
    d3.button("清空选择并返回", key="btn_clear_selected", on_click=clear_selected)

    st.subheader("商品信息与风险证据")
    expand_all = st.checkbox("默认展开全部证据", value=True, key="sel_expand_all")
    for key in keys:
        ev = get_product_evidence(bundle, key)
        row, sigs = ev["product"], ev["signals"]
        title = f"{row['平台']}｜{truncate(row['商品标题'], 36)}｜{row['风险等级']}｜{int(row['风险总分'])}分"
        with st.expander(title, expanded=expand_all):
            c = st.columns(6)
            c[0].metric("到手价", fmt_money(row["到手价"]))
            c[1].metric("现价", fmt_money(row["现价"]))
            c[2].metric("月销量", f"{int(row['月销量']):,}" if pd.notna(row["月销量"]) else "缺失")
            c[3].metric("评论数", "缺失" if pd.isna(row["商品评论数"]) else f"{int(row['商品评论数']):,}")
            c[4].metric("风险等级", row["风险等级"])
            c[5].metric("复核状态", status_map.get(key, "待核查"))
            st.caption(
                f"店铺：{text_or_dash(row['店铺名称'])}　品牌：{text_or_dash(row['品牌'])}　"
                f"类目：{text_or_dash(row['最后一级分类名称'])}　商品ID：{row['商品id']}"
            )
            if sigs.empty:
                st.info("该商品未命中任何规则。")
            for _, sig in sigs.iterrows():
                st.markdown(f"**{sig['风险代码']}　{sig['风险信号']}（{int(sig['基础分'])} 分）**")
                st.write(f"命中规则：{text_or_dash(sig.get('判定规则/阈值'))}")
                st.write(f"命中证据：{text_or_dash(sig['命中证据'])}")
                st.write(f"建议核查动作：{text_or_dash(sig['建议核查动作'])}")
                st.divider()

    st.caption("本页用于批量查看与批量导出；如需登记人工复核，请到“商品详情”页操作。")


def page_detail(bundle: DataBundle) -> None:
    header("商品详情", "逐条查看命中证据、规则阈值与建议核查动作，并登记人工复核结果。")
    flash = st.session_state.pop("flash", None)
    if flash:
        st.success(flash)

    status_map = bundle_review_map()
    keys = list(bundle.products["商品唯一键"])
    default = st.session_state.get("selected_key")
    index = keys.index(default) if default in keys else 0
    labels = dict(zip(bundle.products["商品唯一键"], bundle.products["商品标题"]))

    chosen = st.selectbox(
        "选择商品（也可从风险商品列表点击进入）",
        keys,
        index=index,
        format_func=lambda k: f"{truncate(labels.get(k, k), 40)} | {k}",
    )
    st.session_state["selected_key"] = chosen

    try:
        evidence = get_product_evidence(bundle, chosen)
    except KeyError as exc:
        st.error(str(exc))
        return
    row, signals = evidence["product"], evidence["signals"]

    top = st.columns([4, 1.2, 1.2])
    top[0].subheader(row["商品标题"])
    top[1].metric("风险总分", int(row["风险总分"]))
    top[2].metric("风险等级", row["风险等级"])
    st.caption(
        f"平台：{row['平台']}　店铺：{text_or_dash(row['店铺名称'])}　品牌：{text_or_dash(row['品牌'])}　"
        f"类目：{text_or_dash(row['最后一级分类名称'])}　复核状态：{status_map.get(chosen, '待核查')}"
    )

    m = st.columns(6)
    m[0].metric("到手价", fmt_money(row["到手价"]))
    m[1].metric("现价", fmt_money(row["现价"]))
    m[2].metric("挂牌价", fmt_money(row["挂牌价"]))
    m[3].metric("月销量", f"{int(row['月销量']):,}" if pd.notna(row["月销量"]) else "缺失")
    m[4].metric("评论数", "缺失" if pd.isna(row["商品评论数"]) else f"{int(row['商品评论数']):,}")
    m[5].metric("好评率", "缺失" if pd.isna(row["好评率"]) else f"{row['好评率']:.2f}%")

    if pd.notna(row["现价"]) and row["现价"]:
        st.caption(
            f"价格关系：到手价 / 现价 = {row['到手价']/row['现价']*100:.2f}%；"
            f"唯一键 {row['商品唯一键']}（商品ID {row['商品id']}）"
        )

    st.subheader(f"风险证据（{len(signals)} 条）")
    if signals.empty:
        st.info("该商品当前未命中任何规则，仅表示15项规则未发现异常。")
    else:
        for i, sig in signals.iterrows():
            with st.expander(f"{sig['风险代码']}　{sig['风险信号']}　（{int(sig['基础分'])} 分）", expanded=(i == 0)):
                c1, c2 = st.columns([1.1, 2.4])
                c1.write(f"**风险类别：** {sig['风险类别']}")
                c1.write(f"**命中规则：** {text_or_dash(sig.get('判定规则/阈值'))}")
                c2.write(f"**命中证据：** {text_or_dash(sig['命中证据'])}")
                c2.write(f"**建议核查动作：** {text_or_dash(sig['建议核查动作'])}")

    with st.expander("商品基本信息"):
        info = pd.DataFrame({
            "字段": ["平台", "商品ID", "店铺名称", "品牌", "标准品牌", "平台一级分类", "最后一级分类",
                     "是否在售", "是否跨境", "是否直播商品", "品牌企业统一编码", "商品链接"],
            "值": [
                row["平台"], row["商品id"], text_or_dash(row["店铺名称"]), text_or_dash(row["品牌"]),
                text_or_dash(row["标准品牌"]), text_or_dash(row["平台一级分类名称"]), text_or_dash(row["最后一级分类名称"]),
                yes_no(row["是否在售"]), yes_no(row["是否跨境"]), yes_no(row["是否是直播商品"]),
                text_or_dash(row["品牌企业统一编码"]), text_or_dash(row["商品Url"]),
            ],
        })
        st.dataframe(info, hide_index=True, width="stretch")

    st.subheader("人工复核")
    note("复核记录独立保存于“设计结果/人工复核记录.csv”，不修改原始风险结果，便于保持程序初筛可复现。")
    current_status = status_map.get(chosen, "待核查")
    left, right = st.columns([1.2, 2])
    status = left.radio("复核状态", REVIEW_STATES, index=REVIEW_STATES.index(current_status), key=f"rv_status_{chosen}")
    reason = right.text_area(
        "复核原因（非“待核查”必填）",
        value="",
        height=100,
        placeholder="例如：品牌名误命中、价格差异来自多规格SKU、需补充历史价格或资质",
        key=f"rv_reason_{chosen}",
    )
    reviewer = st.text_input("复核人", value=st.session_state.get("reviewer", ""), key="rv_reviewer")
    st.session_state["reviewer"] = reviewer
    if st.button("保存复核结果", type="primary", key=f"rv_save_{chosen}"):
        try:
            ReviewStore().upsert(chosen, status, reason, reviewer)
        except ValueError as exc:
            st.error(str(exc))
        else:
            bundle_review_map_cached.clear()
            st.session_state["flash"] = f"已保存复核结果：{chosen} → {status}"
            st.rerun()

    existing = ReviewStore().load()
    history = existing[existing["商品唯一键"] == chosen]
    if not history.empty:
        st.caption(
            f"最近一次复核：{history.iloc[-1]['复核状态']}　原因：{history.iloc[-1]['复核原因'] or '—'}　"
            f"复核人：{history.iloc[-1]['复核人']}　时间：{history.iloc[-1]['更新时间']}"
        )


def page_review(bundle: DataBundle) -> None:
    header("复核工作台", "查看程序初筛与人工复核的分布，导出复核记录；复核结论不回写原始风险结果。")
    store = ReviewStore()
    records = store.load()
    counts = store.summary()
    total = len(bundle.products)

    m = st.columns(5)
    m[0].metric("已记录复核", f"{len(records):,}")
    for i, state in enumerate(REVIEW_STATES[1:], start=1):
        m[i].metric(state, counts[state])
    st.caption(
        f"商品总数 {total:,}；尚未人工登记的 {total - counts['待核查']:,} 条默认显示为“待核查”。"
    )

    if records.empty:
        st.info("暂无人工复核记录。请先在“商品详情”页对具体商品登记复核结果。")
        return

    joined = records.merge(
        bundle.products[["商品唯一键", "平台", "商品标题", "风险等级", "风险总分"]],
        on="商品唯一键", how="left",
    )
    choice = st.multiselect("按复核状态筛选", REVIEW_STATES, default=REVIEW_STATES, key="rv_filter")
    shown = joined[joined["复核状态"].isin(choice)]
    st.dataframe(
        shown.rename(columns={"风险等级": "程序风险等级", "风险总分": "程序风险总分"}),
        hide_index=True,
        width="stretch",
        column_config={"平台": st.column_config.TextColumn(width="small")},
    )
    st.download_button(
        "导出复核记录",
        data=to_csv_bytes(shown),
        file_name=f"人工复核记录_{datetime.now():%Y%m%d_%H%M}.csv",
        mime="text/csv",
        key="dl_review",
    )
    note("复核状态含义：已排除（人工判断为误报或正常情形）、需补充材料（需历史价格、资质或订单日志）、"
         "确认问题线索（需进一步处理，但仍非行政处罚结论）。")


def page_rules(bundle: DataBundle) -> None:
    header("规则中心", "公开15项规则的阈值、基础分与命中规模，保证评分可解释、可复核。")
    rules = bundle.rules.copy()
    hit = bundle.signals.groupby("风险代码")["商品唯一键"].nunique()
    rules["命中商品数"] = rules["风险代码"].map(hit).fillna(0).astype(int)
    rules["占全部商品比例(%)"] = (rules["命中商品数"] / len(bundle.products) * 100).round(2)

    categories = st.multiselect("风险类别", sorted(rules["风险类别"].unique()), default=sorted(rules["风险类别"].unique()), key="rule_cat")
    shown = rules[rules["风险类别"].isin(categories)]
    st.dataframe(
        shown[["风险代码", "风险类别", "风险信号", "基础分", "命中商品数", "占全部商品比例(%)", "判定规则/阈值"]],
        hide_index=True,
        width="stretch",
        height=440,
    )
    st.subheader("建议核查动作")
    for _, row in shown.iterrows():
        st.caption(f"{row['风险代码']} {row['风险信号']}：{row['建议核查动作']}")
    note("风险总分为命中规则基础分累加（上限100）：高风险≥45，中风险20—44，低风险1—19。"
         "基础分体现核查优先级，不代表违法概率。T02、L01 因当前数据无有效好评率、直播字段完整而未触发。")


def page_quality(bundle: DataBundle) -> None:
    header("数据质量", "说明字段有效率、特殊缺失编码和使用限制，避免把数据缺失误读为业务异常。")
    report = validate_contract(bundle)
    if report.passed:
        st.success(f"数据契约校验通过：{len(report.checks)} 项检查全部符合约定。")
    else:
        st.error(f"数据契约校验未通过：{len(report.failed)} 项异常，请先修复数据再使用页面。")
        st.dataframe(pd.DataFrame(report.failed), hide_index=True, width="stretch")

    with st.expander("查看契约校验明细"):
        st.dataframe(pd.DataFrame(report.checks), hide_index=True, width="stretch")

    st.subheader("字段有效率")
    quality = bundle.field_quality.copy()
    st.dataframe(
        quality,
        hide_index=True,
        width="stretch",
        height=380,
        column_config={"有效率(%)": st.column_config.ProgressColumn("有效率（%）", min_value=0, max_value=100, format="%.2f")},
    )
    st.subheader("关键字段使用限制")
    limits = [
        ("挂牌价、好评率、商品收藏数、商品评论数中的 -1", "视为缺失，不参与相应规则；评论数 0 保留为真实零评论。"),
        ("优惠信息", "整列为空，因此未设计满减、优惠券类规则。"),
        ("好评率", "全部为 -1，当前无有效值，故低好评率规则未触发。"),
        ("直播字段", "仅直播商品有场次与主播数，非直播商品为空属正常，不判为缺失风险。"),
        ("商品销售额", "与月销量×到手价口径可能不同，不直接用于异常判定。"),
        ("商品ID", "按文本处理与展示，避免19位ID被数值精度截断。"),
    ]
    st.dataframe(
        pd.DataFrame(limits, columns=["字段/情形", "处理方式与原因"]),
        hide_index=True, width="stretch",
    )


def page_about(bundle: DataBundle) -> None:
    header("关于本产品", "电商市场监管风险筛查与核查辅助 Demo：数据来源、方法边界与复现方式。")
    st.markdown(
        """
**产品定位**

帮助监管人员从淘宝、天猫、京东、抖音四平台共 4000 条商品中，快速发现优先核查对象，
并查看可解释的风险证据与建议核查动作。

**数据与方法**

- 数据：4 平台 × 1000 条商品 × 53 个原始字段，经清洗后形成统一商品数据。
- 方法：15 项可解释规则，覆盖价格、宣传、主体、销量评价、跨境、直播六类风险。
- 评分：规则基础分累加（上限 100），高风险 ≥45、中风险 20—44、低风险 1—19。
- 关联：商品唯一键连接商品汇总与风险信号明细，明细保留逐条证据。

**结论边界**

- 风险信号是待核查线索，不是违法违规认定。
- 风险分数表示核查优先级，不代表违法概率或处罚程度。
- 未命中仅表示当前规则未发现异常，不等于商品合规。
- 人工复核记录独立保存，不覆盖程序初筛结果。

**复现方式**

```
.venv\\Scripts\\python.exe -m streamlit run demo_app.py
```

原始风险成果由 `ecommerce_risk_detection.py` 生成，数据契约由 `demo_product_spec.py` 校验，
本 Demo 只读取成果、不重算规则。
        """
    )
    st.subheader("当前数据快照")
    stats = compute_overview(bundle.products)
    st.dataframe(pd.DataFrame(list(stats.items()), columns=["指标", "数值"]), hide_index=True, width="stretch")


# ============================================================
# 五、应用入口
# ============================================================
@st.cache_data(show_spinner="正在加载风险成果……")
def bundle_review_map_cached() -> dict:
    return ReviewStore().status_map()


def bundle_review_map() -> dict:
    return bundle_review_map_cached()


def main() -> None:
    st.set_page_config(
        page_title="电商市场监管风险筛查 Demo",
        page_icon=None,
        layout="wide",
        initial_sidebar_state="expanded",
    )
    inject_style()

    try:
        bundle = load_bundle()
    except FileNotFoundError as exc:
        st.error(str(exc))
        st.stop()
    report = validate_contract(bundle)

    st.sidebar.title("电商监管风险筛查")
    st.sidebar.caption("课题二 · 监管风险识别 Demo")
    st.session_state.setdefault("page", "风险总览")

    def _sync_from_sidebar() -> None:
        st.session_state["page"] = st.session_state["nav_widget"]

    # 在控件创建之前同步，保证程序跳转后侧边栏选中项与页面一致
    st.session_state["nav_widget"] = st.session_state["page"]
    st.sidebar.radio("功能导航", PAGES, key="nav_widget", on_change=_sync_from_sidebar)
    nav = st.session_state["page"]
    st.sidebar.divider()
    st.sidebar.metric("商品总数", f"{len(bundle.products):,}")
    st.sidebar.metric("命中风险商品", f"{int(bundle.products['命中信号数'].gt(0).sum()):,}")
    st.sidebar.metric("高风险商品", int(bundle.products["风险等级"].eq("高").sum()))
    if not report.passed:
        st.sidebar.error(f"数据契约异常 {len(report.failed)} 项")
    st.sidebar.divider()
    st.sidebar.caption(
        "风险分数代表核查优先级，不构成违法认定。\n\n"
        "数据：淘宝/天猫/京东/抖音 各1000条\n\n"
        "版本 1.0 · 2026-11"
    )

    if nav == "风险总览":
        page_overview(bundle)
    elif nav == "风险商品":
        page_products(bundle)
    elif nav == "选中商品":
        page_selected(bundle)
    elif nav == "商品详情":
        page_detail(bundle)
    elif nav == "复核工作台":
        page_review(bundle)
    elif nav == "规则中心":
        page_rules(bundle)
    elif nav == "数据质量":
        page_quality(bundle)
    else:
        page_about(bundle)


if __name__ == "__main__":
    main()
