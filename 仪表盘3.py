# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import requests
import json
import os

st.set_page_config(layout="wide", page_title="跨境BP驾驶舱", page_icon="📊")

# ==================== CSS ====================
st.markdown(
    """
<style>
    .main .block-container { padding-top: 0.5rem !important; padding-bottom: 0rem !important; max-width: 100% !important; }
    .kpi-card { background: white; border-radius: 12px; padding: 12px 16px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); border-left: 4px solid #2ecc71; margin-bottom: 4px; }
    .kpi-card .label { font-size: 12px; color: #6b7a8f; font-weight: 500; }
    .kpi-card .value { font-size: 26px; font-weight: 700; color: #1a2332; }
    .kpi-card .sub { font-size: 11px; color: #6b7a8f; }
    .header-bar { background: #0b1a33; padding: 12px 24px; border-radius: 12px; margin-bottom: 12px; display: flex; justify-content: space-between; align-items: center; }
    .header-bar h1 { color: white; font-size: 20px; font-weight: 600; margin: 0; }
    .header-bar .badge { color: #94a3b8; font-size: 13px; }
    .element-container { margin-bottom: 4px !important; }
    div[data-testid="stVerticalBlock"] { gap: 0.2rem !important; }
    div[data-testid="stHorizontalBlock"] { gap: 0.5rem !important; }
    .stDataFrame { font-size: 12px !important; }
    .stTabs [data-baseweb="tab-list"] { gap: 0px !important; }
    .stTabs [data-baseweb="tab"] { padding: 8px 20px !important; }
    hr { margin: 6px 0 !important; }
</style>
""",
    unsafe_allow_html=True,
)

DEEPSEEK_API_KEY = "sk-你的真实key"
DEEPSEEK_API_URL = "https://api.deepseek.com/v1/chat/completions"

EXCHANGE_RATES = {
    "USD": 1.0,
    "EUR": 1.1,
    "GBP": 1.3,
    "JPY": 0.007,
    "CAD": 0.73,
    "AUD": 0.65,
}

st.markdown(
    """
<div class="header-bar">
    <h1>📊 跨境BP驾驶舱</h1>
    <span class="badge">📅 2026-08-05 更新</span>
</div>
""",
    unsafe_allow_html=True,
)

# ==================== 读取数据 ====================
FILE_PATH = r"C:\Users\Administrator\Desktop\周报分析\SC订单.xlsx"

if not os.path.exists(FILE_PATH):
    st.error(f"文件不存在：{FILE_PATH}")
    st.stop()

try:
    df = pd.read_excel(FILE_PATH)
except Exception as e:
    st.error(f"读取失败：{e}")
    st.stop()

# ==================== 数据清洗 ====================
df.columns = df.columns.str.strip()

date_cols = [
    "订购日期",
    "结算时间",
    "发货时限（最早）",
    "发货时限（最迟）",
    "更新时间",
    "发货时间",
    "收货预计",
]
for col in date_cols:
    if col in df.columns:
        df[col] = pd.to_datetime(df[col], errors="coerce")

numeric_cols = [
    "数量",
    "单价",
    "销售收益",
    "税费",
    "采购成本",
    "头程费用",
    "FBA费",
    "促销费-运费折扣",
    "促销费-商品折扣",
    "平台费",
    "其他",
    "积分成本",
    "站外推广费-本金",
    "站外推广费-佣金",
    "退款总金额",
    "订单总金额",
    "销售额(Item Price)",
]
for col in numeric_cols:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

# ==================== 汇率转换 ====================
if "订单币种" in df.columns:
    df["汇率"] = df["订单币种"].map(EXCHANGE_RATES).fillna(1)
else:
    df["汇率"] = 1

df["销售收益_USD"] = df["销售收益"] / df["汇率"]
df["采购成本_USD"] = df["采购成本"] / df["汇率"]
df["头程费用_USD"] = df["头程费用"] / df["汇率"]
df["FBA费_USD"] = df["FBA费"] / df["汇率"]
df["平台费_USD"] = df["平台费"] / df["汇率"]
df["促销费_USD"] = df["促销费-商品折扣"] / df["汇率"]
df["退款总金额_USD"] = df["退款总金额"] / df["汇率"]
df["站外推广费_USD"] = (df["站外推广费-本金"] + df["站外推广费-佣金"]) / df["汇率"]
df["税费_USD"] = df["税费"] / df["汇率"]
df["其他_USD"] = df["其他"] / df["汇率"]
df["积分成本_USD"] = df["积分成本"] / df["汇率"]

# ==================== 标准财务计算 ====================
# 实际营收 = 销售收益 - 促销 - 退款
df["实际营收_USD"] = df["销售收益_USD"] - df["促销费_USD"] - df["退款总金额_USD"]

# 总成本（全部加总）
df["总成本_USD"] = (
    df["采购成本_USD"]
    + df["头程费用_USD"]
    + df["FBA费_USD"]
    + df["平台费_USD"]
    + df["税费_USD"]
    + df["促销费_USD"]
    + df["站外推广费_USD"]
    + df["其他_USD"]
    + df["积分成本_USD"]
)

# 净利润 = 实际营收 - 总成本
df["净利润"] = df["实际营收_USD"] + df["总成本_USD"]

# 利润率
df["净利率"] = df.apply(
    lambda x: x["净利润"] / x["实际营收_USD"] * 100 if x["实际营收_USD"] > 0 else 0,
    axis=1,
)

# ==================== 店铺汇总 ====================
if "店铺" not in df.columns:
    st.error("缺少'店铺'字段")
    st.stop()

df_shop = (
    df.groupby("店铺")
    .agg(
        {
            "实际营收_USD": "sum",
            "净利润": "sum",
            "销售收益_USD": "sum",
            "采购成本_USD": "sum",
            "头程费用_USD": "sum",
            "FBA费_USD": "sum",
            "平台费_USD": "sum",
            "税费_USD": "sum",
            "促销费_USD": "sum",
            "站外推广费_USD": "sum",
            "其他_USD": "sum",
            "积分成本_USD": "sum",
            "退款总金额_USD": "sum",
            "订单号": "count",
            "数量": "sum",
        }
    )
    .reset_index()
    .rename(columns={"订单号": "订单数", "数量": "销量"})
)

# 店铺级利润率
df_shop["净利率"] = df_shop.apply(
    lambda x: x["净利润"] / x["实际营收_USD"] * 100 if x["实际营收_USD"] > 0 else 0,
    axis=1,
)

# 成本占比
df_shop["平台费占比"] = df_shop.apply(
    lambda x: x["平台费_USD"] / x["实际营收_USD"] * 100 if x["实际营收_USD"] > 0 else 0,
    axis=1,
)
df_shop["FBA费占比"] = df_shop.apply(
    lambda x: x["FBA费_USD"] / x["实际营收_USD"] * 100 if x["实际营收_USD"] > 0 else 0,
    axis=1,
)
df_shop["头程费占比"] = df_shop.apply(
    lambda x: (
        x["头程费用_USD"] / x["实际营收_USD"] * 100 if x["实际营收_USD"] > 0 else 0
    ),
    axis=1,
)
df_shop["促销费占比"] = df_shop.apply(
    lambda x: x["促销费_USD"] / x["实际营收_USD"] * 100 if x["实际营收_USD"] > 0 else 0,
    axis=1,
)
df_shop["站外推广占比"] = df_shop.apply(
    lambda x: (
        x["站外推广费_USD"] / x["实际营收_USD"] * 100 if x["实际营收_USD"] > 0 else 0
    ),
    axis=1,
)
df_shop["退款损失占比"] = df_shop.apply(
    lambda x: (
        x["退款总金额_USD"] / x["实际营收_USD"] * 100 if x["实际营收_USD"] > 0 else 0
    ),
    axis=1,
)
df_shop["税费占比"] = df_shop.apply(
    lambda x: x["税费_USD"] / x["实际营收_USD"] * 100 if x["实际营收_USD"] > 0 else 0,
    axis=1,
)

df_shop["盈利分类"] = pd.cut(
    df_shop["净利润"],
    bins=[-float("inf"), 0, 1000, float("inf")],
    labels=["亏损", "微利", "盈利"],
)

# 28原则分层
df_shop["gmv_rank"] = df_shop["实际营收_USD"].rank(ascending=False, method="dense")
total_shops = len(df_shop)
df_shop["tier"] = "瘦狗"
df_shop.loc[df_shop["gmv_rank"] <= total_shops * 0.2, "tier"] = "明星"
df_shop.loc[
    (df_shop["gmv_rank"] > total_shops * 0.2)
    & (df_shop["gmv_rank"] <= total_shops * 0.4),
    "tier",
] = "现金牛"
df_shop.loc[
    (df_shop["gmv_rank"] > total_shops * 0.4)
    & (df_shop["gmv_rank"] <= total_shops * 0.6),
    "tier",
] = "问题"

# ==================== KPI（去掉盈利店铺，换成平均净利率） ====================
total_actual = df_shop["实际营收_USD"].sum()
total_profit = df_shop["净利润"].sum()
avg_margin = df_shop["净利率"].mean()
total_orders = len(df)
total_qty = df["数量"].sum()

c1, c2, c3, c4, c5 = st.columns(5)
c1.markdown(
    f"""<div class="kpi-card"><div class="label">实际营收</div><div class="value">${total_actual / 1000:.1f}k</div></div>""",
    unsafe_allow_html=True,
)
color = "#2ecc71" if total_profit > 0 else "#e74c3c"
c2.markdown(
    f"""<div class="kpi-card" style="border-left-color:{color};"><div class="label">净利润</div><div class="value" style="color:{color};">${total_profit / 1000:.1f}k</div><div class="sub">净利率 {avg_margin:.1f}%</div></div>""",
    unsafe_allow_html=True,
)
c3.markdown(
    f"""<div class="kpi-card" style="border-left-color:#3498db;"><div class="label">平均净利率</div><div class="value">{avg_margin:.1f}%</div><div class="sub">店铺数 {total_shops}</div></div>""",
    unsafe_allow_html=True,
)
c4.markdown(
    f"""<div class="kpi-card" style="border-left-color:#f39c12;"><div class="label">总销量</div><div class="value">{total_qty:,.0f}</div><div class="sub">订单数 {total_orders}</div></div>""",
    unsafe_allow_html=True,
)
risk = df[df["净利润"] < 0].shape[0]
risk_pct = risk / total_orders * 100 if total_orders > 0 else 0
c5.markdown(
    f"""<div class="kpi-card" style="border-left-color:#e74c3c;"><div class="label">亏损订单</div><div class="value">{risk}</div><div class="sub">占比 {risk_pct:.1f}%</div></div>""",
    unsafe_allow_html=True,
)

st.divider()

# ==================== Tabs ====================
tab1, tab2, tab3, tab4 = st.tabs([" 收入 & 利润", " 成本分析", " 库存分析", " AI 顾问"])

with tab1:
    col_left, col_right = st.columns([2, 1])
    with col_left:
        st.markdown("**📌 店铺分层气泡图**")
        fig = px.scatter(
            df_shop,
            x="实际营收_USD",
            y="净利润",
            text="店铺",
            color="tier",
            size="实际营收_USD",
            size_max=50,
            color_discrete_map={
                "明星": "#2ecc71",
                "现金牛": "#3498db",
                "问题": "#f39c12",
                "瘦狗": "#95a5a6",
            },
            labels={"实际营收_USD": "实际营收 (USD)", "净利润": "净利润 (USD)"},
        )
        fig.add_hline(y=0, line_dash="dash", line_color="red")
        fig.update_traces(textposition="top center", textfont_size=10)
        fig.update_layout(
            plot_bgcolor="white", height=380, margin=dict(l=10, r=10, t=10, b=10)
        )
        st.plotly_chart(fig, use_container_width=True)

    with col_right:
        st.markdown("**分层汇总**")
        tier_sum = (
            df_shop.groupby("tier")
            .agg(
                店铺数=("店铺", "count"),
                实际营收=("实际营收_USD", "sum"),
                净利润=("净利润", "sum"),
            )
            .reset_index()
        )
        tier_sum["净利率"] = tier_sum.apply(
            lambda x: x["净利润"] / x["实际营收"] * 100 if x["实际营收"] > 0 else 0,
            axis=1,
        )
        st.dataframe(
            tier_sum.style.format(
                {"实际营收": "${:,.0f}", "净利润": "${:,.0f}", "净利率": "{:.1f}%"}
            ),
            use_container_width=True,
            height=220,
        )

    st.divider()

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**店铺营收 & 净利润**")
        # 按实际营收排序，取前15
        df_top15 = df_shop.nlargest(15, "实际营收_USD").sort_values(
            "实际营收_USD", ascending=True
        )

        fig_bar = go.Figure()
        # 营收用浅蓝色柱
        fig_bar.add_trace(
            go.Bar(
                x=df_top15["实际营收_USD"],
                y=df_top15["店铺"],
                name="实际营收",
                marker_color="#3498db",
                orientation="h",
                text=df_top15["实际营收_USD"].apply(lambda x: f"${x:,.0f}"),
                textposition="outside",
                textfont=dict(size=9),
            )
        )
        # 净利润用绿色/红色柱叠加
        colors_bar = ["#2ecc71" if p > 0 else "#e74c3c" for p in df_top15["净利润"]]
        fig_bar.add_trace(
            go.Bar(
                x=df_top15["净利润"],
                y=df_top15["店铺"],
                name="净利润",
                marker_color=colors_bar,
                orientation="h",
                text=df_top15["净利润"].apply(lambda x: f"${x:,.0f}"),
                textposition="outside",
                textfont=dict(size=9),
            )
        )
        fig_bar.update_layout(
            barmode="group",
            plot_bgcolor="white",
            height=420,
            margin=dict(l=10, r=80, t=10, b=10),
            legend=dict(orientation="h", y=1.05, font_size=11),
            xaxis_title="USD",
            yaxis=dict(autorange="reversed"),
        )
        st.plotly_chart(fig_bar, use_container_width=True)

    with c2:
        st.markdown("**🏆 Top5 盈利 / 📉 Bottom5 亏损**")
        tt, tb = st.tabs(["盈利", "亏损"])
        with tt:
            st.dataframe(
                df_shop.nlargest(5, "净利润")[
                    ["店铺", "净利润", "实际营收_USD", "净利率"]
                ].style.format(
                    {
                        "净利润": "${:,.0f}",
                        "实际营收_USD": "${:,.0f}",
                        "净利率": "{:.1f}%",
                    }
                ),
                use_container_width=True,
                height=200,
            )
        with tb:
            st.dataframe(
                df_shop.nsmallest(5, "净利润")[
                    ["店铺", "净利润", "实际营收_USD", "净利率"]
                ].style.format(
                    {
                        "净利润": "${:,.0f}",
                        "实际营收_USD": "${:,.0f}",
                        "净利率": "{:.1f}%",
                    }
                ),
                use_container_width=True,
                height=200,
            )

    st.markdown("**🔴 亏损订单明细**")
    df_loss_orders = df[df["净利润"] < 0].copy()

    if df_loss_orders.empty:
        st.success("🎉 本月无亏损订单")
    else:
        st.warning(f"发现 {len(df_loss_orders)} 笔亏损订单")

        # ==================== 合并标签 ====================
        def get_loss_reason(row):
            reasons = []

            if row.get("是否促销") in ["是", True, "TRUE", "True", 1]:
                promo_ratio = (
                    row["促销费_USD"] / row["销售收益_USD"] * 100
                    if row["销售收益_USD"] > 0
                    else 0
                )
                if promo_ratio > 50:
                    reasons.append("大额促销")
                elif promo_ratio > 30:
                    reasons.append("促销")
                else:
                    reasons.append("小额促销")

            if row.get("是否退款") in ["是", True, "TRUE", "True", 1]:
                reasons.append("退款")

            if row.get("是否退货") in ["是", True, "TRUE", "True", 1]:
                reasons.append("退货")

            if row.get("换货订单") in ["是", True, "TRUE", "True", 1]:
                reasons.append("换货")

            if row.get("是否B2B") in ["是", True, "TRUE", "True", 1]:
                reasons.append("B2B")

            if row.get("是否会员") in ["是", True, "TRUE", "True", 1]:
                reasons.append("会员")

            if row.get("是否优质配送") in ["是", True, "TRUE", "True", 1]:
                reasons.append("优质配送")

            if not reasons:
                reasons.append("其他")

            return " · ".join(reasons)

        df_loss_orders["亏损原因"] = df_loss_orders.apply(get_loss_reason, axis=1)

        # ==================== CSS美化 ====================
        st.markdown(
            """
        <style>
        .loss-container {
            background: white;
            border-radius: 16px;
            padding: 20px 20px 10px 20px;
            box-shadow: 0 2px 12px rgba(0,0,0,0.06);
            border: 1px solid #f0f2f5;
            margin-bottom: 12px;
        }
        .loss-container .section-title {
            font-size: 14px;
            font-weight: 600;
            color: #1a2332;
            margin-bottom: 12px;
        }
        .loss-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
        }
        .loss-table th {
            text-align: left;
            padding: 10px 12px;
            background: #f8fafc;
            color: #475569;
            font-weight: 600;
            font-size: 11px;
            text-transform: uppercase;
            letter-spacing: 0.3px;
            border-bottom: 2px solid #e9edf2;
        }
        .loss-table td {
            padding: 9px 12px;
            border-bottom: 1px solid #f1f4f8;
            color: #1e293b;
        }
        .loss-table tr:hover td {
            background: #fafbfc;
        }
        .loss-table .rank-num {
            display: inline-block;
            width: 22px;
            height: 22px;
            line-height: 22px;
            text-align: center;
            border-radius: 50%;
            font-size: 11px;
            font-weight: 700;
            background: #f1f4f8;
            color: #64748b;
        }
        .loss-table .rank-1 { background: #fef3c7; color: #92400e; }
        .loss-table .rank-2 { background: #e8f0fe; color: #1a4d8a; }
        .loss-table .rank-3 { background: #fce4ec; color: #c62828; }
        .loss-table .badge-tag {
            display: inline-block;
            padding: 2px 10px;
            border-radius: 20px;
            font-size: 11px;
            font-weight: 500;
            margin: 1px 2px;
            background: #f1f4f8;
            color: #475569;
        }
        .loss-table .badge-tag.tag-promo { background: #fef3c7; color: #92400e; }
        .loss-table .badge-tag.tag-refund { background: #fce4ec; color: #c62828; }
        .loss-table .badge-tag.tag-return { background: #ffccbc; color: #bf360c; }
        .loss-table .badge-tag.tag-exchange { background: #e3f2fd; color: #0d47a1; }
        .loss-table .badge-tag.tag-other { background: #f1f4f8; color: #64748b; }
        .loss-table .badge-tag.tag-b2b { background: #e8f5e9; color: #1b5e20; }
        .loss-table .badge-tag.tag-vip { background: #f3e5f5; color: #6a1b9a; }
        .loss-table .badge-tag.tag-shipping { background: #e0f7fa; color: #00695c; }
        </style>
        """,
            unsafe_allow_html=True,
        )

        # ==================== 拆分标签统计 ====================
        all_tags = []
        for tags in df_loss_orders["亏损原因"]:
            for tag in tags.split(" · "):
                all_tags.append(tag)
        tag_counts = pd.Series(all_tags).value_counts().reset_index()
        tag_counts.columns = ["原因", "订单数"]

        # 给原因打颜色标签
        def tag_color(name):
            if "促销" in name:
                return "tag-promo"
            elif "退款" in name:
                return "tag-refund"
            elif "退货" in name:
                return "tag-return"
            elif "换货" in name:
                return "tag-exchange"
            elif "B2B" in name:
                return "tag-b2b"
            elif "会员" in name:
                return "tag-vip"
            elif "配送" in name:
                return "tag-shipping"
            else:
                return "tag-other"

        col_left, col_right = st.columns([1, 1])

        with col_left:
            st.markdown('<div class="loss-container">', unsafe_allow_html=True)
        display_cols = [
            "店铺",
            "订单号",
            "亏损原因",
            "实际营收_USD",
            "净利润",
            "净利率",
            "销售收益_USD",
        ]
        extra_cols = [
            "促销费_USD",
            "退款总金额_USD",
            "采购成本_USD",
            "头程费用_USD",
            "FBA费_USD",
        ]
        for col in extra_cols:
            if col in df_loss_orders.columns:
                display_cols.append(col)

        display_cols = [col for col in display_cols if col in df_loss_orders.columns]
        df_display = df_loss_orders[display_cols].sort_values("净利润", ascending=True)

        fmt_dict = {}
        for col in display_cols:
            if "USD" in col or "营收" in col or "利润" in col or "收益" in col:
                fmt_dict[col] = "${:,.0f}"
        fmt_dict["净利率"] = "{:.1f}%"

        st.caption(
            f"📊 共 {len(df_loss_orders)} 笔亏损订单 ｜ 总亏损 ${df_loss_orders['净利润'].sum():,.0f} ｜ 平均 ${df_loss_orders['净利润'].mean():,.0f} ｜ {df_loss_orders['店铺'].nunique()} 家店铺"
        )

        csv = df_loss_orders[display_cols].to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            label="📥 导出CSV",
            data=csv,
            file_name=f"亏损订单_{pd.Timestamp.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            use_container_width=False,
        )
        st.dataframe(
            df_display.style.format(fmt_dict).background_gradient(
                subset=["净利润"], cmap="Reds_r"
            ),
            use_container_width=True,
            height=350,
        )


with tab2:
    # ==================== 费用日历热力图 ====================
    st.markdown("**📅 店铺费用日历（日维度）**")
    
    shop_list = ['全部'] + sorted(df['店铺'].unique().tolist())
    selected_shop = st.selectbox("选择店铺", shop_list, key="calendar_shop")
    
    if selected_shop == '全部':
        df_calendar = df.copy()
    else:
        df_calendar = df[df['店铺'] == selected_shop].copy()
    
    if '订购日期' not in df_calendar.columns:
        st.warning("数据中缺少'订购日期'字段，无法展示日历")
    elif df_calendar.empty:
        st.info("该店铺暂无数据")
    else:
        df_calendar['日期'] = pd.to_datetime(df_calendar['订购日期'], errors='coerce').dt.date
        df_calendar = df_calendar[df_calendar['日期'].notna()].copy()
        
        if df_calendar.empty:
            st.info("该时段无有效日期数据")
        else:
            # 按天汇总
            daily = df_calendar.groupby('日期').agg({
                '实际营收_USD': 'sum',
                '采购成本_USD': 'sum',
                '头程费用_USD': 'sum',
                'FBA费_USD': 'sum',
                '平台费_USD': 'sum',
                '税费_USD': 'sum',
                '促销费_USD': 'sum',
                '站外推广费_USD': 'sum',
                '其他_USD': 'sum',
                '积分成本_USD': 'sum'
            }).reset_index()
            
            # 计算总费用 = 所有费用加总（费用本身是负数）
            daily['总费用'] = daily[['采购成本_USD', '头程费用_USD', 'FBA费_USD', '平台费_USD', 
                                    '税费_USD', '促销费_USD', '站外推广费_USD', '其他_USD', '积分成本_USD']].sum(axis=1)
            
            # 费率 = 总费用 * (-1) / 实际营收 * 100，正数表示费用占比
            daily['费率'] = daily.apply(lambda x: x['总费用'] * (-1) / x['实际营收_USD'] * 100 if x['实际营收_USD'] != 0 else 0, axis=1)
            
            if daily.empty:
                st.info("无有效数据")
            else:
                # 获取月份
                daily['月份'] = pd.to_datetime(daily['日期']).dt.month
                daily['日'] = pd.to_datetime(daily['日期']).dt.day
                daily['年份'] = pd.to_datetime(daily['日期']).dt.year
                
                # 统计信息
                avg_rate = daily['费率'].mean()
                st.caption(f"📊 共 {len(daily)} 天数据 ｜ 平均费率 {avg_rate:.1f}%")
                
                # ========== 日历和饼图两列布局 ==========
                col_cal, col_pie = st.columns([3, 1])
                
                with col_cal:
                    import plotly.express as px
                    
                    for month in sorted(daily['月份'].unique()):
                        month_data = daily[daily['月份'] == month]
                        year = month_data['年份'].iloc[0]
                        
                        st.markdown(f"**{year}年{month}月**")
                        
                        days_in_month = pd.Period(f"{year}-{month:02d}").days_in_month
                        first_day = pd.Timestamp(f"{year}-{month:02d}-01").dayofweek
                        
                        matrix = []
                        week = []
                        
                        for _ in range(first_day):
                            week.append(None)
                        
                        for d in range(1, days_in_month + 1):
                            row = month_data[month_data['日'] == d]
                            if not row.empty:
                                rate = row['费率'].iloc[0]
                                week.append(rate)
                            else:
                                week.append(None)
                            
                            if len(week) == 7:
                                matrix.append(week)
                                week = []
                        
                        if week:
                            while len(week) < 7:
                                week.append(None)
                            matrix.append(week)
                        
                        fig = px.imshow(
                            matrix,
                            text_auto='.0f',
                            aspect='equal',
                            color_continuous_scale=['#3498db', '#2ecc71', '#f1c40f', '#e74c3c'],
                            range_color=[0, 100],
                            labels=dict(x="", y="", color="费率%")
                        )
                        fig.update_xaxes(
                            ticktext=['日', '一', '二', '三', '四', '五', '六'],
                            tickvals=[0, 1, 2, 3, 4, 5, 6]
                        )
                        fig.update_yaxes(
                            ticktext=[f"第{i+1}周" for i in range(len(matrix))], 
                            tickvals=list(range(len(matrix)))
                        )
                        fig.update_layout(
                            height=320,
                            margin=dict(l=10, r=10, t=10, b=10),
                            coloraxis_colorbar=dict(title="费率%")
                        )
                        st.plotly_chart(fig, use_container_width=True)
                        
                        month_total = month_data['实际营收_USD'].sum()
                        month_cost = month_data['总费用'].sum() * (-1)
                        month_rate = month_cost / month_total * 100 if month_total != 0 else 0
                        st.caption(f"📊 {year}年{month}月 ｜ 营收 ${month_total:,.0f} ｜ 费用 ${month_cost:,.0f} ｜ 费率 {month_rate:.1f}%")
                        st.divider()
                
                with col_pie:
                    st.markdown("****")
                    
                    # 汇总该店铺的所有费用
                    cost_summary = {
                        '采购成本': daily['采购成本_USD'].sum() * (-1),
                        '头程费用': daily['头程费用_USD'].sum() * (-1),
                        'FBA费': daily['FBA费_USD'].sum() * (-1),
                        '平台费': daily['平台费_USD'].sum() * (-1),
                        '税费': daily['税费_USD'].sum() * (-1),
                        '促销费': daily['促销费_USD'].sum() * (-1),
                        '站外推广': daily['站外推广费_USD'].sum() * (-1),
                        '其他杂费': (daily['其他_USD'].sum() + daily['积分成本_USD'].sum()) * (-1)
                    }
                    
                    # 过滤掉0或负值
                    cost_df = pd.DataFrame({
                        '费用项': list(cost_summary.keys()),
                        '金额': list(cost_summary.values())
                    })
                    cost_df = cost_df[cost_df['金额'] > 0].copy()
                    
                    if not cost_df.empty:
                        fig_pie = px.pie(
                            cost_df,
                            values='金额',
                            names='费用项',
                            title=f'{selected_shop} 费用占比' if selected_shop != '全部' else '全部店铺 费用占比',
                            color_discrete_sequence=px.colors.qualitative.Set2,
                            hole=0.4
                        )
                        fig_pie.update_traces(textposition='inside', textinfo='percent+label')
                        fig_pie.update_layout(
                            height=400,
                            margin=dict(l=10, r=10, t=30, b=10)
                        )
                        st.plotly_chart(fig_pie, use_container_width=True)
                    else:
                        st.info("该店铺暂无费用数据")
                        
                        
with tab3:
    st.markdown("**📦 库存库龄分析**")
    
    path = "C:\\Users\\Administrator\\Desktop\\周报分析\\库存.xlsx"
    df_inv = pd.read_excel(path)
    
    bins = [-1, 90, 120, 180, float('inf')]
    labels = ['0-90天', '91-120天', '121-180天', '>181天']
    df_inv['库龄分段'] = pd.cut(df_inv['库龄天数'], bins=bins, labels=labels)
    
    df_amz = df_inv[df_inv['平台'] == '亚马逊']
    df_multi = df_inv[df_inv['平台'] == '多平台']
    
    # ========== 公共函数 ==========
    def get_alert_data(df, site_filter=None):
        if site_filter:
            df = df[df['站点'] == site_filter]
        alert = df.groupby(['SKU', '名称', '站点']).apply(
            lambda x: pd.Series({
                '总数量': x['数量'].sum(),
                '加权平均库龄': (x['数量'] * x['库龄天数']).sum() / x['数量'].sum() if x['数量'].sum() > 0 else 0
            })
        ).reset_index()
        alert = alert[alert['加权平均库龄'] > 30]
        alert = alert.sort_values('加权平均库龄', ascending=False)
        return alert
    
    tab_all, tab_amz, tab_multi = st.tabs(["全部", "亚马逊", "多平台"])
    
    # ========== 滑动条：超期天数阈值 ==========
    st.markdown("---")
    col_slider, col_info = st.columns([2, 1])
    with col_slider:
        days_threshold = st.slider(
            "📌 滑动调整：超期天数阈值",
            min_value=30,
            max_value=180,
            value=60,
            step=10,
            help="拖动滑块，查看不同超期天数下的SKU数量"
        )
    with col_info:
        st.metric("当前阈值", f">{days_threshold} 天")
    
    st.markdown("---")
    
    # ========== 通用预警函数（带阈值） ==========
    def get_alert_data_with_threshold(df, threshold, site_filter=None):
        if site_filter:
            df = df[df['站点'] == site_filter]
        alert = df.groupby(['SKU', '名称', '站点']).apply(
            lambda x: pd.Series({
                '总数量': x['数量'].sum(),
                '加权平均库龄': (x['数量'] * x['库龄天数']).sum() / x['数量'].sum() if x['数量'].sum() > 0 else 0
            })
        ).reset_index()
        alert = alert[alert['加权平均库龄'] > threshold]
        alert = alert.sort_values('加权平均库龄', ascending=False)
        return alert
    
    # ========== Tab全部 ==========
    with tab_all:
        st.markdown("#### 📊 库存分布")
        pivot = pd.pivot_table(df_inv, values='数量', index='库龄分段', columns='站点', aggfunc='sum', fill_value=0)
        c1, c2 = st.columns([1, 2])
        with c1:
            st.dataframe(pivot, use_container_width=True, height=280)
        with c2:
            df_plot = pivot.reset_index().melt(id_vars='库龄分段', var_name='站点', value_name='数量')
            fig = px.bar(df_plot, x='站点', y='数量', color='库龄分段', barmode='group', title='全部库存库龄分布')
            fig.update_layout(height=280, plot_bgcolor='white', margin=dict(l=10, r=10, t=30, b=10))
            st.plotly_chart(fig, use_container_width=True)
        
        st.markdown("---")
        st.markdown(f"#### ⚠️ 超 {days_threshold} 天库存预警")
        
        alert_all = get_alert_data_with_threshold(df_inv, days_threshold)
        if len(alert_all) > 0:
            st.warning(f"发现 {len(alert_all)} 个SKU加权平均库龄超过 {days_threshold} 天")
            st.dataframe(alert_all, use_container_width=True, height=300)
            
            st.markdown("##### 按站点查看")
            sites = alert_all['站点'].unique().tolist()
            cols = st.columns(min(len(sites), 4))
            for idx, site in enumerate(sites):
                with cols[idx % 4]:
                    site_data = alert_all[alert_all['站点'] == site]
                    st.metric(f"📌 {site}", f"{len(site_data)} 个SKU")
                    st.dataframe(site_data[['SKU', '名称', '总数量', '加权平均库龄']].head(5), use_container_width=True)
        else:
            st.success(f"✅ 所有SKU加权平均库龄均在 {days_threshold} 天以内")
    
    # ========== Tab亚马逊 ==========
    with tab_amz:
        pivot = pd.pivot_table(df_amz, values='数量', index='库龄分段', columns='站点', aggfunc='sum', fill_value=0)
        c1, c2 = st.columns([1, 2])
        with c1:
            st.dataframe(pivot, use_container_width=True, height=280)
        with c2:
            df_plot = pivot.reset_index().melt(id_vars='库龄分段', var_name='站点', value_name='数量')
            fig = px.bar(df_plot, x='站点', y='数量', color='库龄分段', barmode='group', title='亚马逊库存库龄分布')
            fig.update_layout(height=280, plot_bgcolor='white', margin=dict(l=10, r=10, t=30, b=10))
            st.plotly_chart(fig, use_container_width=True)
        
        st.markdown("---")
        st.markdown(f"#### ⚠️ 亚马逊超 {days_threshold} 天库存预警")
        alert_amz = get_alert_data_with_threshold(df_amz, days_threshold)
        if len(alert_amz) > 0:
            st.warning(f"发现 {len(alert_amz)} 个SKU加权平均库龄超过 {days_threshold} 天")
            st.dataframe(alert_amz, use_container_width=True, height=300)
        else:
            st.success(f"✅ 所有SKU加权平均库龄均在 {days_threshold} 天以内")
    
    # ========== Tab多平台 ==========
    with tab_multi:
        pivot = pd.pivot_table(df_multi, values='数量', index='库龄分段', columns='站点', aggfunc='sum', fill_value=0)
        c1, c2 = st.columns([1, 2])
        with c1:
            st.dataframe(pivot, use_container_width=True, height=280)
        with c2:
            df_plot = pivot.reset_index().melt(id_vars='库龄分段', var_name='站点', value_name='数量')
            fig = px.bar(df_plot, x='站点', y='数量', color='库龄分段', barmode='group', title='多平台库存库龄分布')
            fig.update_layout(height=280, plot_bgcolor='white', margin=dict(l=10, r=10, t=30, b=10))
            st.plotly_chart(fig, use_container_width=True)
        
        st.markdown("---")
        st.markdown(f"#### ⚠️ 多平台超 {days_threshold} 天库存预警")
        alert_multi = get_alert_data_with_threshold(df_multi, days_threshold)
        if len(alert_multi) > 0:
            st.warning(f"发现 {len(alert_multi)} 个SKU加权平均库龄超过 {days_threshold} 天")
            st.dataframe(alert_multi, use_container_width=True, height=300)
        else:
            st.success(f"✅ 所有SKU加权平均库龄均在 {days_threshold} 天以内")
# ==================== Tab 4: AI 顾问 ====================
with tab4:
    st.markdown("**🤖 AI 财务顾问**")
    
    # ---- API Key 输入 ----
    col_key1, col_key2 = st.columns([3, 1])
    with col_key1:
        api_key_input = st.text_input(
            "🔑 DeepSeek API Key",
            type="password",
            placeholder="请输入你的 DeepSeek API Key（sk-...）",
            help="输入后将在本次会话中生效，不会保存",
            key="ai_tab_api_key"
        )
    with col_key2:
        if api_key_input:
            st.success("✅ 已配置")
        else:
            st.info("ℹ️ 请输入")
    
    st.markdown("---")
    
    # ---- 预设问题 ----
    st.markdown("**📌 快捷提问**")
    preset_qs = [
        "本月经营情况总结",
        "哪些店铺需要重点关注？",
        "给出降本增效建议",
        "库存风险在哪里？",
        "利润下滑原因分析"
    ]
    cols = st.columns(len(preset_qs))
    for idx, q in enumerate(preset_qs):
        with cols[idx]:
            if st.button(q, key=f"ai_tab_preset_{idx}"):
                st.session_state["ai_tab_question"] = q
    
    st.markdown("---")
    
    # ---- 输入框 ----
    question = st.text_area(
        "💬 输入你的问题",
        value=st.session_state.get("ai_tab_question", ""),
        placeholder="例如：美国站最近两个月利润为什么下降？哪些SKU应该清货？",
        height=80
    )
    
    # ---- 数据摘要（供AI参考） ----
    # 【在这里加上 cost_cols2 的定义】
    cost_cols2 = [
        "平台费占比",
        "FBA费占比",
        "头程费占比",
        "促销费占比",
        "站外推广占比",
        "退款损失占比",
        "税费占比",
    ]
    
    ds = {
        "总实际营收_USD": float(total_actual),
        "总净利润_USD": float(total_profit),
        "平均净利率_%": float(avg_margin),
        "店铺数": int(total_shops),
        "分层分布": df_shop["tier"].value_counts().to_dict(),
        "盈利店铺数": int(len(df_shop[df_shop["净利润"] > 0])),
        "亏损店铺数": int(len(df_shop[df_shop["净利润"] < 0])),
        "总订单数": int(total_orders),
        "总销量": int(total_qty),
        "亏损订单数": int(risk),
    }
    for c in cost_cols2:
        if c in df_shop.columns:
            ds[c] = float(df_shop[c].mean())
    
    # ---- 提问按钮 ----
    col_btn1, col_btn2 = st.columns([1, 5])
    with col_btn1:
        ask_btn = st.button("🚀 提问", type="primary", use_container_width=True)
    
    with col_btn2:
        if st.button("🗑️ 清空", use_container_width=True):
            st.session_state["ai_tab_question"] = ""
            st.rerun()
    
    # ---- 调用 AI ----
    def call_ai(question, api_key, data_summary):
        if not api_key:
            return "⚠️ 请先输入 DeepSeek API Key"
        
        if not question:
            return "⚠️ 请输入问题"
        
        headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        }
        
        system_prompt = """你是跨境电商财务BP分析师。根据提供的业务数据，给出专业的分析建议。
        回答要求：
        1. 简洁、有数据支撑、可执行
        2. 用分点列出建议
        3. 聚焦利润优化、成本控制、风险预警、运营策略
        4. 数据用具体数字说话"""
        
        full_prompt = f"业务数据摘要：{json.dumps(data_summary, ensure_ascii=False, default=str)}\n\n问题：{question}"
        
        payload = {
            "model": "deepseek-chat",
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": full_prompt}
            ],
            "temperature": 0.7,
            "max_tokens": 800
        }
        
        try:
            response = requests.post(
                "https://api.deepseek.com/v1/chat/completions",
                headers=headers,
                json=payload,
                timeout=30
            )
            if response.status_code == 200:
                return response.json()["choices"][0]["message"]["content"]
            else:
                return f"❌ API请求失败：{response.status_code}\n{response.text}"
        except requests.exceptions.Timeout:
            return "⏰ 请求超时，请稍后重试"
        except Exception as e:
            return f"❌ 异常：{str(e)}"
    
    if ask_btn:
        with st.spinner("🤔 AI 正在分析..."):
            answer = call_ai(question, api_key_input, ds)
            st.markdown("---")
            st.markdown(f"**📋 回答：**\n\n{answer}")
    
    # ---- 历史记录 ----
    if "ai_tab_history" not in st.session_state:
        st.session_state["ai_tab_history"] = []
    
    if ask_btn and question and api_key_input:
        st.session_state["ai_tab_history"].append({
            "问题": question,
            "回答": answer,
            "时间": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M")
        })
    
    if st.session_state["ai_tab_history"]:
        with st.expander("📜 历史记录"):
            for h in st.session_state["ai_tab_history"][-5:]:
                st.markdown(f"**{h['时间']}**")
                st.markdown(f"**Q:** {h['问题']}")
                st.markdown(f"**A:** {h['回答'][:200]}...")
                st.markdown("---")
    
    # ---- 使用说明 ----
    with st.expander("📖 使用说明"):
        st.markdown("""
        1. **输入 API Key**：在上方输入你的 DeepSeek API Key
        2. **选择问题**：点击快捷按钮或手动输入问题
        3. **点击提问**：AI 将基于当前业务数据给出分析
        4. **查看历史**：最近5次问答记录会保存
        
        **提示：** API Key 仅本次会话有效，刷新页面需要重新输入
        """)