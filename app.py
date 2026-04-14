import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import io
import warnings
from pathlib import Path

# --- 1. 网页基础配置 ---
warnings.filterwarnings("ignore")
st.set_page_config(page_title="疑似经济困难学生初筛系统", layout="wide")

# ==========================================
# --- 2. 核心算法逻辑 (已适配网页上传模式) ---
# ==========================================

# 将你原来的全局参数设为默认值，并允许在网页侧边栏修改
MIN_ACTIVE_DAYS = 15
MIN_SPEND_COUNT = 20
HIGH_AMOUNT_THRESHOLD = 20
A_MEDIAN_TXN_MAX = 3.0
A_ONE_YUAN_RATIO_MIN = 0.50
A_DOMINANT_POS_RATIO_MIN = 0.70
A_POS_NUNIQUE_MAX = 2
A_HASH_RATIO_MIN = 0.60
A_MODE_RATIO_MIN = 0.70

WEIGHTS = {
    "score_total_spend": 0.35,
    "score_daily_spend": 0.30,
    "score_median_txn": 0.20,
    "score_high_amount_ratio": 0.10,
    "score_avg_deposit": 0.05,
}


def inverse_percentile_score(s):
    s = pd.to_numeric(s, errors="coerce")
    out = pd.Series(0.5, index=s.index, dtype=float)
    valid = s.notna()
    if valid.sum() == 0: return out
    pct = s[valid].rank(pct=True, method="average")
    out.loc[valid] = 1 - pct
    return out.clip(0, 1)


def safe_mode_ratio(s):
    s = s.dropna()
    if len(s) == 0: return np.nan
    vc = s.value_counts(dropna=True)
    return float(vc.iloc[0] / len(s))


def first_notna(s):
    s = s.dropna()
    return s.iloc[0] if len(s) > 0 else np.nan


def load_and_clean_data(uploaded_file):
    # 适配修改：从上传的文件对象读取，而不是本地路径
    df = pd.read_excel(uploaded_file)
    df = df.rename(columns={
        "入帐时间": "入帐时间", "入帐时间.1": "入帐时间_重复", "帐号": "帐号",
        "对方帐号": "对方帐号", "POS号": "POS号", "交易额": "交易额",
        "消费类型": "消费类型", "一卡通号": "一卡通号",
    })
    df["入帐时间"] = pd.to_datetime(df["入帐时间"], errors="coerce")
    df["交易额"] = pd.to_numeric(df["交易额"], errors="coerce")
    df["POS号"] = df["POS号"].astype(str).str.strip()
    df["消费类型"] = df["消费类型"].astype(str).str.strip()
    df["一卡通号"] = df["一卡通号"].astype(str).str.strip()
    df = df.dropna(subset=["入帐时间", "交易额"])
    df = df[df["一卡通号"].notna() & (df["一卡通号"] != "")]
    df["日期"] = df["入帐时间"].dt.date
    df["小时"] = df["入帐时间"].dt.hour
    df["金额绝对值"] = df["交易额"].abs()
    return df


def split_transactions(df):
    spend_df = df[df["交易额"] < 0].copy()
    spend_df["支出金额"] = spend_df["交易额"].abs()
    deposit_df = df[df["交易额"] > 0].copy()
    deposit_df["充值金额"] = deposit_df["交易额"]
    zero_df = df[df["交易额"] == 0].copy()
    return spend_df, deposit_df, zero_df


def build_spend_features(spend_df):
    def group_metrics(g):
        spend_amount = g["支出金额"]
        pos_counts = g["POS号"].value_counts(dropna=True)
        dominant_pos_ratio = float(pos_counts.iloc[0] / len(g)) if len(pos_counts) > 0 else np.nan
        amount_mode_ratio = safe_mode_ratio(spend_amount.round(2))
        breakfast_cnt = ((g["小时"] >= 5) & (g["小时"] < 9)).sum()
        lunch_cnt = ((g["小时"] >= 10) & (g["小时"] < 14)).sum()
        dinner_cnt = ((g["小时"] >= 16) & (g["小时"] < 20)).sum()
        meal_cnt = breakfast_cnt + lunch_cnt + dinner_cnt
        night_cnt = ((g["小时"] >= 21) | (g["小时"] < 5)).sum()
        high_cnt = (spend_amount >= HIGH_AMOUNT_THRESHOLD).sum()
        one_yuan_cnt = ((spend_amount >= 0.99) & (spend_amount <= 1.01)).sum()
        hash_type_cnt = (g["消费类型"] == "###").sum()
        return pd.Series({
            "帐号": first_notna(g["帐号"]) if "帐号" in g.columns else np.nan,
            "总支出": spend_amount.sum(), "消费笔数": len(g),
            "活跃消费天数": g["日期"].nunique(),
            "日均活跃支出": spend_amount.sum() / max(g["日期"].nunique(), 1),
            "单笔支出中位数": spend_amount.median(),
            "高额消费笔数": high_cnt, "高额消费占比": high_cnt / max(len(g), 1),
            "餐时消费笔数": meal_cnt, "餐时消费占比": meal_cnt / max(len(g), 1),
            "夜间消费占比": night_cnt / max(len(g), 1),
            "POS种类数": g["POS号"].nunique(dropna=True),
            "主导POS占比": dominant_pos_ratio, "金额众数占比": amount_mode_ratio,
            "一元消费占比": one_yuan_cnt / max(len(g), 1),
            "###类型占比": hash_type_cnt / max(len(g), 1),
        })

    return spend_df.groupby("一卡通号").apply(group_metrics).reset_index()


def build_deposit_features(deposit_df):
    if deposit_df.empty:
        return pd.DataFrame(columns=["一卡通号", "平均单次充值额"])
    return deposit_df.groupby("一卡通号", as_index=False).agg(平均单次充值额=("充值金额", "mean"))


def flag_group_a(df):
    cond_a = (
            (df["活跃消费天数"] >= MIN_ACTIVE_DAYS) & (df["消费笔数"] >= MIN_SPEND_COUNT) &
            (df["单笔支出中位数"] <= A_MEDIAN_TXN_MAX) & (df["一元消费占比"] >= A_ONE_YUAN_RATIO_MIN) &
            (df["主导POS占比"] >= A_DOMINANT_POS_RATIO_MIN) & (df["POS种类数"] <= A_POS_NUNIQUE_MAX) &
            ((df["###类型占比"] >= A_HASH_RATIO_MIN) | (df["金额众数占比"] >= A_MODE_RATIO_MIN))
    )
    df["高疑似A"] = np.where(cond_a, 1, 0)
    return df


def score_candidates(feature_df):
    df = feature_df.copy()
    df["活跃样本"] = np.where((df["活跃消费天数"] >= MIN_ACTIVE_DAYS) & (df["消费笔数"] >= MIN_SPEND_COUNT), 1, 0)
    active_non_a = df[(df["活跃样本"] == 1) & (df["高疑似A"] == 0)].copy()
    if active_non_a.empty: return df, active_non_a

    for key in ["总支出", "日均活跃支出", "单笔支出中位数", "高额消费占比"]:
        active_non_a[f"score_{key}"] = inverse_percentile_score(active_non_a[key])

    active_non_a["score_avg_deposit"] = inverse_percentile_score(active_non_a["平均单次充值额"])
    active_non_a["疑似经济困难得分"] = 0.0
    for col, w in WEIGHTS.items():
        if col in active_non_a.columns:
            active_non_a["疑似经济困难得分"] += active_non_a[col] * w

    p95 = active_non_a["疑似经济困难得分"].quantile(0.95)
    p85 = active_non_a["疑似经济困难得分"].quantile(0.85)
    active_non_a["疑似等级"] = "低疑似"
    active_non_a.loc[active_non_a["疑似经济困难得分"] >= p85, "疑似等级"] = "中疑似"
    active_non_a.loc[active_non_a["疑似经济困难得分"] >= p95, "疑似等级"] = "高疑似B_一般低消费组"
    return df, active_non_a


# ==========================================
# --- 3. 网页界面布局 (Streamlit UI) ---
# ==========================================

st.title("🛡️ 疑似经济困难学生初筛系统")
st.markdown("---")

# 侧边栏参数微调
with st.sidebar:
    st.header("⚙️ 核心参数微调")
    MIN_ACTIVE_DAYS = st.slider("活跃天数阈值", 5, 30, 15)
    HIGH_AMOUNT_THRESHOLD = st.number_input("高额消费认定标准(元)", value=20)

# 文件上传区
uploaded_file = st.file_uploader("📂 请上传一卡通流水 Excel 文件", type=["xlsx"])

if uploaded_file:
    with st.status("🚀 正在运行核心算法进行深度初筛...", expanded=True) as status:
        # 依次运行你的逻辑函数
        raw_df = load_and_clean_data(uploaded_file)
        spend_df, deposit_df, _ = split_transactions(raw_df)
        spend_features = build_spend_features(spend_df)
        deposit_features = build_deposit_features(deposit_df)
        feature_df = spend_features.merge(deposit_features, on="一卡通号", how="left")
        feature_df = flag_group_a(feature_df)
        all_feats, scored_df = score_candidates(feature_df)
        status.update(label="✅ 分析完成！", state="complete")

    # --- 1. 数据概览 (对应图4顶部卡片) ---
    st.subheader("📊 疑似等级分布")
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    kpi1.metric("分析总人数", f"{len(feature_df)} 人")
    kpi2.metric("高疑似A (特殊重复消费)", f"{int(all_feats['高疑似A'].sum())} 人")
    kpi3.metric("高疑似B (一般低消费)", f"{len(scored_df[scored_df['疑似等级'] == '高疑似B_一般低消费组'])} 人")
    kpi4.metric("中疑似", f"{len(scored_df[scored_df['疑似等级'] == '中疑似'])} 人")

    # --- 2. 可视化分析 (对应图4中间图表) ---
    col_l, col_r = st.columns(2)
    with col_l:
        st.write("**疑似得分分布直方图**")
        fig_hist = px.histogram(scored_df, x="疑似经济困难得分", color="疑似等级",
                                nbins=30, color_discrete_sequence=px.colors.qualitative.Pastel)
        st.plotly_chart(fig_hist, use_container_width=True)

    with col_r:
        st.write("**消费结构主要占比**")
        avg_meal = scored_df["餐时消费占比"].mean()
        avg_night = scored_df["夜间消费占比"].mean()
        fig_pie = px.pie(values=[avg_meal, avg_night, 1 - avg_meal - avg_night],
                         names=['餐时消费', '夜间消费', '其他'], hole=0.4)
        st.plotly_chart(fig_pie, use_container_width=True)

    # --- 3. 详细名单与下载 ---
    st.subheader("📋 疑似名单明细")
    st.dataframe(scored_df[["一卡通号", "总支出", "日均活跃支出", "疑似等级", "疑似经济困难得分"]].sort_values(
        "疑似经济困难得分", ascending=False))

    # 导出 CSV
    csv = scored_df.to_csv(index=False).encode('utf-8-sig')
    st.download_button("📥 导出分析结果清单 (CSV)", data=csv, file_name="分析结果.csv", mime="text/csv")