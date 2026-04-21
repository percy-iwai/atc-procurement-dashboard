"""
航空管制調達DB インタラクティブダッシュボード
データソース: data/db/atc_procurement.db
  - contracts テーブル (10,578件: 航空局本局XLSX + TCAB + WCAB)
  - pas_airport_contracts テーブル (2,240件: 地方整備局PAS)
対象期間: FY2020-2024（一部FY2019/2025含む）

カラム:
  vendor_normalized : 法人番号ベース名寄せ後のベンダー名（ゴミはNULL）
  mega_category     : モノの種類ベース大括り分類
                      管制システム / 無線・レーダー / 通信設備 / 灯火・標識 /
                      電源・電気設備 / 土木・建築 / 航空機・検査 / 一般物品・庁費
  work_type         : 作業特性
                      製造・納入 / 工事・設置 / 保守・運用 / 設計・調査 / その他
"""

import json
import sqlite3
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

# ── ページ設定 ───────────────────────────────────────────────────
st.set_page_config(
    page_title="航空管制調達ダッシュボード",
    page_icon="✈️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── スタイル ────────────────────────────────────────────────────
st.markdown("""
<style>
  #MainMenu {visibility: hidden;}
  footer {visibility: hidden;}
  header {visibility: hidden;}
  div[data-testid="metric-container"] {
    background: #1e1e2e;
    border: 1px solid #313244;
    border-radius: 8px;
    padding: 14px 18px;
  }
  div[data-testid="metric-container"] label {
    font-size: 0.78rem;
    color: #a6b0cf;
  }
  .stTabs [data-baseweb="tab"] {font-size: 0.9rem;}
  .block-container {padding-top: 1.5rem;}
</style>
""", unsafe_allow_html=True)

# ── パス定数 ─────────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
DB_PATH = BASE_DIR / "data" / "atc_procurement.db"
BUDGET_PATH = BASE_DIR / "data" / "output" / "budget_vs_db_5yr.json"
TEMPLATE = "plotly_dark"

# ── ヘルパー関数 ──────────────────────────────────────────────────
def classify_bid(bm: str | None) -> str:
    if not bm:
        return "その他"
    if "一般競争" in bm or "公募型競争" in bm:
        return "一般競争"
    if "随意" in bm:
        return "随意契約"
    if "企画競争" in bm or "プロポーザル" in bm:
        return "企画競争"
    if "指名競争" in bm:
        return "指名競争"
    return "その他"


SOURCE_LABEL = {
    None: "航空局本局",
    "tcab": "東京航空局(TCAB)",
    "wcab": "大阪航空局(WCAB)",
}


def fmt_oku(v: float) -> str:
    """億円単位の金額を人間が読みやすい形式にフォーマット"""
    if v >= 10_000:
        return f"{v / 10_000:.2f} 兆円"
    if v >= 1:
        return f"{v:,.1f} 億円"
    return f"{v * 100:.1f} 百万円"


def fmt_man(v: float) -> str:
    """万円単位でフォーマット"""
    if v >= 1e4:
        return f"{v / 1e4:,.1f} 億円"
    return f"{v:,.0f} 万円"


# ── データ読み込み（キャッシュ付き） ────────────────────────────────
@st.cache_data
def load_contracts() -> pd.DataFrame:
    con = sqlite3.connect(DB_PATH)
    df = pd.read_sql("SELECT * FROM contracts", con)
    con.close()

    df["bid_type"] = df["bid_method"].apply(classify_bid)
    df["source_org"] = df["source"].map(SOURCE_LABEL).fillna("不明")
    df["amount_oku"] = df["contract_amount"].fillna(0) / 1e8
    df["fiscal_year"] = df["fiscal_year"].astype("Int64")

    # vendor_normalized フォールバック（カラムがない旧DBでも動く）
    if "vendor_normalized" not in df.columns:
        df["vendor_normalized"] = df["vendor_name"]
    if "mega_category" not in df.columns:
        df["mega_category"] = df["category"]
    if "work_type" not in df.columns:
        df["work_type"] = "その他"

    # 月次パース
    df["year_month"] = (
        pd.to_datetime(df["contract_date"], format="%Y%m%d", errors="coerce")
        .dt.to_period("M")
        .astype(str)
        .replace("NaT", pd.NA)
    )
    return df


@st.cache_data
def load_pas() -> pd.DataFrame:
    con = sqlite3.connect(DB_PATH)
    df = pd.read_sql("SELECT * FROM pas_airport_contracts", con)
    con.close()

    df["bid_type"] = df["bid_method"].apply(classify_bid)
    df["amount_oku"] = df["award_amount"].fillna(0) / 1e8
    df["fiscal_year"] = df["fiscal_year"].astype("Int64")
    df["source_org"] = "地方整備局(PAS)"

    if "vendor_normalized" not in df.columns:
        df["vendor_normalized"] = df["vendor_name"]
    if "mega_category" not in df.columns:
        df["mega_category"] = df["category"]

    df["year_month"] = (
        pd.to_datetime(df["award_date"], format="%Y%m%d", errors="coerce")
        .dt.to_period("M")
        .astype(str)
        .replace("NaT", pd.NA)
    )
    return df


@st.cache_data
def load_budget() -> dict:
    with open(BUDGET_PATH, encoding="utf-8") as f:
        return json.load(f)


def make_combined(fc: pd.DataFrame, fp: pd.DataFrame) -> pd.DataFrame:
    """contracts と PAS を共通スキーマで結合"""
    c = fc.assign(
        amount=fc["contract_amount"].fillna(0),
        organization=fc["organization"],
        source_type="contracts",
    )[["fiscal_year", "bid_type", "contract_name", "vendor_name", "vendor_normalized",
       "amount", "organization", "category", "mega_category", "work_type", "area_type",
       "system_name", "year_month", "source_org", "source_type", "amount_oku"]]

    p = fp.assign(
        amount=fp["award_amount"].fillna(0),
        organization=fp["bureau"] + "地整局",
        contract_name=fp["contract_name"],
        system_name=fp.get("work_type", pd.Series(dtype=str)),
        work_type=pd.NA,
        source_type="pas",
    )[["fiscal_year", "bid_type", "contract_name", "vendor_name", "vendor_normalized",
       "amount", "organization", "category", "mega_category", "work_type", "area_type",
       "system_name", "year_month", "source_org", "source_type", "amount_oku"]]

    return pd.concat([c, p], ignore_index=True)


# ── メイン ──────────────────────────────────────────────────────
def main():
    # タイトル
    st.markdown(
        "## ✈️ 航空管制調達DB ダッシュボード",
        unsafe_allow_html=False,
    )
    st.caption(
        "国土交通省 航空局・地方整備局 | FY2020–2024 | "
        "contracts (10,578件) + PAS (2,240件)"
    )

    # データ読み込み
    df_c = load_contracts()
    df_p = load_pas()
    budget = load_budget()

    # ── サイドバー ──────────────────────────────────────────────
    with st.sidebar:
        st.header("🔍 フィルタ")

        # 会計年度
        all_fy = sorted(
            set(df_c["fiscal_year"].dropna().astype(int).tolist())
            | set(df_p["fiscal_year"].dropna().astype(int).tolist())
        )
        sel_fy = st.multiselect("会計年度 (FY)", all_fy, default=all_fy)

        # 大カテゴリ (mega_category) — モノの種類ベース
        MEGA_ORDER = [
            "管制システム", "無線・レーダー", "通信設備", "灯火・標識",
            "電源・電気設備", "土木・建築", "航空機・検査", "一般物品・庁費",
        ]
        all_megas = [m for m in MEGA_ORDER if m in (
            set(df_c["mega_category"].dropna()) | set(df_p["mega_category"].dropna())
        )]
        sel_megas = st.multiselect("大カテゴリ（モノの種類）", all_megas, default=all_megas)

        # 作業種別 (work_type)
        WORK_TYPE_ORDER = ["製造・納入", "工事・設置", "保守・運用", "設計・調査", "その他"]
        all_wts = [w for w in WORK_TYPE_ORDER if w in set(df_c["work_type"].dropna())]
        sel_wts = st.multiselect("作業種別（work_type）", all_wts, default=all_wts)

        # カテゴリ
        all_cats = sorted(
            (set(df_c["category"].dropna()) | set(df_p["category"].dropna()))
            - {"不明"}
        ) + ["不明"]
        sel_cats = st.multiselect("カテゴリ（詳細）", all_cats, default=all_cats)

        # エリアタイプ
        all_areas = sorted(
            set(df_c["area_type"].dropna()) | set(df_p["area_type"].dropna())
        )
        sel_areas = st.multiselect("エリアタイプ", all_areas, default=all_areas)

        # システム名（contracts）
        all_systems = sorted(df_c["system_name"].dropna().unique())
        sel_systems = st.multiselect("システム名", all_systems, default=all_systems)

        # 組織（contracts）
        all_orgs = sorted(df_c["organization"].dropna().unique())
        sel_orgs = st.multiselect(
            "組織（contracts）",
            all_orgs,
            default=all_orgs,
            help="航空局本局・TCAB・WCAB の組織名",
        )

        # 地整局（PAS）
        all_bureaux = sorted(df_p["bureau"].dropna().unique())
        sel_bureaux = st.multiselect(
            "地方整備局（PAS）",
            all_bureaux,
            default=all_bureaux,
        )

        # 入札方式
        bid_options = ["一般競争", "随意契約", "企画競争", "指名競争", "その他"]
        sel_bids = st.multiselect("入札方式", bid_options, default=bid_options)

        # 金額レンジ（億円）
        max_oku = float(
            max(
                df_c["amount_oku"].max() if len(df_c) > 0 else 0,
                df_p["amount_oku"].max() if len(df_p) > 0 else 0,
            )
        )
        amount_range = st.slider(
            "金額レンジ（億円）",
            0.0,
            max_oku,
            (0.0, max_oku),
            step=0.1,
        )

        st.divider()
        data_src = st.radio(
            "表示データソース",
            ["両方（contracts + PAS）", "contracts のみ", "PAS のみ"],
            index=0,
        )

    # ── フィルタ適用 ───────────────────────────────────────────
    fc = df_c.copy()
    fp = df_p.copy()

    if sel_fy:
        fc = fc[fc["fiscal_year"].isin(sel_fy)]
        fp = fp[fp["fiscal_year"].isin(sel_fy)]
    if sel_megas:
        fc = fc[fc["mega_category"].fillna("その他").isin(sel_megas)]
        fp = fp[fp["mega_category"].fillna("その他").isin(sel_megas)]
    if sel_cats:
        fc = fc[fc["category"].fillna("不明").isin(sel_cats)]
        fp = fp[fp["category"].fillna("不明").isin(sel_cats)]
    if sel_areas:
        fc = fc[fc["area_type"].fillna("不明").isin(sel_areas)]
        fp = fp[fp["area_type"].fillna("不明").isin(sel_areas)]
    if sel_systems:
        fc = fc[fc["system_name"].isin(sel_systems) | fc["system_name"].isna()]
    if sel_orgs:
        fc = fc[fc["organization"].isin(sel_orgs) | fc["organization"].isna()]
    if sel_bureaux:
        fp = fp[fp["bureau"].isin(sel_bureaux)]
    if sel_bids:
        fc = fc[fc["bid_type"].isin(sel_bids)]
        fp = fp[fp["bid_type"].isin(sel_bids)]
    if sel_wts:
        fc = fc[fc["work_type"].fillna("その他").isin(sel_wts)]

    fc = fc[
        (fc["amount_oku"] >= amount_range[0]) & (fc["amount_oku"] <= amount_range[1])
    ]
    fp = fp[
        (fp["amount_oku"] >= amount_range[0]) & (fp["amount_oku"] <= amount_range[1])
    ]

    # データソース選択
    if data_src == "contracts のみ":
        fp = fp.iloc[0:0]
    elif data_src == "PAS のみ":
        fc = fc.iloc[0:0]

    combined = make_combined(fc, fp)

    # ── KPI 行 ─────────────────────────────────────────────────
    total_count = len(combined)
    total_oku = combined["amount"].sum() / 1e8
    avg_man = (combined["amount"].mean() / 1e4) if total_count > 0 else 0
    unique_cats = combined["category"].nunique()

    st.markdown("### 📊 KPI サマリー")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("総件数", f"{total_count:,} 件")
    k2.metric("総金額", fmt_oku(total_oku))
    k3.metric("平均単価", fmt_man(avg_man))
    k4.metric("カテゴリ数", str(unique_cats))

    sub1, sub2, sub3 = st.columns(3)
    sub1.metric(
        "contracts",
        f"{len(fc):,} 件",
        f"{fc['amount_oku'].sum():.0f} 億円",
    )
    sub2.metric(
        "PAS（地方整備局）",
        f"{len(fp):,} 件",
        f"{fp['amount_oku'].sum():.0f} 億円",
    )
    c_rate = (
        (combined["bid_type"] == "随意契約").sum() / total_count * 100
        if total_count > 0
        else 0
    )
    sub3.metric("随意契約率（件数）", f"{c_rate:.1f} %")

    st.divider()

    # ── Row 1: カテゴリ別 + FY トレンド ───────────────────────
    col_a, col_b = st.columns(2)

    with col_a:
        st.markdown("#### カテゴリ別 金額（横棒）")
        if not combined.empty:
            cat_df = (
                combined.groupby("category")["amount"]
                .sum()
                .reset_index()
            )
            cat_df["億円"] = cat_df["amount"] / 1e8
            cat_df = cat_df.sort_values("億円", ascending=True).tail(20)
            fig = px.bar(
                cat_df,
                x="億円",
                y="category",
                orientation="h",
                color="億円",
                color_continuous_scale="Blues_r",
                labels={"category": "カテゴリ", "億円": "金額（億円）"},
                template=TEMPLATE,
            )
            fig.update_layout(
                coloraxis_showscale=False,
                height=480,
                margin=dict(l=8, r=8, t=8, b=8),
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("データなし")

    with col_b:
        st.markdown("#### FY別 金額トレンド（ソース別積み上げ）")
        if not combined.empty:
            fy_df = (
                combined.groupby(["fiscal_year", "source_type"])["amount"]
                .sum()
                .reset_index()
            )
            fy_df["億円"] = fy_df["amount"] / 1e8
            fy_df["ソース"] = fy_df["source_type"].map(
                {"contracts": "本局+TCAB+WCAB", "pas": "地方整備局(PAS)"}
            )
            fig = px.bar(
                fy_df,
                x="fiscal_year",
                y="億円",
                color="ソース",
                barmode="stack",
                labels={"fiscal_year": "会計年度", "億円": "金額（億円）"},
                template=TEMPLATE,
                color_discrete_sequence=["#4472c4", "#70ad47"],
                text="億円",
            )
            fig.update_traces(texttemplate="%{text:.0f}", textposition="inside")
            fig.update_layout(
                height=480,
                margin=dict(l=8, r=8, t=8, b=8),
                legend=dict(orientation="h", yanchor="bottom", y=1.01),
                xaxis=dict(type="category"),
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("データなし")

    # ── Row 2: システム名 + ベンダー上位 ───────────────────────
    col_c, col_d = st.columns(2)

    with col_c:
        st.markdown("#### システム名別 金額（contracts）")
        sys_data = fc[fc["system_name"].notna()]
        if not sys_data.empty:
            sys_df = (
                sys_data.groupby("system_name")["contract_amount"]
                .sum()
                .reset_index()
            )
            sys_df["億円"] = sys_df["contract_amount"] / 1e8
            sys_df = sys_df.sort_values("億円", ascending=True)
            fig = px.bar(
                sys_df,
                x="億円",
                y="system_name",
                orientation="h",
                color="億円",
                color_continuous_scale="Teal",
                labels={"system_name": "システム", "億円": "金額（億円）"},
                template=TEMPLATE,
            )
            fig.update_layout(
                coloraxis_showscale=False,
                height=420,
                margin=dict(l=8, r=8, t=8, b=8),
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("システム名データなし")

    with col_d:
        st.markdown("#### ベンダー上位 20社（名寄せ後・金額）")
        v_data = combined[combined["vendor_normalized"].notna()]
        if not v_data.empty:
            v_df = (
                v_data.groupby("vendor_normalized")["amount"]
                .sum()
                .reset_index()
            )
            v_df["億円"] = v_df["amount"] / 1e8
            v_df = v_df.sort_values("億円", ascending=True).tail(20)
            v_df["表示名"] = v_df["vendor_normalized"].str[:22]
            fig = px.bar(
                v_df,
                x="億円",
                y="表示名",
                orientation="h",
                color="億円",
                color_continuous_scale="Oranges",
                labels={"表示名": "ベンダー", "億円": "金額（億円）"},
                template=TEMPLATE,
            )
            fig.update_layout(
                coloraxis_showscale=False,
                height=520,
                margin=dict(l=8, r=8, t=8, b=8),
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("データなし")

    # ── Row 3: 組織別ドーナツ + 月別推移 ───────────────────────
    col_e, col_f = st.columns(2)

    with col_e:
        st.markdown("#### 組織別 金額（ドーナツ）")
        if not combined.empty:
            org_df = (
                combined.groupby("source_org")["amount"]
                .sum()
                .reset_index()
            )
            org_df["億円"] = org_df["amount"] / 1e8
            fig = px.pie(
                org_df,
                values="億円",
                names="source_org",
                hole=0.5,
                template=TEMPLATE,
                color_discrete_sequence=px.colors.qualitative.Set2,
            )
            fig.update_traces(
                textinfo="percent+label",
                hovertemplate="%{label}<br>%{value:.1f}億円<br>%{percent}",
            )
            fig.update_layout(
                height=420,
                margin=dict(l=8, r=8, t=30, b=8),
                showlegend=True,
                legend=dict(orientation="v", x=1.0, y=0.5),
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("データなし")

    with col_f:
        st.markdown("#### 月別 契約件数推移")
        m_data = combined[combined["year_month"].notna() & (combined["year_month"] != "NaT")]
        if not m_data.empty:
            m_df = (
                m_data.groupby("year_month")
                .agg(件数=("amount", "count"), 合計億円=("amount_oku", "sum"))
                .reset_index()
                .sort_values("year_month")
            )
            fig = make_subplots(specs=[[{"secondary_y": True}]])
            fig.add_trace(
                go.Bar(
                    x=m_df["year_month"],
                    y=m_df["件数"],
                    name="件数",
                    marker_color="#4472c4",
                    opacity=0.8,
                ),
                secondary_y=False,
            )
            fig.add_trace(
                go.Scatter(
                    x=m_df["year_month"],
                    y=m_df["合計億円"],
                    name="金額（億円）",
                    mode="lines+markers",
                    line=dict(color="#ffc000", width=2),
                    marker=dict(size=4),
                ),
                secondary_y=True,
            )
            fig.update_layout(
                template=TEMPLATE,
                height=420,
                margin=dict(l=8, r=8, t=8, b=8),
                legend=dict(orientation="h", yanchor="bottom", y=1.01),
                xaxis=dict(tickangle=-45),
            )
            fig.update_yaxes(title_text="件数", secondary_y=False)
            fig.update_yaxes(title_text="金額（億円）", secondary_y=True)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("日付データなし")

    # ── 大カテゴリ（mega_category）グラフ ──────────────────────
    st.markdown("#### 大カテゴリ別 金額・件数（モノの種類）")
    if not combined.empty:
        MEGA_COLORS = {
            "管制システム":   "#4472c4",
            "無線・レーダー": "#9e2a2b",
            "通信設備":       "#5e8c61",
            "灯火・標識":     "#ffc000",
            "電源・電気設備": "#ed7d31",
            "土木・建築":     "#70ad47",
            "航空機・検査":   "#8e44ad",
            "一般物品・庁費": "#7f7f7f",
        }
        MEGA_ORDER = [
            "管制システム", "無線・レーダー", "通信設備", "灯火・標識",
            "電源・電気設備", "土木・建築", "航空機・検査", "一般物品・庁費",
        ]
        mg_col1, mg_col2 = st.columns(2)

        with mg_col1:
            mg_df = (
                combined.groupby("mega_category")
                .agg(億円=("amount", lambda x: x.sum() / 1e8), 件数=("amount", "count"))
                .reset_index()
                .sort_values("億円", ascending=False)
            )
            fig_mg = px.bar(
                mg_df,
                x="mega_category",
                y="億円",
                color="mega_category",
                color_discrete_map=MEGA_COLORS,
                category_orders={"mega_category": MEGA_ORDER},
                text="件数",
                labels={"mega_category": "大カテゴリ", "億円": "金額（億円）"},
                template=TEMPLATE,
            )
            fig_mg.update_traces(texttemplate="%{text}件", textposition="outside")
            fig_mg.update_layout(
                showlegend=False,
                height=380,
                margin=dict(l=8, r=8, t=8, b=8),
                xaxis=dict(tickangle=-25),
            )
            st.plotly_chart(fig_mg, use_container_width=True)

        with mg_col2:
            # FY × mega_category の積み上げ
            mg_fy = (
                combined.groupby(["fiscal_year", "mega_category"])["amount"]
                .sum()
                .reset_index()
            )
            mg_fy["億円"] = mg_fy["amount"] / 1e8
            fig_fy = px.bar(
                mg_fy,
                x="fiscal_year",
                y="億円",
                color="mega_category",
                color_discrete_map=MEGA_COLORS,
                category_orders={"mega_category": MEGA_ORDER},
                barmode="stack",
                labels={"fiscal_year": "会計年度", "億円": "金額（億円）",
                        "mega_category": "大カテゴリ"},
                template=TEMPLATE,
            )
            fig_fy.update_layout(
                height=380,
                margin=dict(l=8, r=8, t=8, b=8),
                legend=dict(orientation="h", yanchor="bottom", y=1.01,
                            font=dict(size=9)),
                xaxis=dict(type="category"),
            )
            st.plotly_chart(fig_fy, use_container_width=True)

    # ── work_type グラフ ────────────────────────────────────────
    st.markdown("#### 作業種別（work_type）× 大カテゴリ — contracts")
    if not fc.empty:
        WORK_COLORS = {
            "製造・納入": "#4472c4",
            "工事・設置": "#70ad47",
            "保守・運用": "#ffc000",
            "設計・調査": "#9dc3e6",
            "その他":     "#7f7f7f",
        }
        WORK_ORDER = ["製造・納入", "工事・設置", "保守・運用", "設計・調査", "その他"]
        wt_col1, wt_col2 = st.columns(2)

        with wt_col1:
            wt_df = (
                fc.groupby("work_type")
                .agg(億円=("amount_oku", "sum"), 件数=("amount_oku", "count"))
                .reset_index()
            )
            fig_wt = px.bar(
                wt_df,
                x="work_type",
                y="億円",
                color="work_type",
                color_discrete_map=WORK_COLORS,
                category_orders={"work_type": WORK_ORDER},
                text="件数",
                labels={"work_type": "作業種別", "億円": "金額（億円）"},
                template=TEMPLATE,
            )
            fig_wt.update_traces(texttemplate="%{text}件", textposition="outside")
            fig_wt.update_layout(
                showlegend=False, height=340,
                margin=dict(l=8, r=8, t=8, b=8),
            )
            st.plotly_chart(fig_wt, use_container_width=True)

        with wt_col2:
            # mega × work_type ヒートマップ
            heat = (
                fc.groupby(["mega_category", "work_type"])["amount_oku"]
                .sum()
                .reset_index()
                .pivot(index="mega_category", columns="work_type", values="amount_oku")
                .fillna(0)
                .reindex(columns=WORK_ORDER, fill_value=0)
                .reindex(MEGA_ORDER[::-1], fill_value=0)  # 上から大→小
            )
            fig_heat = go.Figure(go.Heatmap(
                z=heat.values,
                x=heat.columns.tolist(),
                y=heat.index.tolist(),
                colorscale="Blues",
                text=[[f"{v:.0f}億" for v in row] for row in heat.values],
                texttemplate="%{text}",
                hovertemplate="%{y} × %{x}<br>%{z:.1f}億円<extra></extra>",
            ))
            fig_heat.update_layout(
                template=TEMPLATE,
                height=340,
                margin=dict(l=8, r=8, t=8, b=8),
                xaxis_title="作業種別",
                yaxis_title="大カテゴリ",
            )
            st.plotly_chart(fig_heat, use_container_width=True)

    # ── 組織別棒グラフ（contracts） ─────────────────────────────
    st.markdown("#### 組織別 金額・件数（contracts）")
    org_c = fc[fc["organization"].notna()]
    if not org_c.empty:
        org_bar = (
            org_c.groupby("organization")
            .agg(億円=("amount_oku", "sum"), 件数=("contract_amount", "count"))
            .reset_index()
            .sort_values("億円", ascending=False)
            .head(20)
        )
        fig = px.bar(
            org_bar,
            x="organization",
            y="億円",
            color="億円",
            color_continuous_scale="Viridis",
            text="件数",
            labels={"organization": "組織", "億円": "金額（億円）"},
            template=TEMPLATE,
        )
        fig.update_traces(
            texttemplate="%{text}件",
            textposition="outside",
        )
        fig.update_layout(
            coloraxis_showscale=False,
            height=380,
            margin=dict(l=8, r=8, t=8, b=8),
            xaxis=dict(tickangle=-30),
        )
        st.plotly_chart(fig, use_container_width=True)

    # ── 予算比較（5年トレンド） ─────────────────────────────────
    st.markdown("#### 📈 予算比較 — 調達母数 vs DB収録額（5年トレンド）")
    bgt_rows = []
    for fy_key in sorted(budget["fiscal_years"].keys()):
        d = budget["fiscal_years"][fy_key]
        bgt_rows.append(
            {
                "FY": int(fy_key),
                "調達母数（億円）": d["budget"]["procurement_base_oku"],
                "DB収録額（億円）": d["db"]["combined_oku"],
                "網羅率(%)": d["coverage"]["coverage_pct"],
            }
        )
    bgt_df = pd.DataFrame(bgt_rows)

    fig2 = make_subplots(specs=[[{"secondary_y": True}]])
    fig2.add_trace(
        go.Bar(
            name="調達母数",
            x=bgt_df["FY"],
            y=bgt_df["調達母数（億円）"],
            marker_color="#4472c4",
            opacity=0.7,
        ),
        secondary_y=False,
    )
    fig2.add_trace(
        go.Bar(
            name="DB収録額",
            x=bgt_df["FY"],
            y=bgt_df["DB収録額（億円）"],
            marker_color="#70ad47",
            opacity=0.9,
        ),
        secondary_y=False,
    )
    fig2.add_trace(
        go.Scatter(
            name="網羅率(%)",
            x=bgt_df["FY"],
            y=bgt_df["網羅率(%)"],
            mode="lines+markers+text",
            line=dict(color="#ffc000", width=2),
            text=[f"{v:.1f}%" for v in bgt_df["網羅率(%)"]],
            textposition="top center",
        ),
        secondary_y=True,
    )
    fig2.update_layout(
        template=TEMPLATE,
        barmode="group",
        height=420,
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        margin=dict(l=8, r=8, t=8, b=8),
        xaxis=dict(type="category"),
    )
    fig2.update_yaxes(title_text="金額（億円）", secondary_y=False)
    fig2.update_yaxes(
        title_text="網羅率（%）",
        secondary_y=True,
        range=[0, 170],
        ticksuffix="%",
    )
    st.plotly_chart(fig2, use_container_width=True)

    # ── データテーブル ──────────────────────────────────────────
    st.markdown("#### 📋 フィルタ済みデータ一覧")
    tab1, tab2 = st.tabs(
        [f"contracts ({len(fc):,} 件)", f"PAS — 地方整備局 ({len(fp):,} 件)"]
    )

    with tab1:
        show_c = fc[
            [
                "fiscal_year", "bid_type", "source_org", "organization",
                "mega_category", "work_type", "category", "area_type", "system_name",
                "contract_name", "vendor_normalized", "vendor_name",
                "contract_amount", "contract_date",
            ]
        ].copy()
        show_c.columns = [
            "FY", "入札方式", "ソース機関", "組織",
            "大カテゴリ", "作業種別", "カテゴリ", "エリア", "システム",
            "件名", "ベンダー（名寄せ）", "ベンダー（元）",
            "金額（円）", "契約日",
        ]
        st.dataframe(show_c, use_container_width=True, height=320)
        csv_c = show_c.to_csv(index=False, encoding="utf-8-sig")
        st.download_button(
            "📥 contracts CSV ダウンロード",
            csv_c,
            file_name="contracts_filtered.csv",
            mime="text/csv",
        )

    with tab2:
        show_p = fp[
            [
                "fiscal_year", "bid_type", "bureau",
                "mega_category", "category", "area_type", "work_type",
                "contract_name", "vendor_normalized", "award_amount", "award_date",
            ]
        ].copy()
        show_p.columns = [
            "FY", "入札方式", "地整局",
            "大カテゴリ", "カテゴリ", "エリア", "工種",
            "件名", "ベンダー（名寄せ）", "落札額（円）", "落札日",
        ]
        st.dataframe(show_p, use_container_width=True, height=320)
        csv_p = show_p.to_csv(index=False, encoding="utf-8-sig")
        st.download_button(
            "📥 PAS CSV ダウンロード",
            csv_p,
            file_name="pas_filtered.csv",
            mime="text/csv",
        )


if __name__ == "__main__":
    main()
