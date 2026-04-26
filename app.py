"""
国交省航空関連調達DB インタラクティブダッシュボード
データソース: data/db/atc_procurement.db
  - contracts テーブル (航空局本局XLSX + TCAB + WCAB)
  - pas_airport_contracts テーブル (地方整備局PAS)
対象期間: FY2013〜FY2025

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
    page_title="国交省航空関連調達DB ダッシュボード",
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
  div[data-testid="stDialog"] > div {
    width: 95vw !important;
    max-width: 95vw !important;
    height: 90vh !important;
    max-height: 90vh !important;
  }
</style>
""", unsafe_allow_html=True)

# ── パス定数 ─────────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
DB_PATH = BASE_DIR / "data" / "atc_procurement.db"
BUDGET_PATH = BASE_DIR / "data" / "output" / "budget_vs_db_5yr.json"
ALL_YEARS_BUDGET_PATH = BASE_DIR / "data" / "output" / "budget_all_years.json"
TEMPLATE = "plotly_dark"

# ── システム名 → 参照先対応 ─────────────────────────────────────
PDF_BASE  = "https://www.mlit.go.jp/koku/content/001743241.pdf"
CIDAI_BASE = "https://www.mlit.go.jp/koku/cidai/"

# ("PDF", page) → 航空保安業務の概要PDF
# ("URL", url)  → 外部サイト（CIDAI等）
SYSTEM_REFERENCES: dict[str, tuple[str, int | str]] = {
    "TAPS": ("PDF", 56), "TEPS": ("PDF", 54), "TOPS": ("PDF", 60), "TEAM": ("PDF", 58), "FACE": ("PDF", 62),
    "HARP": ("PDF", 63), "ICAP": ("PDF", 63),
    "ILS": ("PDF", 42), "VOR/DME": ("PDF", 44), "ASR": ("PDF", 46), "ARSR": ("PDF", 48),
    "MLAT": ("PDF", 50), "SSR": ("PDF", 46), "SBAS": ("PDF", 51), "GBAS": ("PDF", 51),
    "NDB": ("PDF", 44), "RVA": ("PDF", 50),
    "CPDLC": ("PDF", 60), "ADS-B": ("PDF", 50),
    "PAPI": ("PDF", 52), "UPS": ("PDF", 66),
    "ATFM": ("PDF", 28), "ATIS": ("PDF", 30), "SWIM": ("PDF", 73), "MASS": ("PDF", 73),
    "RCAG": ("PDF", 48), "MAPS": ("PDF", 34), "RISE": ("PDF", 34), "MISE": ("PDF", 38), "SMC": ("PDF", 38),
    "ADEX": ("PDF", 60), "VOR": ("PDF", 45), "DME": ("PDF", 45), "DLCS": ("PDF", 90),
    # CIDAI（日本の空港技術ショーケース）掲載システム — PDFに記載なし
    "CCS":  ("URL", f"{CIDAI_BASE}#No_3-23"),  # 通信制御装置
    "TRCS": ("URL", f"{CIDAI_BASE}#No_3-4"),   # 非常用ターミナルレーダー管制装置（ARTS/VCCS言及あり）
    "ACTS": ("URL", f"{CIDAI_BASE}#No_2-5"),   # 飛行場管制訓練システム
    "EVA":  ("URL", f"{CIDAI_BASE}#No_3-5"),   # 非常用管制塔システム
    "ARTS": ("URL", "https://www.mlit.go.jp/koku/jans/dogu/system.html"),
    "CRMS": ("URL", "https://www.mlit.go.jp/koku/jans/dogu/system.html"),
    "FDPS": ("URL", "https://ja.wikipedia.org/wiki/%E9%A3%9B%E8%A1%8C%E8%A8%88%E7%94%BB%E6%83%85%E5%A0%B1%E7%AE%A1%E7%90%86%E3%82%B7%E3%82%B9%E3%83%86%E3%83%A0"),
}

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


def show_drilldown(df: pd.DataFrame, label: str, max_rows: int = 200):
    """クリックされた要素のドリルダウンテーブルを表示"""
    if df.empty:
        st.info(f"{label}: データなし")
        return
    DESIRED_COLS = {
        "fiscal_year": "FY",
        "bid_type": "入札方式",
        "source_org": "ソース機関",
        "organization": "組織",
        "mega_category": "大カテゴリ",
        "work_type": "作業種別",
        "category": "カテゴリ",
        "system_name": "システム",
        "contract_name": "件名",
        "vendor_normalized": "ベンダー",
        "amount": "金額（円）",
        "year_month": "年月",
    }
    cols = {k: v for k, v in DESIRED_COLS.items() if k in df.columns}
    disp = df.sort_values("amount", ascending=False)[list(cols.keys())].head(max_rows).rename(columns=cols)
    if "金額（円）" in disp.columns:
        disp = disp.copy()
        disp["金額（円）"] = disp["金額（円）"].apply(lambda x: f"{x:,.0f}" if pd.notna(x) else "")
    st.markdown(f"**▼ {label} — {len(df):,} 件**（最大{max_rows}行表示）")
    st.dataframe(disp, use_container_width=True, height=240)


# ── モーダルダイアログ（Streamlit 1.35+） ────────────────────────────
try:
    _st_ver = tuple(int(x) for x in st.__version__.split(".")[:2])
    _HAS_DIALOG = _st_ver >= (1, 35)
except Exception:
    _HAS_DIALOG = False

_DRILLDOWN_COLS = {
    "fiscal_year": "FY",
    "bid_type": "入札方式",
    "source_org": "ソース機関",
    "organization": "組織",
    "mega_category": "大カテゴリ",
    "work_type": "作業種別",
    "category": "カテゴリ",
    "system_name": "システム",
    "contract_name": "件名",
    "vendor_normalized": "ベンダー",
    "amount": "金額（円）",
    "year_month": "年月",
}

if _HAS_DIALOG:
    @st.dialog("ドリルダウン", width="large")
    def show_drilldown_dialog(df: pd.DataFrame, title: str, max_rows: int = 200):
        cols = {k: v for k, v in _DRILLDOWN_COLS.items() if k in df.columns}
        disp = df.sort_values("amount", ascending=False)[list(cols.keys())].head(max_rows).rename(columns=cols)
        if "金額（円）" in disp.columns:
            disp = disp.copy()
            disp["金額（円）"] = disp["金額（円）"].apply(lambda x: f"{x:,.0f}" if pd.notna(x) else "")
        st.subheader(title)
        st.dataframe(disp, use_container_width=True, height=600)
        st.write(f"**{len(df):,}件** / **{df['amount'].sum() / 1e8:,.1f}億円**")
else:
    def show_drilldown_dialog(df: pd.DataFrame, title: str, max_rows: int = 200):
        show_drilldown(df, title, max_rows)


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

    if "vendor_normalized" not in df.columns:
        df["vendor_normalized"] = df["vendor_name"]
    if "mega_category" not in df.columns:
        df["mega_category"] = df["category"]
    if "work_type" not in df.columns:
        df["work_type"] = "その他"

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


@st.cache_data
def load_all_years_budget() -> dict:
    with open(ALL_YEARS_BUDGET_PATH, encoding="utf-8") as f:
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
    # 変更1: タイトル変更
    st.markdown("## ✈️ 国交省航空関連調達DB ダッシュボード")

    # 変更2: データソース説明（折りたたみ、デフォルト閉じ）
    with st.expander("データソース・注記", expanded=False):
        st.markdown("""
**データソース:**
- [航空局本局](https://www.mlit.go.jp/koku/koku_tk1_000033.html)
- [東京航空局](https://www.cab.mlit.go.jp/tcab/contract/)
- [大阪航空局](https://www.cab.mlit.go.jp/wcab/contract/published.html)
- [地方整備局PAS](https://www.pas.ysk.nilim.go.jp/)

**期間:** FY2013〜FY2025（過去分は[WARP（国立国会図書館）](https://warp.ndl.go.jp/)で収集）
FY2019の地方整備局データはアーカイブ欠落のため未収録。
**カテゴリ分類:** [航空保安業務の概要（2025）](https://www.mlit.go.jp/koku/content/001743241.pdf)を参照。
""")

    # データ読み込み
    df_c = load_contracts()
    df_p = load_pas()
    budget = load_budget()

    # ── 大カテゴリ定数（サイドバーと本体で共用） ───────────────
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

    # ── サイドバー ──────────────────────────────────────────────
    with st.sidebar:
        st.header("🔍 フィルタ")

        all_fy = sorted(
            set(df_c["fiscal_year"].dropna().astype(int).tolist())
            | set(df_p["fiscal_year"].dropna().astype(int).tolist())
        )
        sel_fy = st.multiselect("会計年度 (FY)", all_fy, default=all_fy)

        all_megas = [m for m in MEGA_ORDER if m in (
            set(df_c["mega_category"].dropna()) | set(df_p["mega_category"].dropna())
        )]
        sel_megas = st.multiselect("大カテゴリ（モノの種類）", all_megas, default=all_megas)

        WORK_TYPE_ORDER = ["製造・納入", "工事・設置", "保守・運用", "設計・調査", "その他"]
        all_wts = [w for w in WORK_TYPE_ORDER if w in set(df_c["work_type"].dropna())]
        sel_wts = st.multiselect("作業種別（work_type）", all_wts, default=all_wts)

        all_cats = sorted(
            (set(df_c["category"].dropna()) | set(df_p["category"].dropna()))
            - {"不明"}
        ) + ["不明"]
        sel_cats = st.multiselect("カテゴリ（詳細）", all_cats, default=all_cats)

        all_areas = sorted(
            set(df_c["area_type"].dropna()) | set(df_p["area_type"].dropna())
        )
        sel_areas = st.multiselect("エリアタイプ", all_areas, default=all_areas)

        all_systems = sorted(df_c["system_name"].dropna().unique())
        sel_systems = st.multiselect("システム名", all_systems, default=all_systems)

        all_orgs = sorted(df_c["organization"].dropna().unique())
        sel_orgs = st.multiselect(
            "組織（contracts）",
            all_orgs,
            default=all_orgs,
            help="航空局本局・TCAB・WCAB の組織名",
        )

        all_bureaux = sorted(df_p["bureau"].dropna().unique())
        sel_bureaux = st.multiselect(
            "地方整備局（PAS）",
            all_bureaux,
            default=all_bureaux,
        )

        bid_options = ["一般競争", "随意契約", "企画競争", "指名競争", "その他"]
        sel_bids = st.multiselect("入札方式", bid_options, default=bid_options)

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
        f"{fc['amount_oku'].sum():,.0f} 億円",
    )
    sub2.metric(
        "PAS（地方整備局）",
        f"{len(fp):,} 件",
        f"{fp['amount_oku'].sum():,.0f} 億円",
    )
    c_rate = (
        (combined["bid_type"] == "随意契約").sum() / total_count * 100
        if total_count > 0
        else 0
    )
    sub3.metric("随意契約率（件数）", f"{c_rate:.1f} %")

    st.divider()

    # ── Row 1: 大カテゴリ別金額（変更3） + FY トレンド ──────────
    col_a, col_b = st.columns(2)

    with col_a:
        st.markdown("#### 大カテゴリ別 金額（横棒）")
        if not combined.empty:
            mega_df = (
                combined.groupby("mega_category")["amount"]
                .sum()
                .reset_index()
            )
            mega_df["億円"] = mega_df["amount"] / 1e8
            mega_df = mega_df.sort_values("億円", ascending=True)
            fig_a = px.bar(
                mega_df,
                x="億円",
                y="mega_category",
                orientation="h",
                color="mega_category",
                color_discrete_map=MEGA_COLORS,
                labels={"mega_category": "大カテゴリ", "億円": "金額（億円）"},
                template=TEMPLATE,
            )
            fig_a.update_layout(
                showlegend=False,
                height=380,
                margin=dict(l=8, r=8, t=8, b=8),
                yaxis=dict(categoryorder="total ascending", dtick=1),
            )
            fig_a.update_traces(hovertemplate="%{y}<br>%{x:,.1f}億円<extra></extra>")
            ev_a = st.plotly_chart(fig_a, use_container_width=True, on_select="rerun", key="chart_mega")
            if ev_a and ev_a.selection and ev_a.selection.points:
                sel_mega = ev_a.selection.points[0].get("y")
                if sel_mega:
                    show_drilldown_dialog(
                        combined[combined["mega_category"] == sel_mega],
                        f"大カテゴリ: {sel_mega}",
                    )
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
            fig_b = px.bar(
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
            fig_b.update_traces(texttemplate="%{text:,.0f}", textposition="inside", hovertemplate="FY%{x}<br>%{y:,.1f}億円<extra></extra>")
            fig_b.update_layout(
                height=380,
                margin=dict(l=8, r=8, t=8, b=8),
                legend=dict(orientation="h", yanchor="bottom", y=1.01),
                xaxis=dict(type="category"),
            )
            ev_b = st.plotly_chart(fig_b, use_container_width=True, on_select="rerun", key="chart_fy")
            if ev_b and ev_b.selection and ev_b.selection.points:
                sel_fy_val = ev_b.selection.points[0].get("x")
                if sel_fy_val is not None:
                    try:
                        sel_fy_int = int(str(sel_fy_val))
                        show_drilldown_dialog(
                            combined[combined["fiscal_year"] == sel_fy_int],
                            f"FY{sel_fy_int}",
                        )
                    except (ValueError, TypeError):
                        pass
        else:
            st.info("データなし")
        st.caption("※ FY2019のPASデータはアーカイブ欠損のため未収録")

    # ── Row 2: システム名（変更4・5） + ベンダー上位 ─────────────
    col_c, col_d = st.columns(2)

    with col_c:
        st.markdown("#### システム名別 金額")
        sys_data = fc[fc["system_name"].notna()]
        if not sys_data.empty:
            sys_df = (
                sys_data.groupby("system_name")["contract_amount"]
                .sum()
                .reset_index()
            )
            sys_df["億円"] = sys_df["contract_amount"] / 1e8
            sys_df = sys_df.sort_values("億円", ascending=True)
            fig_c = px.bar(
                sys_df,
                x="億円",
                y="system_name",
                orientation="h",
                color="億円",
                color_continuous_scale="Teal",
                labels={"system_name": "システム", "億円": "金額（億円）"},
                template=TEMPLATE,
            )
            # 変更4: 全ラベル強制表示
            fig_c.update_layout(
                coloraxis_showscale=False,
                height=480,
                margin=dict(l=8, r=8, t=8, b=8),
                yaxis=dict(dtick=1),
            )
            fig_c.update_traces(hovertemplate="%{y}<br>%{x:,.1f}億円<extra></extra>")
            ev_c = st.plotly_chart(fig_c, use_container_width=True, on_select="rerun", key="chart_system")
            if ev_c and ev_c.selection and ev_c.selection.points:
                sel_sys = ev_c.selection.points[0].get("y")
                if sel_sys:
                    show_drilldown_dialog(
                        combined[combined["system_name"] == sel_sys],
                        f"システム: {sel_sys}",
                    )

            # 変更5: システム名 → 参照先リンクテーブル（PDF or CIDAI）
            existing_systems = sorted(sys_data["system_name"].unique())
            ref_rows = []
            for sname in existing_systems:
                ref = SYSTEM_REFERENCES.get(sname)
                if ref and ref[0] == "PDF":
                    page = ref[1]
                    ref_rows.append({
                        "システム名": sname,
                        "参照先": "PDF",
                        "ページ": str(page),
                        "開く": f"{PDF_BASE}#page={page}&view=Fit",
                    })
                elif ref and ref[0] == "URL":
                    ref_rows.append({
                        "システム名": sname,
                        "参照先": "CIDAI",
                        "ページ": "—",
                        "開く": ref[1],
                    })
                else:
                    ref_rows.append({
                        "システム名": sname,
                        "参照先": "—",
                        "ページ": "—",
                        "開く": "",
                    })
            if ref_rows:
                ref_df = pd.DataFrame(ref_rows)
                st.caption("▼ システム参照先（航空保安業務の概要PDF / CIDAIサイト）")
                st.dataframe(
                    ref_df,
                    use_container_width=True,
                    height=min(38 * len(ref_rows) + 38, 280),
                    column_config={
                        "開く": st.column_config.LinkColumn("開く", display_text="開く"),
                    },
                    hide_index=True,
                )
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
            fig_d = px.bar(
                v_df,
                x="億円",
                y="表示名",
                orientation="h",
                color="億円",
                color_continuous_scale="Oranges",
                labels={"表示名": "ベンダー", "億円": "金額（億円）"},
                template=TEMPLATE,
            )
            fig_d.update_layout(
                coloraxis_showscale=False,
                height=520,
                margin=dict(l=8, r=8, t=8, b=8),
                yaxis=dict(dtick=1),
            )
            fig_d.update_traces(hovertemplate="%{y}<br>%{x:,.1f}億円<extra></extra>")
            ev_d = st.plotly_chart(fig_d, use_container_width=True, on_select="rerun", key="chart_vendor")
            if ev_d and ev_d.selection and ev_d.selection.points:
                sel_display = ev_d.selection.points[0].get("y")
                if sel_display:
                    matched = v_df[v_df["表示名"] == sel_display]["vendor_normalized"].values
                    if len(matched) > 0:
                        show_drilldown_dialog(
                            combined[combined["vendor_normalized"] == matched[0]],
                            f"ベンダー: {matched[0]}",
                        )
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
            fig_e = px.pie(
                org_df,
                values="億円",
                names="source_org",
                hole=0.5,
                template=TEMPLATE,
                color_discrete_sequence=px.colors.qualitative.Set2,
            )
            fig_e.update_traces(
                textinfo="percent+label",
                hovertemplate="%{label}<br>%{value:,.1f}億円<br>%{percent}",
            )
            fig_e.update_layout(
                height=420,
                margin=dict(l=8, r=8, t=30, b=8),
                showlegend=True,
                legend=dict(orientation="v", x=1.0, y=0.5),
            )
            ev_e = st.plotly_chart(fig_e, use_container_width=True, on_select="rerun", key="chart_org_pie")
            if ev_e and ev_e.selection and ev_e.selection.points:
                sel_org = ev_e.selection.points[0].get("label")
                if sel_org:
                    show_drilldown_dialog(
                        combined[combined["source_org"] == sel_org],
                        f"組織: {sel_org}",
                    )
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
            fig_f = make_subplots(specs=[[{"secondary_y": True}]])
            fig_f.add_trace(
                go.Bar(
                    x=m_df["year_month"],
                    y=m_df["件数"],
                    name="件数",
                    marker_color="#4472c4",
                    opacity=0.8,
                    hovertemplate="%{x}<br>%{y:,}件<extra></extra>",
                ),
                secondary_y=False,
            )
            fig_f.add_trace(
                go.Scatter(
                    x=m_df["year_month"],
                    y=m_df["合計億円"],
                    name="金額（億円）",
                    mode="lines+markers",
                    line=dict(color="#ffc000", width=2),
                    marker=dict(size=4),
                    hovertemplate="%{x}<br>%{y:,.1f}億円<extra></extra>",
                ),
                secondary_y=True,
            )
            fig_f.update_layout(
                template=TEMPLATE,
                height=420,
                margin=dict(l=8, r=8, t=8, b=8),
                legend=dict(orientation="h", yanchor="bottom", y=1.01),
                xaxis=dict(tickangle=-45),
            )
            fig_f.update_yaxes(title_text="件数", secondary_y=False)
            fig_f.update_yaxes(title_text="金額（億円）", secondary_y=True)
            ev_f = st.plotly_chart(fig_f, use_container_width=True, on_select="rerun", key="chart_monthly")
            if ev_f and ev_f.selection and ev_f.selection.points:
                sel_month = ev_f.selection.points[0].get("x")
                if sel_month:
                    show_drilldown_dialog(
                        combined[combined["year_month"] == str(sel_month)],
                        f"月: {sel_month}",
                    )
        else:
            st.info("日付データなし")

    # ── 大カテゴリ FY推移（積み上げ） ──────────────────────────
    st.markdown("#### 大カテゴリ別 FY推移（積み上げ棒）")
    if not combined.empty:
        mg_fy = (
            combined.groupby(["fiscal_year", "mega_category"])["amount"]
            .sum()
            .reset_index()
        )
        mg_fy["億円"] = mg_fy["amount"] / 1e8
        fig_mg_fy = px.bar(
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
        fig_mg_fy.update_layout(
            height=380,
            margin=dict(l=8, r=8, t=8, b=8),
            legend=dict(orientation="h", yanchor="bottom", y=1.01, font=dict(size=9)),
            xaxis=dict(type="category"),
        )
        fig_mg_fy.update_traces(hovertemplate="FY%{x}<br>%{y:,.1f}億円<extra></extra>")
        ev_mg = st.plotly_chart(fig_mg_fy, use_container_width=True, on_select="rerun", key="chart_mega_fy")
        if ev_mg and ev_mg.selection and ev_mg.selection.points:
            pt = ev_mg.selection.points[0]
            sel_fy_mg = pt.get("x")
            if sel_fy_mg is not None:
                try:
                    sel_fy_int_mg = int(str(sel_fy_mg))
                    filtered_mg = combined[combined["fiscal_year"] == sel_fy_int_mg]
                    show_drilldown_dialog(filtered_mg, f"FY{sel_fy_int_mg}")
                except (ValueError, TypeError):
                    pass

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
            fig_wt.update_traces(texttemplate="%{text}件", textposition="outside", hovertemplate="%{x}<br>%{y:,.1f}億円<extra></extra>")
            fig_wt.update_layout(
                showlegend=False, height=340,
                margin=dict(l=8, r=8, t=8, b=8),
            )
            st.plotly_chart(fig_wt, use_container_width=True)

        with wt_col2:
            heat = (
                fc.groupby(["mega_category", "work_type"])["amount_oku"]
                .sum()
                .reset_index()
                .pivot(index="mega_category", columns="work_type", values="amount_oku")
                .fillna(0)
                .reindex(columns=WORK_ORDER, fill_value=0)
                .reindex(MEGA_ORDER[::-1], fill_value=0)
            )
            fig_heat = go.Figure(go.Heatmap(
                z=heat.values,
                x=heat.columns.tolist(),
                y=heat.index.tolist(),
                colorscale="Blues",
                text=[[f"{v:,.0f}億" for v in row] for row in heat.values],
                texttemplate="%{text}",
                hovertemplate="%{y} × %{x}<br>%{z:,.1f}億円<extra></extra>",
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
        fig_org = px.bar(
            org_bar,
            x="organization",
            y="億円",
            color="億円",
            color_continuous_scale="Viridis",
            text="件数",
            labels={"organization": "組織", "億円": "金額（億円）"},
            template=TEMPLATE,
        )
        fig_org.update_traces(texttemplate="%{text}件", textposition="outside", hovertemplate="%{x}<br>%{y:,.1f}億円<extra></extra>")
        fig_org.update_layout(
            coloraxis_showscale=False,
            height=380,
            margin=dict(l=8, r=8, t=8, b=8),
            xaxis=dict(tickangle=-30),
        )
        ev_org = st.plotly_chart(fig_org, use_container_width=True, on_select="rerun", key="chart_org_bar")
        if ev_org and ev_org.selection and ev_org.selection.points:
            sel_org_name = ev_org.selection.points[0].get("x")
            if sel_org_name:
                show_drilldown_dialog(
                    combined[combined["organization"] == sel_org_name],
                    f"組織: {sel_org_name}",
                )

    # ── 予算比較（FY2013-2025 全期間トレンド） ────────────────────
    st.markdown("#### 📈 予算比較 — 調達母数 vs DB収録額（FY2013〜2025 全期間）")
    st.caption(
        "調達母数 = 予算歳出合計 − 人件費（推定） − 借入金償還 − "
        "非調達的経費（市町村交付金・土地借料・ICAO分担金等）"
    )

    # FY別調達母数（全年度の調達母数は予算書PDFから算出）
    # 空港整備勘定歳出合計から人件費・財投借入金償還・非調達的経費を除いた額
    _PROC_BASE_OKU = {
        2013: 1010, 2014: 1510, 2015: 1590, 2016: 1895,
        2017: 1975, 2018: 2320, 2019: 2430,
        2020: 2152, 2021: 2146, 2022: 2137, 2023: 2147, 2024: 2407,
        2025: 2500,
    }
    _PROC_BASE_CONFIRMED = set(_PROC_BASE_OKU.keys())

    _all_bgt = load_all_years_budget()

    # DB実績：全データ（フィルタ適用前）から FY別集計
    _c_fy = df_c.groupby("fiscal_year")["contract_amount"].sum() / 1e8
    _p_fy = df_p.groupby("fiscal_year")["award_amount"].sum() / 1e8
    _db_fy = _c_fy.add(_p_fy, fill_value=0)

    bgt_rows2 = []
    for fy in range(2013, 2026):
        fy_key = f"FY{fy}"
        bgt_total = _all_bgt.get(fy_key, {}).get("budget_total", None)
        if bgt_total is None:
            continue
        proc_base = _PROC_BASE_OKU.get(fy, max(bgt_total - 1750, 0))
        non_proc = max(bgt_total - proc_base, 0)
        db_oku = float(_db_fy.get(fy, 0))
        confirmed = fy in _PROC_BASE_CONFIRMED
        bgt_rows2.append({
            "FY": fy,
            "調達母数": proc_base,
            "調達母数外": non_proc,
            "DB収録額": db_oku,
            "確定": confirmed,
        })
    bgt_df2 = pd.DataFrame(bgt_rows2)

    fig2 = go.Figure()
    fig2.add_trace(go.Bar(
        name="調達母数（下段・濃色）",
        x=bgt_df2["FY"],
        y=bgt_df2["調達母数"],
        marker_color="#4472c4",
        opacity=0.85,
        hovertemplate="%{x}<br>調達母数: %{y:,.0f}億円<extra></extra>",
    ))
    fig2.add_trace(go.Bar(
        name="調達母数外（上段・薄色）",
        x=bgt_df2["FY"],
        y=bgt_df2["調達母数外"],
        marker_color="#b4c7e7",
        opacity=0.55,
        hovertemplate="%{x}<br>調達母数外: %{y:,.0f}億円<extra></extra>",
    ))
    fig2.add_trace(go.Scatter(
        name="DB収録額（折れ線）",
        x=bgt_df2["FY"],
        y=bgt_df2["DB収録額"],
        mode="lines+markers+text",
        line=dict(color="#ffc000", width=2),
        marker=dict(size=7),
        text=[f"{v:,.0f}" for v in bgt_df2["DB収録額"]],
        textposition="top center",
        textfont=dict(size=9),
        hovertemplate="%{x}<br>DB収録額: %{y:,.1f}億円<extra></extra>",
    ))
    fig2.update_layout(
        template=TEMPLATE,
        barmode="stack",
        height=460,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, font=dict(size=10)),
        margin=dict(l=8, r=8, t=8, b=8),
        xaxis=dict(type="category", title="会計年度"),
        yaxis=dict(title="金額（億円）"),
    )
    st.plotly_chart(fig2, use_container_width=True)

    st.caption(
        "※ FY2020（令和2年度）は新型コロナ対策補正予算により歳出が急増。"
        "当初予算 3,836億円 → 補正後 6,623億円（差分 2,787億円が補正）。"
        "補正分の主な内訳: 空港における感染症対策（サーモグラフィ・消毒設備等）、"
        "航空会社支援（旅行需要急減対策）、那覇空港第二滑走路整備前倒し等。"
        "これらは物品役務調達ではなく補助金・支援金が大半のため調達母数（2,152億）はほぼ不変。"
    )
    st.caption(
        "※ 予算は歳出予算額ベース、DB収録額は契約締結額ベース。"
        "多年度契約を含むため直接比較はオーダー感の把握用。"
    )
    st.caption(
        "※ 全年度の調達母数は予算書PDFから算出。"
        "空港整備勘定歳出合計から人件費・財投借入金償還・非調達的経費を除いた額。"
    )

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
        csv_c = show_c.to_csv(index=False, encoding="utf-8-sig")
        show_c_disp = show_c.copy()
        show_c_disp["金額（円）"] = show_c_disp["金額（円）"].apply(lambda x: f"{x:,.0f}" if pd.notna(x) else "")
        st.dataframe(show_c_disp, use_container_width=True, height=320)
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
        csv_p = show_p.to_csv(index=False, encoding="utf-8-sig")
        show_p_disp = show_p.copy()
        show_p_disp["落札額（円）"] = show_p_disp["落札額（円）"].apply(lambda x: f"{x:,.0f}" if pd.notna(x) else "")
        st.dataframe(show_p_disp, use_container_width=True, height=320)
        st.download_button(
            "📥 PAS CSV ダウンロード",
            csv_p,
            file_name="pas_filtered.csv",
            mime="text/csv",
        )

    # ── 情報ソース一覧 ─────────────────────────────────────────
    _all_bgt_src = load_all_years_budget()
    show_source_links(_all_bgt_src)


def show_source_links(all_years_budget: dict):
    """情報ソース一覧 expander"""
    with st.expander("情報ソース一覧", expanded=False):
        st.markdown(
            "<style>.source-links a { font-size: 0.82rem; }</style>",
            unsafe_allow_html=True,
        )
        st.markdown("##### 適正化データ（契約実績）")
        st.markdown(
            "- **航空局本局**: [koku_tk1_000033.html](https://www.mlit.go.jp/koku/koku_tk1_000033.html)"
            "　｜　WARP(H25-R3): [warp.ndl.go.jp](https://warp.ndl.go.jp/web/20241003183110/https://www.mlit.go.jp:8088/koku/koku_tk1_000033.html)\n"
            "- **東京航空局(TCAB)**: [tcab/contract/](https://www.cab.mlit.go.jp/tcab/contract/)"
            "　｜　WARP(H24-R2): [warp.ndl.go.jp](https://warp.ndl.go.jp/web/20201001151609/https://www.cab.mlit.go.jp/tcab/contract/contract_04/post_271.html)\n"
            "- **大阪航空局(WCAB)**: [wcab/contract/published.html](https://www.cab.mlit.go.jp/wcab/contract/published.html)"
            "　｜　WARP(H27-R4): [warp.ndl.go.jp](https://warp.ndl.go.jp/web/20230601134447/https://www.cab.mlit.go.jp/wcab/contract/published.html)\n"
            "- **地方整備局PAS**: [pas.ysk.nilim.go.jp](https://www.pas.ysk.nilim.go.jp/)"
        )

        st.markdown("##### 予算概要（航空保安勘定 年度別PDF）")
        bgt_lines = []
        for fy_key in sorted(all_years_budget.keys()):
            entry = all_years_budget[fy_key]
            src = entry.get("source", "")
            nen = entry.get("nen_go", "")
            note = entry.get("note", "")
            total = entry.get("budget_total", "")
            if src:
                label = f"{fy_key}（{nen}）{total}億 {note}"
                bgt_lines.append(f"- [{label}]({src})")
        st.markdown("\n".join(bgt_lines))

        st.markdown("##### その他")
        st.markdown(
            "- [航空保安業務の概要（2025年版）](https://www.mlit.go.jp/koku/content/001743241.pdf)\n"
            "- [GEPSポータル](https://www.p-portal.go.jp/)\n"
            "- [行政事業レビューシステム](https://rssystem.go.jp/)\n"
            "- [WARP（国立国会図書館ウェブアーカイブ）](https://warp.ndl.go.jp/)"
        )


if __name__ == "__main__":
    main()
