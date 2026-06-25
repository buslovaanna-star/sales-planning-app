import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import warnings
import json
import re
import requests
warnings.filterwarnings('ignore')

st.set_page_config(
    page_title="Планування продажів",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
.rec-card { border-radius:8px; padding:12px 16px; margin-bottom:8px; }
.rec-low  { background:#e8f5e9; color:#1b5e20; }
.rec-high { background:#fff3e0; color:#e65100; }
.rec-new  { background:#e3f2fd; color:#0d47a1; }
</style>
""", unsafe_allow_html=True)

# ── Session state ─────────────────────────────────────────────────────────────
for k, v in [
    ("df_fact", None), ("df_plan_auto", None), ("df_plan_edited", None),
    ("seasonality", None), ("store_starts", {}), ("plan_month", None),
    ("excel_plans", {}),
    ("bias_corrections", {}),
    ("promo_calendar", pd.DataFrame()),
    ("promo_static_loaded", False),
]:
    if k not in st.session_state:
        st.session_state[k] = v




# ── Core helpers ──────────────────────────────────────────────────────────────

def parse_seasonality(xl):
    seas_raw = xl.get("Коеф. Сезонності")
    seasonality = {}
    if seas_raw is None:
        return seasonality
    for _, row in seas_raw.iterrows():
        try:
            m = int(float(row.iloc[14]))
            c = float(row.iloc[15])
            if 1 <= m <= 12:
                seasonality[m] = c
        except Exception:
            pass
    return seasonality


def parse_fact_sheet(xl):
    raw = xl.get("ДАНІ ВНОСИТИ")
    if raw is None:
        st.error("Не знайдено лист 'ДАНІ ВНОСИТИ'")
        return None, {}, {}

    stores = list(raw.iloc[0, 2:].dropna())
    dates_raw = raw.iloc[3:, 0].reset_index(drop=True)
    data_block = raw.iloc[3:, 2:2+len(stores)].reset_index(drop=True)
    data_block.columns = stores
    df = data_block.copy()
    df.insert(0, "date", pd.to_datetime(dates_raw, errors="coerce"))
    df = df[df["date"].notna()].set_index("date")
    df = df.apply(pd.to_numeric, errors="coerce")

    store_starts = {}
    for s in stores:
        nonzero = df[s].replace(0, np.nan).dropna()
        store_starts[s] = nonzero.index[0] if len(nonzero) > 0 else None

    seasonality = parse_seasonality(xl)
    return df, store_starts, seasonality


def parse_excel_plans(xl: dict) -> dict:
    """
    Extract store-level plan values from any available calc sheet.
    Priority: розрахунки >рік+доповн > розрахунки >заг метрика похибки > розрахунки <рік
    The 'заг метрика' sheet has 5 cols per store: fact, plan, deviation, APE, abs_dev
    Returns dict: (store_name, month_timestamp) -> plan_value
    """
    plans = {}

    # Sheet configs: (sheet_name, col_step, plan_col_offset, skip_keywords)
    sheet_configs = [
        ("розрахунки >рік+доповн",       5, 1, ["ПЛАН", "виконання"]),
        ("розрахунки >заг метрика похибки", 5, 1, ["ПЛАН", "відхилення", "APE", "модуль"]),
        ("розрахунки <рік",              5, 1, ["ПЛАН", "виконання"]),
    ]

    for sheet_name, step, plan_offset, skip_kw in sheet_configs:
        df = xl.get(sheet_name)
        if df is None:
            continue
        dates = pd.to_datetime(df.iloc[1:, 0], errors="coerce").reset_index(drop=True)
        for c in range(1, df.shape[1], step):
            store_raw = str(df.iloc[0, c])
            if any(kw in store_raw for kw in skip_kw) or store_raw == "nan":
                continue
            store_name = store_raw.strip()
            plan_col = c + plan_offset
            if plan_col >= df.shape[1]:
                continue
            for i, d in enumerate(dates):
                if pd.isna(d):
                    continue
                key = (store_name, d.to_period("M").to_timestamp())
                if key in plans:   # already set by higher-priority sheet
                    continue
                try:
                    v = float(df.iloc[i + 1, plan_col])
                    if v > 0:
                        plans[key] = v
                except Exception:
                    pass
    return plans


# ── Promo calendar helpers ────────────────────────────────────────────────────

# Weights matched by substring — order matters (more specific first)
PROMO_TYPE_WEIGHT = {
    "ТОВАР МІСЯЦЯ":           3.0,   # ТОВАР МІСЯЦЯ / ТОВАР МІСЯЦЯ МАГАЗИНИ / САЙТ
    "1+1=3":                  2.5,   # 1+1=3 та товари / 2=3 та товари
    "2=3":                    2.5,
    "-50":                    2.5,   # -50 на другу одиницю
    "чорна п'ятниця":         3.0,   # Black Friday
    "ніч знижок":             2.0,
    "25%":                    2.0,   # глибокі знижки ≥25%
    "40%":                    2.5,
    "промокод":               1.5,
    "знижка":                 1.0,
    "акція":                  1.0,
    "день":                   0.8,   # День матері, День Києва тощо
    "тиждень":                1.2,   # MY NUTRI WEEK
    "розсилка":               0.3,   # не впливає на продажі магазину
    "банер":                  0.3,
    "промо":                  0.5,
}
PROMO_DEEP_TYPES = {"ТОВАР МІСЯЦЯ", "1+1=3", "2=3", "-50",
                    "чорна п'ятниця", "40%"}

# ── Hardcoded config (set once, never needs UI) ───────────────────────────────
# Файл з акціями за 2025 рік у репозиторії (відносний шлях від app.py)
# Залиште порожнім "" якщо файлу немає
STATIC_PROMO_FILE = "promo_2025.xlsx"   # або "promo_2025.csv"

# Посилання на Google Таблицю з акціями 2026 (відкрита для перегляду)
# Залиште порожнім "" щоб вводити вручну у вкладці Промокалендар
GSHEET_PROMO_URL = "https://docs.google.com/spreadsheets/d/1D1wI3WF3sc8zaJLU0TqD_wENhTRRWeLOSf0JKTbC8co/edit?pli=1&gid=0#gid=0"   # вставте сюди посилання, напр.: "https://docs.google.com/spreadsheets/d/..."
# ─────────────────────────────────────────────────────────────────────────────


# ── Auto-load promo calendar from repo/config (runs once per session) ─────────
import os as _os
if not st.session_state.get("_promo_autoloaded", False):
    st.session_state["_promo_autoloaded"] = True
    _frames = []

    # 1. Static file from repository
    if STATIC_PROMO_FILE:
        _path = _os.path.join(_os.path.dirname(__file__), STATIC_PROMO_FILE)
        if _os.path.exists(_path):
            try:
                _raw = (pd.read_excel(_path, header=None) if _path.endswith(".xlsx")
                        else pd.read_csv(_path, header=None))
                _frames.append(parse_promo_df(_raw))
            except Exception:
                pass

    # 2. Google Sheets from config
    if GSHEET_PROMO_URL:
        try:
            _frames.append(load_promo_from_gsheet(GSHEET_PROMO_URL))
        except Exception:
            pass

    if _frames:
        _combined = pd.concat(_frames, ignore_index=True).drop_duplicates(
            subset=["date_start", "date_end", "promo_type"]
        ).reset_index(drop=True)
        st.session_state.promo_calendar       = _combined
        st.session_state.promo_static_loaded  = True
# ──────────────────────────────────────────────────────────────────────────────


def parse_promo_df(df_raw: pd.DataFrame) -> pd.DataFrame:
    """
    Normalize a raw promo DataFrame to columns:
    date_start, date_end, promo_type, name
    Accepts the exact layout: cols A=date_start, B=date_end, C=type, D=sku, E=name
    Row 0 may be a header or data.
    """
    df = df_raw.copy()

    # Detect header row: first row where col 0 is text (not a date)
    header_row = None
    for i in range(min(5, len(df))):
        val = df.iloc[i, 0]
        if isinstance(val, str) and len(val) > 2:
            try:
                pd.to_datetime(val, dayfirst=True)
            except Exception:
                header_row = i
                break

    if header_row is not None:
        df.columns = [str(c).strip() for c in df.iloc[header_row].values]
        df = df.iloc[header_row + 1:].reset_index(drop=True)

    cols = df.columns.tolist()

    def find_col(keywords):
        for kw in keywords:
            for c in cols:
                if kw.lower() in str(c).lower():
                    return c
        return None

    col_start = find_col(["поча", "start"]) or cols[0]
    col_end   = find_col(["закін", "end"])   or cols[1]
    col_type  = find_col(["тип", "type"])    or cols[2]
    col_name  = find_col(["назва", "name"])

    result = pd.DataFrame()
    result["date_start"] = pd.to_datetime(df[col_start], errors="coerce", dayfirst=True)
    result["date_end"]   = pd.to_datetime(df[col_end],   errors="coerce", dayfirst=True)
    result["promo_type"] = df[col_type].astype(str).str.strip()
    result["name"]       = df[col_name].astype(str).str.strip() if col_name else ""

    # Fill missing date_end with date_start (one-day events)
    result["date_end"] = result["date_end"].fillna(result["date_start"])

    # Drop rows where we have no date at all, or type is empty/nan
    result = result[
        result["date_start"].notna() &
        ~result["promo_type"].isin(["nan", "", "None"])
    ].reset_index(drop=True)
    return result


def load_promo_from_gsheet(url: str) -> pd.DataFrame:
    """
    Load promo calendar from a Google Sheets share link.
    Accepts both /edit and /pub URLs.
    """
    # Extract sheet ID
    match = re.search(r'/spreadsheets/d/([a-zA-Z0-9_-]+)', url)
    if not match:
        raise ValueError("Не вдалося розпізнати посилання на Google Sheets")
    sheet_id = match.group(1)

    # Try to extract gid (sheet tab id)
    gid_match = re.search(r'gid=(\d+)', url)
    gid = gid_match.group(1) if gid_match else "0"

    csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
    resp = requests.get(csv_url, timeout=15)
    resp.raise_for_status()
    from io import StringIO
    df_raw = pd.read_csv(StringIO(resp.text), header=None)
    return parse_promo_df(df_raw)


def compute_monthly_promo_score(promo_df: pd.DataFrame, target_month: pd.Timestamp) -> dict:
    """
    For a given month, compute:
    - intensity:      weighted promo coverage (sum of days×weight / days_in_month)
    - has_deep:       bool — any deep promo (ТОВАР МІСЯЦЯ / 1+1=3 / -50)
    - has_anchor:     bool — ТОВАР МІСЯЦЯ -40% present
    - deep_days:      number of days with deep promo active
    - lift_factor:    multiplier to apply to base forecast (calibrated from data)
    """
    m_start = target_month
    m_end   = target_month + pd.offsets.MonthEnd(1)
    days_in_month = m_end.day

    # Filter promos overlapping this month
    overlap = promo_df[
        (promo_df["date_start"] <= m_end) &
        (promo_df["date_end"]   >= m_start)
    ].copy()

    if len(overlap) == 0:
        return {"intensity": 0, "has_deep": False, "has_anchor": False,
                "deep_days": 0, "lift_factor": 1.0, "n_promos": 0}

    # Compute active days per promo (clipped to month)
    clipped_end   = overlap["date_end"].clip(upper=m_end)
    clipped_start = overlap["date_start"].clip(lower=m_start)
    overlap["active_days"] = ((clipped_end - clipped_start).dt.days + 1).clip(lower=0)

    # Map weights
    def _weight(t):
        t_lo = t.lower()
        for k, v in PROMO_TYPE_WEIGHT.items():
            if k.lower() in t_lo:
                return v
        return 0.5  # unknown type → minimal weight

    def _is_deep(t):
        t_lo = t.lower()
        return any(k.lower() in t_lo for k in PROMO_DEEP_TYPES)

    overlap["weight"]    = overlap["promo_type"].map(_weight)
    overlap["is_deep"]   = overlap["promo_type"].map(_is_deep)
    overlap["is_anchor"] = overlap["promo_type"].str.upper().str.contains("ТОВАР МІСЯЦЯ", na=False)

    intensity  = (overlap["active_days"] * overlap["weight"]).sum() / days_in_month
    has_deep   = overlap["is_deep"].any()
    has_anchor = overlap["is_anchor"].any()
    deep_days  = int((overlap.loc[overlap["is_deep"], "active_days"]).sum())

    # Lift factor: calibrated from analysis
    # Base: 1.0 (no adjustment)
    # + anchor promo:      +0.08 (ТОВАР МІСЯЦЯ always present, partially in baseline)
    # + deep mechanic:     +0.06 additional
    # Soft promos (intensity only): marginal
    lift = 1.0
    if has_anchor:
        lift += 0.08
    if has_deep and not has_anchor:
        lift += 0.05
    elif has_deep and has_anchor:
        lift += 0.06

    return {
        "intensity": round(intensity, 2),
        "has_deep":  bool(has_deep),
        "has_anchor": bool(has_anchor),
        "deep_days": deep_days,
        "lift_factor": round(lift, 3),
        "n_promos": len(overlap),
    }


def deseasonalise(series: pd.Series, seasonality: dict) -> pd.Series:
    if not seasonality:
        return series
    return series / series.index.map(lambda d: seasonality.get(d.month, 1.0))


def compute_plan_store(series: pd.Series, target: pd.Timestamp,
                       seasonality: dict, store_start,
                       growth_old: float, growth_young: float,
                       age_threshold: int,
                       short_weight: float = 0.0,
                       short_window: int = 6) -> float:
    """
    Mature store (>= age_threshold months):
        Blend of long-term trend (all data) and short-term trend (last short_window months).
        short_weight=0 -> pure long-term; short_weight=1 -> pure short-term.
    Young store (< age_threshold):
        Last 12 months only, growth_young applied.
    """
    if store_start is None:
        return 0.0

    age = (target.year - store_start.year) * 12 + (target.month - store_start.month)
    growth = growth_old if age >= age_threshold else growth_young

    hist = series.replace(0, np.nan).dropna()
    hist = hist[hist.index < target]
    if len(hist) == 0:
        return 0.0

    deseas = deseasonalise(hist, seasonality) if seasonality else hist.copy()
    seas_coef = seasonality.get(target.month, 1.0) if seasonality else 1.0

    def trend_projection(window):
        n = len(window)
        if n >= 3:
            x = np.arange(n, dtype=float)
            slope, intercept = np.polyfit(x, window.values.astype(float), 1)
            return max(slope * n + intercept, 0.0)
        elif n > 0:
            return float(window.mean())
        return 0.0

    if age >= age_threshold:
        proj_long  = trend_projection(deseas)
        short_data = deseas.iloc[-short_window:] if len(deseas) >= short_window else deseas
        proj_short = trend_projection(short_data)
        projected  = (1 - short_weight) * proj_long + short_weight * proj_short
    else:
        projected = trend_projection(deseas.iloc[-12:])

    return round(max(projected, 0) * seas_coef * (1 + growth), 2)


def compute_auto_plan(df, store_starts, seasonality, target_month,
                      growth_old=0.0, growth_young=0.15, age_threshold=12,
                      short_weight=0.0, short_window=6):
    return pd.Series({
        store: compute_plan_store(
            df[store], target_month, seasonality,
            store_starts.get(store), growth_old, growth_young, age_threshold,
            short_weight, short_window
        )
        for store in df.columns
    })


def compute_bias_corrections(df: pd.DataFrame, store_starts: dict,
                              seasonality: dict, target_month: pd.Timestamp,
                              excel_plans: dict,
                              lookback_months: int = 12,
                              short_weight: float = 0.5,
                              short_window: int = 6) -> dict:
    """
    Compute MPE per store. Rules:
    - Only mature stores (age >= 12 months at target_month)
    - Only months where plan_v >= 1000 (skip near-zero bootstrap plans)
    - Require >= 3 valid observations
    - Cap correction at +-25%
    """
    MAX_CORR = 0.25
    MIN_PLAN = 1000

    has_data = df.replace(0, np.nan).dropna(how="all")
    avail = [m for m in has_data.index if m < target_month]
    months = avail[-lookback_months:] if len(avail) >= lookback_months else avail

    corrections = {}
    for store in df.columns:
        start = store_starts.get(store)
        if start is None:
            continue
        age_at_target = (target_month.year - start.year) * 12 + (target_month.month - start.month)
        if age_at_target < 12:
            continue
        mpes = []
        for month in months:
            fact_v = df.loc[month, store] if month in df.index else 0
            if pd.isna(fact_v) or fact_v == 0:
                continue
            plan_v = excel_plans.get((store, month), 0)
            if plan_v == 0:
                hist = df[df.index < month]
                p = compute_auto_plan(hist, store_starts, seasonality, month,
                                      growth_old=0.0, growth_young=0.15,
                                      age_threshold=12,
                                      short_weight=short_weight,
                                      short_window=short_window)
                plan_v = p.get(store, 0)
            if plan_v >= MIN_PLAN:
                mpes.append((fact_v - plan_v) / plan_v)
        if len(mpes) >= 3:
            raw = float(np.mean(mpes))
            corrections[store] = max(-MAX_CORR, min(MAX_CORR, raw))
    return corrections


def execution_table(fact: pd.Series, plan: pd.Series) -> pd.DataFrame:
    df = pd.DataFrame({"Факт": fact, "План": plan})
    df["Виконання %"] = np.where(df["План"] > 0,
                                  (df["Факт"] / df["План"] * 100).round(1), np.nan)
    df["Відхилення"] = (df["Факт"] - df["План"]).round(2)
    df["Статус"] = df["Виконання %"].apply(
        lambda v: "✅" if pd.notna(v) and v >= 100
        else ("⚠️" if pd.notna(v) and v >= 90 else ("🔴" if pd.notna(v) else "—"))
    )
    return df


def color_exec(val):
    if pd.isna(val):
        return ""
    if val >= 100:
        return "background-color:#c8e6c9;color:#1b5e20"
    if val >= 90:
        return "background-color:#fff9c4;color:#f57f17"
    return "background-color:#ffcdd2;color:#b71c1c"


def to_excel(df_fact, plan_auto, plan_edited, plan_month):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        if df_fact is not None:
            df_fact.reset_index().to_excel(w, sheet_name="Факт", index=False)
        if plan_auto is not None:
            pd.DataFrame({
                "Магазин": plan_auto.index,
                "Авто-план": plan_auto.values,
                "Відред. план": plan_edited.values if plan_edited is not None else plan_auto.values,
                "Місяць": str(plan_month)[:7],
            }).to_excel(w, sheet_name="План", index=False)
    buf.seek(0)
    return buf


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.title("📊 Планування продажів")
    st.markdown("---")

    uploaded = st.file_uploader("Завантажте Excel-файл", type=["xlsx"],
                                help="Лист 'ДАНІ ВНОСИТИ'")
    if uploaded:
        xl = pd.read_excel(uploaded, sheet_name=None, header=None)
        df_fact, store_starts, seasonality = parse_fact_sheet(xl)
        if df_fact is not None:
            st.session_state.df_fact = df_fact
            st.session_state.store_starts = store_starts
            st.session_state.seasonality = seasonality
            st.session_state.excel_plans = parse_excel_plans(xl)
            st.success(f"✓ Завантажено {len(df_fact.columns)} магазинів")

    st.markdown("---")
    st.markdown("**Параметри плану**")

    next_month = (pd.Timestamp.today().to_period("M") + 1).to_timestamp()
    month_str = st.text_input("Місяць плану (РРРР-ММ)", value=next_month.strftime("%Y-%m"))
    try:
        plan_month = pd.Timestamp(month_str + "-01")
    except Exception:
        plan_month = next_month
    st.session_state.plan_month = plan_month

    growth_old   = st.slider("Приріст, зрілі магазини (%)",  -20, 50, 0,  1) / 100
    growth_young = st.slider("Приріст, нові магазини (%)",   -20, 50, 15, 1) / 100
    age_threshold = st.slider("Поріг «молодий магазин» (міс.)", 6, 24, 12, 1)

    st.markdown("---")
    st.markdown("**Короткостроковий тренд**")
    short_weight = st.slider(
        "Вага короткострокового тренду (%)",
        min_value=0, max_value=100, value=50, step=5,
        help="0% = тільки довгостроковий тренд (вся історія)\n"
             "50% = оптимально за backtesting (MAPE −39%)\n"
             "100% = тільки останні 6 місяців"
    ) / 100
    short_window = st.slider("Вікно короткого тренду (міс.)", 3, 12, 6, 1)

    st.markdown("---")
    st.markdown("**Корекція зміщення (bias)**")
    use_bias = st.toggle(
        "Застосувати bias correction",
        value=True,
        help="Коригує план на основі середнього MPE кожного магазину за останній рік.\n"
             "Якщо магазин стабільно недовиконує план на 10% → план знижується на 10%."
    )
    bias_lookback = st.slider("Період для розрахунку MPE (міс.)", 6, 24, 12, 1,
                              disabled=not use_bias)

    if st.button("🤖 Розрахувати план", type="primary", use_container_width=True):
        if st.session_state.df_fact is not None:
            plan = compute_auto_plan(
                st.session_state.df_fact, st.session_state.store_starts,
                st.session_state.seasonality, plan_month,
                growth_old, growth_young, age_threshold,
                short_weight, short_window,
            )
            # Apply promo lift factor if calendar is loaded
            promo_df_sidebar = st.session_state.promo_calendar
            if len(promo_df_sidebar) > 0:
                promo_score = compute_monthly_promo_score(promo_df_sidebar, plan_month)
                lift = promo_score["lift_factor"]
                if lift != 1.0:
                    plan = (plan * lift).round(2)
            # Compute and apply bias corrections
            if use_bias:
                corrections = compute_bias_corrections(
                    st.session_state.df_fact,
                    st.session_state.store_starts,
                    st.session_state.seasonality,
                    plan_month,
                    st.session_state.excel_plans,
                    lookback_months=bias_lookback,
                    short_weight=short_weight,
                    short_window=short_window,
                )
                st.session_state.bias_corrections = corrections
                corrected = {}
                for store in plan.index:
                    corr = corrections.get(store, 0.0)
                    corrected[store] = round(plan[store] * (1 + corr), 2)
                plan = pd.Series(corrected)
            else:
                st.session_state.bias_corrections = {}

            st.session_state.df_plan_auto   = plan
            st.session_state.df_plan_edited = plan.copy()
            st.success("✓ План розраховано")
        else:
            st.warning("Спочатку завантажте файл")


# ══════════════════════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════════════════════
tab1, tab2, tab3, tab4, tab5 = st.tabs(["📁 Дані", "✏️ Редагування плану", "📊 Аналітика", "💡 Рекомендації", "📅 Промокалендар"])


# ── TAB 1: DATA ───────────────────────────────────────────────────────────────
with tab1:
    st.subheader("Завантажені дані — факт продажів по магазинах")

    if st.session_state.df_fact is None:
        st.info("👈 Завантажте Excel-файл у боковій панелі")
    else:
        df   = st.session_state.df_fact
        plan = st.session_state.df_plan_auto
        stores = list(df.columns)

        # Only show months that have at least one non-zero value
        has_data = df.replace(0, np.nan).dropna(how="all")

        col1, col2, col3 = st.columns(3)
        col1.metric("Магазинів", len(stores))
        col2.metric("Місяців з даними", len(has_data))
        col3.metric("Останній місяць", has_data.index[-1].strftime("%B %Y"))

        st.markdown("---")
        c1, c2 = st.columns([3, 1])
        search   = c1.text_input("Пошук магазину", placeholder="Введіть назву...")
        show_n   = c2.selectbox("Показати місяців", [6, 12, 24, 99], index=1)

        filtered = [s for s in stores if search.lower() in s.lower()] if search else stores
        # Only rows with data
        disp = has_data[filtered].iloc[-show_n:].copy()
        disp.index = disp.index.strftime("%b %Y")
        st.dataframe(
            disp.style.format("{:,.0f}", na_rep="—").highlight_null(color="#fff3cd"),
            use_container_width=True, height=380
        )

        st.markdown("---")
        st.markdown("**Тренд продажів — факт та план по магазину**")
        sel = st.selectbox("Магазин для графіку", filtered, key="chart_store")

        chart_fact = has_data[[sel]].rename(columns={sel: "Факт"})

        # Add plan line if plan is calculated
        if plan is not None and sel in plan.index:
            plan_val = plan[sel]
            pm = st.session_state.plan_month
            plan_row = pd.DataFrame({"План": [plan_val]}, index=[pm])
            chart_plan = plan_row
        else:
            chart_plan = None

        if chart_plan is not None:
            combined = chart_fact.join(chart_plan, how="outer")
            st.line_chart(combined, use_container_width=True)
        else:
            st.line_chart(chart_fact, use_container_width=True)
            st.caption("Розрахуйте план для відображення планової лінії")


# ── TAB 2: PLAN EDITING ───────────────────────────────────────────────────────
with tab2:
    pm_label = st.session_state.plan_month.strftime("%B %Y") if st.session_state.plan_month else "—"
    st.subheader(f"Редагування плану на {pm_label}")

    if st.session_state.df_plan_auto is None:
        st.info("👈 Розрахуйте план у боковій панелі")
    else:
        plan_auto   = st.session_state.df_plan_auto
        plan_edited = st.session_state.df_plan_edited.copy()
        df          = st.session_state.df_fact
        pm          = st.session_state.plan_month

        # ── Previous month summary ──────────────────────────────────────────
        prev_month = pm - pd.offsets.MonthBegin(1)
        has_data   = df.replace(0, np.nan).dropna(how="all")
        months_with_data = has_data.index

        if prev_month in months_with_data:
            # compute what plan would have been for prev month
            plan_prev = compute_auto_plan(
                df, st.session_state.store_starts,
                st.session_state.seasonality, prev_month,
                growth_old=0.0, growth_young=0.15, age_threshold=12,
            )
            fact_prev = df.loc[prev_month].replace(0, np.nan)
            exec_prev = execution_table(fact_prev, plan_prev)
            exec_prev = exec_prev.dropna(subset=["Факт"])

            total_fact_prev = exec_prev["Факт"].sum()
            total_plan_prev = exec_prev["План"].sum()
            exec_pct_prev   = total_fact_prev / total_plan_prev * 100 if total_plan_prev > 0 else 0

            st.markdown(f"**Підсумок попереднього місяця — {prev_month.strftime('%B %Y')}**")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Факт (сума)", f"{total_fact_prev:,.0f}")
            c2.metric("План (сума)", f"{total_plan_prev:,.0f}")
            c3.metric("Виконання мережі", f"{exec_pct_prev:.1f}%",
                      delta=f"{exec_pct_prev-100:+.1f}%",
                      delta_color="normal" if exec_pct_prev >= 100 else "inverse")
            above = (exec_prev["Виконання %"] >= 100).sum()
            c4.metric("Виконали план", f"{above} / {len(exec_prev)}")
            st.markdown("---")

        # ── Global adjustment ───────────────────────────────────────────────
        st.markdown("**Глобальне коригування**")
        cg1, cg2, cg3 = st.columns(3)
        global_pct = cg1.number_input("Приріст до авто-плану (%)", -50.0, 100.0, 0.0, 0.5)
        if cg2.button("Застосувати до всіх"):
            st.session_state.df_plan_edited = (plan_auto * (1 + global_pct / 100)).round(2)
            st.rerun()
        if cg3.button("↩ Скинути до авто-плану"):
            st.session_state.df_plan_edited = plan_auto.copy()
            st.rerun()

        st.markdown("---")
        st.markdown("**Ручне редагування по кожному магазину**")

        corrections = st.session_state.bias_corrections
        has_bias = len(corrections) > 0

        # Build edit table — add bias column when corrections are present
        edit_data = {
            "Магазин": plan_auto.index,
            "Авто-план": plan_auto.values.round(2),
        }
        if has_bias:
            edit_data["Bias MPE %"] = [
                round(corrections.get(s, 0.0) * 100, 1) for s in plan_auto.index
            ]
        edit_data["Ваш план"] = st.session_state.df_plan_edited.values.round(2)

        edit_df = pd.DataFrame(edit_data).set_index("Магазин")

        col_config = {
            "Авто-план": st.column_config.NumberColumn(
                "Авто-план (до корекції)" if has_bias else "Авто-план",
                disabled=True, format="%.0f"
            ),
            "Ваш план": st.column_config.NumberColumn("Ваш план ✏️", format="%.0f", min_value=0),
        }
        if has_bias:
            col_config["Bias MPE %"] = st.column_config.NumberColumn(
                "Bias MPE %",
                disabled=True,
                format="%.1f%%",
                help="Середнє відхилення факту від плану за обраний період. "
                     "Негативне = план завищений, позитивне = план занижений."
            )

        if has_bias:
            n_corrected = sum(1 for v in corrections.values() if abs(v) >= 0.03)
            st.caption(
                f"✅ Bias correction активна — скориговано {n_corrected} магазинів "
                f"(|MPE| ≥ 3%). Колонка «Авто-план» показує план ДО корекції, "
                f"«Ваш план» — вже з поправкою."
            )

        edited = st.data_editor(
            edit_df, use_container_width=True,
            column_config=col_config,
            height=550,
        )
        st.session_state.df_plan_edited = pd.Series(edited["Ваш план"].values, index=edited.index)

        st.markdown("---")
        c1, c2, c3 = st.columns(3)
        c1.metric("Авто-план (сума)",  f"{plan_auto.sum():,.0f} USD")
        c2.metric("Ваш план (сума)",   f"{st.session_state.df_plan_edited.sum():,.0f} USD")
        delta = st.session_state.df_plan_edited.sum() - plan_auto.sum()
        c3.metric("Різниця", f"{delta:+,.0f} USD",
                  delta_color="normal" if delta >= 0 else "inverse")

        st.markdown("---")
        buf = to_excel(df, plan_auto, st.session_state.df_plan_edited, pm)
        st.download_button("⬇️ Завантажити план у Excel", data=buf,
                           file_name=f"план_{pm.strftime('%Y_%m')}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)


# ── TAB 3: ANALYTICS ─────────────────────────────────────────────────────────
with tab3:
    st.subheader("Аналітика виконання плану")

    if st.session_state.df_fact is None:
        st.info("👈 Завантажте Excel-файл у боковій панелі")
    elif st.session_state.df_plan_auto is None:
        st.info("👈 Спочатку розрахуйте план")
    else:
        df   = st.session_state.df_fact
        pm   = st.session_state.plan_month

        # Only months WITH actual data
        has_data = df.replace(0, np.nan).dropna(how="all")
        avail    = [m for m in has_data.index if m < pm]

        if not avail:
            st.info("Немає місяців з даними для порівняння")
        else:
            compare_month = st.selectbox(
                "Порівняти факт з планом за місяць",
                options=avail[-12:],
                format_func=lambda x: x.strftime("%B %Y"),
                index=len(avail[-12:]) - 1,
            )

            plan_for_month = compute_auto_plan(
                df[df.index < compare_month],
                st.session_state.store_starts,
                st.session_state.seasonality,
                compare_month,
                growth_old=0.0, growth_young=0.15, age_threshold=12,
            )
            fact_m = has_data.loc[compare_month]

            exec_df = execution_table(fact_m, plan_for_month)
            show_df = exec_df.dropna(subset=["Факт"]).sort_values("Виконання %", ascending=False)

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Магазинів з даними", len(show_df))
            c2.metric("Виконали план ✅", (show_df["Виконання %"] >= 100).sum())
            c3.metric("90–100% ⚠️",
                      ((show_df["Виконання %"] >= 90) & (show_df["Виконання %"] < 100)).sum())
            c4.metric("Нижче 90% 🔴", (show_df["Виконання %"] < 90).sum())

            st.markdown("---")
            st.dataframe(
                show_df[["Факт", "План", "Виконання %", "Відхилення", "Статус"]]
                .style
                .format({"Факт": "{:,.0f}", "План": "{:,.0f}",
                         "Виконання %": "{:.1f}%", "Відхилення": "{:+,.0f}"})
                .map(color_exec, subset=["Виконання %"]),
                use_container_width=True, height=500,
            )

            st.markdown("---")
            ca, cb = st.columns(2)
            with ca:
                st.caption("🏆 Топ-10 найкращих")
                top = show_df.head(10)[["Виконання %"]]
                top.index = [s.replace("Магазин - ", "") for s in top.index]
                st.bar_chart(top)
            with cb:
                st.caption("⚠️ Зона уваги (останні 10)")
                bot = show_df.tail(10)[["Виконання %"]].sort_values("Виконання %")
                bot.index = [s.replace("Магазин - ", "") for s in bot.index]
                st.bar_chart(bot)


# ── TAB 4: RECOMMENDATIONS ───────────────────────────────────────────────────
with tab4:
    st.subheader("💡 Рекомендації по плануванню")

    if st.session_state.df_fact is None:
        st.info("👈 Завантажте Excel-файл у боковій панелі")
    elif st.session_state.df_plan_auto is None:
        st.info("👈 Спочатку розрахуйте план")
    else:
        df          = st.session_state.df_fact
        store_starts= st.session_state.store_starts
        seasonality = st.session_state.seasonality
        pm          = st.session_state.plan_month

        has_data = df.replace(0, np.nan).dropna(how="all")
        avail    = [m for m in has_data.index if m < pm]
        last_n   = avail[-6:] if len(avail) >= 6 else avail

        # ── Build execution history: Excel plans first, app-plan as fallback ──
        excel_plans = st.session_state.excel_plans
        today = pd.Timestamp.today()

        # Pre-compute app fallback plans for stores missing from Excel
        fallback_plans = {}  # store -> {month -> plan_value}
        stores_no_excel = set()
        for store in df.columns:
            for month in last_n:
                fact_v = df.loc[month, store] if month in df.index else 0
                if pd.isna(fact_v) or fact_v == 0:
                    continue
                if excel_plans.get((store, month), 0) == 0:
                    stores_no_excel.add(store)
        for store in stores_no_excel:
            for month in last_n:
                if fallback_plans.get(store, {}).get(month): continue
                p = compute_auto_plan(
                    df[df.index < month], store_starts, seasonality, month,
                    growth_old=0.0, growth_young=0.15, age_threshold=12,
                )
                fallback_plans.setdefault(store, {})[month] = p.get(store, 0)

        exec_hist   = {}  # store -> [exec_pct, ...]
        exec_detail = {}  # store -> [(month, fact, plan, pct, plan_source_str), ...]
        MIN_PLAN_FOR_REC = 1000   # ignore near-zero bootstrap plans in recommendations

        for month in last_n:
            for store in df.columns:
                fact_v = df.loc[month, store] if month in df.index else 0
                if pd.isna(fact_v) or fact_v == 0:
                    continue
                excel_pv = excel_plans.get((store, month), 0)
                if excel_pv >= MIN_PLAN_FOR_REC:
                    pv = excel_pv
                    src_label = "Excel"
                else:
                    pv = fallback_plans.get(store, {}).get(month, 0)
                    src_label = "авто"
                    if pv < MIN_PLAN_FOR_REC:
                        continue   # skip months where plan was near-zero (new store ramp-up)
                pct = fact_v / pv * 100
                exec_hist.setdefault(store, []).append(pct)
                exec_detail.setdefault(store, []).append(
                    (month, round(fact_v), round(pv), round(pct, 1), src_label)
                )

        # ── Compute trend direction for each store (deseasoned last 6 months) ──
        def store_trend_pct(store):
            hist = df[store].replace(0, np.nan).dropna()
            hist = hist[hist.index < pm].iloc[-6:]
            if len(hist) < 3:
                return 0.0
            deseas = hist / hist.index.map(lambda d: seasonality.get(d.month, 1.0)) if seasonality else hist
            x = np.arange(len(deseas), dtype=float)
            slope, _ = np.polyfit(x, deseas.values.astype(float), 1)
            return slope / deseas.mean() * 100 if deseas.mean() > 0 else 0.0

        # ── Build recommendations ──
        recs_low, recs_high, recs_new, recs_trend = [], [], [], []

        for store, pcts in exec_hist.items():
            if len(pcts) < 2:
                continue
            avg   = np.mean(pcts)
            name  = store.replace("Магазин - ", "")
            start = store_starts.get(store)
            age   = ((today.year - start.year) * 12 + today.month - start.month) if start else 999
            detail_rows  = exec_detail.get(store, [])
            trend_pct_mo = store_trend_pct(store)
            trend_label  = (f"📈 зростає +{trend_pct_mo:.1f}%/міс" if trend_pct_mo > 1
                            else f"📉 спадає {trend_pct_mo:.1f}%/міс" if trend_pct_mo < -1
                            else "➡️ стабільний")

            # New stores: use broader band ±25% (plans are inherently less accurate)
            is_new = age < 12
            threshold_high = 125 if is_new else 115
            threshold_low  = 75  if is_new else 85

            if avg > threshold_high:
                gap = avg - 100
                trend_note = (f" Тренд {trend_label} — план варто підняти агресивніше."
                              if trend_pct_mo > 1 else
                              f" Тренд {trend_label} — перевірте чи ріст не сповільнився.")
                if is_new:
                    recs_new.append((name, avg, len(pcts), detail_rows, trend_pct_mo,
                        f"Новий магазин ({age} міс.), виконання {avg:.0f}% — план занижений. "
                        f"Тренд: {trend_label}. Рекомендуємо план = факт останнього місяця × (1 + сезонний коеф.)."))
                else:
                    recs_low.append((name, avg, len(pcts), detail_rows, trend_pct_mo,
                        f"Середнє виконання {avg:.0f}% за {len(pcts)} міс. — план занижений на ≈{gap:.0f}%.{trend_note}"))

            elif avg < threshold_low:
                gap = 100 - avg
                if is_new:
                    recs_new.append((name, avg, len(pcts), detail_rows, trend_pct_mo,
                        f"Новий магазин ({age} міс.), виконання {avg:.0f}% — план завищений. "
                        f"Тренд: {trend_label}. Рекомендуємо план = факт останнього місяця × (1 + сезонний коеф.)."))
                else:
                    trend_note = (f" Тренд {trend_label} — спад може продовжитись."
                                  if trend_pct_mo < -1 else
                                  f" Тренд {trend_label} — перевірте чи є разові причини.")
                    recs_high.append((name, avg, len(pcts), detail_rows, trend_pct_mo,
                        f"Середнє виконання {avg:.0f}% за {len(pcts)} міс. — план завищений на ≈{gap:.0f}%.{trend_note}"))

            else:
                # Plan is within acceptable range — flag only strong trend divergence for mature stores
                if abs(trend_pct_mo) > 3 and not is_new:
                    action = "підвищити" if trend_pct_mo > 0 else "знизити"
                    recs_trend.append((name, avg, len(pcts), detail_rows, trend_pct_mo,
                        f"Виконання в нормі ({avg:.0f}%), але тренд {trend_label} — "
                        f"варто {action} план на наступний місяць."))

        total = len(recs_low) + len(recs_high) + len(recs_new) + len(recs_trend)
        n_excel = sum(1 for (s, m), v in excel_plans.items() if v > 0)
        st.markdown(
            f"Знайдено **{total} рекомендацій** на основі **{len(last_n)} місяців** "
            f"({n_excel} планових значень з Excel, {len(stores_no_excel)} магазинів — авто-план)"
        )

        def _render_detail(detail_rows):
            if not detail_rows:
                return
            has_src = len(detail_rows[0]) == 5
            rows_html = ""
            for row in sorted(detail_rows, key=lambda x: x[0]):
                m, fact, plan, pct = row[0], row[1], row[2], row[3]
                src_lbl = row[4] if has_src else ""
                clr = "#2e7d32" if pct >= 100 else "#e65100" if pct >= 85 else "#c62828"
                src_badge = (f"<span style='font-size:10px;padding:1px 5px;border-radius:3px;"
                             f"background:#e3e3e3;color:#555;margin-left:4px'>{src_lbl}</span>"
                             if src_lbl else "")
                rows_html += (
                    f"<tr>"
                    f"<td style='padding:2px 10px 2px 0;color:var(--color-text-secondary);font-size:12px'>{m.strftime('%b %Y')}</td>"
                    f"<td style='padding:2px 10px 2px 0;font-size:12px'>{fact:,}</td>"
                    f"<td style='padding:2px 10px 2px 0;font-size:12px'>{plan:,}{src_badge}</td>"
                    f"<td style='padding:2px 0;font-size:12px;font-weight:500;color:{clr}'>{pct:.0f}%</td>"
                    f"</tr>"
                )
            st.markdown(
                f"<table style='margin-top:6px'>"
                f"<tr>"
                f"<th style='padding:2px 10px 2px 0;font-size:11px;color:var(--color-text-secondary)'>Місяць</th>"
                f"<th style='padding:2px 10px 2px 0;font-size:11px;color:var(--color-text-secondary)'>Факт</th>"
                f"<th style='padding:2px 10px 2px 0;font-size:11px;color:var(--color-text-secondary)'>План</th>"
                f"<th style='padding:2px 0;font-size:11px;color:var(--color-text-secondary)'>Вик.%</th>"
                f"</tr>{rows_html}</table>",
                unsafe_allow_html=True
            )

        if recs_low:
            with st.expander(f"🔼 План занижений — {len(recs_low)} магазин(и)", expanded=True):
                for name, avg, n, detail_rows, trend_pct, msg in sorted(recs_low, key=lambda x: -x[1]):
                    st.markdown(
                        f'<div class="rec-card rec-low">'
                        f'<strong>{name}</strong> — виконання {avg:.0f}%<br>'
                        f'<span style="font-size:0.9rem">{msg}</span></div>',
                        unsafe_allow_html=True
                    )
                    _render_detail(detail_rows)

        if recs_high:
            with st.expander(f"🔽 План завищений — {len(recs_high)} магазин(и)", expanded=True):
                for name, avg, n, detail_rows, trend_pct, msg in sorted(recs_high, key=lambda x: x[1]):
                    st.markdown(
                        f'<div class="rec-card rec-high">'
                        f'<strong>{name}</strong> — виконання {avg:.0f}%<br>'
                        f'<span style="font-size:0.9rem">{msg}</span></div>',
                        unsafe_allow_html=True
                    )
                    _render_detail(detail_rows)

        if recs_new:
            with st.expander(f"🆕 Нові магазини — {len(recs_new)} магазин(и)", expanded=True):
                for name, avg, n, detail_rows, trend_pct, msg in sorted(recs_new, key=lambda x: x[1]):
                    st.markdown(
                        f'<div class="rec-card rec-new">'
                        f'<strong>{name}</strong> — виконання {avg:.0f}%<br>'
                        f'<span style="font-size:0.9rem">{msg}</span></div>',
                        unsafe_allow_html=True
                    )
                    _render_detail(detail_rows)

        if recs_trend:
            with st.expander(f"📈 Тренд розходиться з планом — {len(recs_trend)} магазин(и)", expanded=True):
                for name, avg, n, detail_rows, trend_pct, msg in sorted(recs_trend, key=lambda x: -abs(x[4])):
                    clr = "rec-low" if trend_pct > 0 else "rec-high"
                    st.markdown(
                        f'<div class="rec-card {clr}">'
                        f'<strong>{name}</strong> — виконання {avg:.0f}%, тренд {trend_pct:+.1f}%/міс<br>'
                        f'<span style="font-size:0.9rem">{msg}</span></div>',
                        unsafe_allow_html=True
                    )
                    _render_detail(detail_rows)

        if total == 0:
            st.success("✅ Всі магазини виконують план у нормі (85–115%) і тренд відповідає плану. Рекомендацій немає.")

        # ── Store age table ──────────────────────────────────────────────────
        st.markdown("---")
        st.markdown("**Вік магазинів мережі**")
        age_rows = []
        for store, start in store_starts.items():
            if start:
                age = (today.year-start.year)*12 + today.month-start.month
                age_rows.append({
                    "Магазин": store.replace("Магазин - ", ""),
                    "Дата відкриття": start.strftime("%b %Y"),
                    "Вік (міс.)": age,
                    "Категорія": "Зрілий (>12 міс.)" if age >= 12 else "Новий (≤12 міс.)"
                })
        age_df = pd.DataFrame(age_rows).sort_values("Вік (міс.)", ascending=False)
        c1, c2 = st.columns(2)
        c1.metric("Зрілих магазинів", (age_df["Вік (міс.)"] >= 12).sum())
        c2.metric("Нових (≤12 міс.)", (age_df["Вік (міс.)"] < 12).sum())
        st.dataframe(age_df, use_container_width=True, hide_index=True, height=300)


# ── TAB 5: PROMO CALENDAR ────────────────────────────────────────────────────
with tab5:
    st.subheader("📅 Промокалендар")

    promo_df = st.session_state.promo_calendar
    has_promo = len(promo_df) > 0

    # ── Status: auto-loaded sources ─────────────────────────────────────────
    import os
    _static_exists = bool(STATIC_PROMO_FILE and os.path.exists(
        os.path.join(os.path.dirname(__file__), STATIC_PROMO_FILE)))
    _gsheet_config = bool(GSHEET_PROMO_URL)

    if _static_exists or _gsheet_config:
        src_parts = []
        if _static_exists:
            src_parts.append(f"файл `{STATIC_PROMO_FILE}` з репозиторію")
        if _gsheet_config:
            src_parts.append("Google Таблиця з конфігу")
        if has_promo:
            st.success(f"✓ Промокалендар завантажено автоматично з: {', '.join(src_parts)} "
                       f"({len(promo_df)} акцій)")
        else:
            st.warning(f"Налаштовано джерела ({', '.join(src_parts)}), але завантаження не вдалось. "
                       f"Завантажте вручну нижче.")
    else:
        st.info("Джерела промокалендаря не налаштовані в коді. "
                "Завантажте вручну або додайте шлях/посилання в `STATIC_PROMO_FILE` / `GSHEET_PROMO_URL`.")

    # ── Section 1: Manual upload (fallback / add extra data) ────────────────
    with st.expander("Завантажити файл вручну (Excel/CSV)", expanded=not has_promo):
        st.caption("Формат: колонки ДАТА ПОЧАТКУ, ДАТА ЗАКІНЧЕННЯ, ТИП, НАЗВА")
        static_file = st.file_uploader(
            "Файл з акціями",
            type=["xlsx", "csv"],
            key="promo_static_upload",
        )
        if static_file:
            try:
                _raw = pd.read_csv(static_file, header=None) if static_file.name.endswith(".csv")                        else pd.read_excel(static_file, header=None)
                parsed = parse_promo_df(_raw)
                if has_promo:
                    merged = pd.concat([promo_df, parsed], ignore_index=True).drop_duplicates(
                        subset=["date_start", "date_end", "promo_type"]).reset_index(drop=True)
                else:
                    merged = parsed
                st.session_state.promo_calendar = merged
                st.session_state.promo_static_loaded = True
                promo_df = merged
                has_promo = True
                st.success(f"✓ Додано {len(parsed)} акцій (всього: {len(merged)})")
            except Exception as e:
                st.error(f"Помилка читання файлу: {e}")

    # ── Section 2: Google Sheets link (manual override / refresh) ───────────
    with st.expander("Оновити з Google Sheets вручну", expanded=not _gsheet_config):
        st.caption(
            "Таблиця має бути відкрита для перегляду: Файл → Поділитися → "
            "Будь-хто з посиланням може переглядати."
        )
        default_url = GSHEET_PROMO_URL or ""
        gsheet_url = st.text_input(
            "Посилання на Google Таблицю",
            value=default_url,
            placeholder="https://docs.google.com/spreadsheets/d/...",
            key="gsheet_url_input",
        )
        if st.button("Завантажити / оновити з Google Sheets"):
            url_to_use = gsheet_url.strip() or GSHEET_PROMO_URL
            if url_to_use:
                try:
                    with st.spinner("Завантаження..."):
                        gdf = load_promo_from_gsheet(url_to_use)
                    base = promo_df if has_promo else pd.DataFrame()
                    combined = pd.concat([base, gdf], ignore_index=True).drop_duplicates(
                        subset=["date_start", "date_end", "promo_type"]).reset_index(drop=True)
                    st.session_state.promo_calendar = combined
                    promo_df = combined
                    has_promo = True
                    st.success(f"✓ Оновлено: {len(gdf)} акцій з Google Sheets (всього: {len(combined)})")
                except requests.exceptions.HTTPError:
                    st.error("Немає доступу. Відкрийте таблицю для перегляду.")
                except Exception as e:
                    st.error(f"Помилка: {e}")
            else:
                st.warning("Введіть посилання")

    # ── Section 3: Preview and monthly summary ──────────────────────────────
    if has_promo:
        st.markdown("---")
        pm = st.session_state.plan_month

        col_a, col_b, col_c = st.columns(3)
        col_a.metric("Акцій у календарі", len(promo_df))
        col_b.metric("Діапазон дат",
                     f"{promo_df['date_start'].min().strftime('%b %Y')} – "
                     f"{promo_df['date_end'].max().strftime('%b %Y')}")

        # Monthly summary table
        st.markdown("**Місячний підсумок**")
        months_range = pd.date_range(
            promo_df["date_start"].min().to_period("M").to_timestamp(),
            promo_df["date_end"].max().to_period("M").to_timestamp(),
            freq="MS"
        )
        summary_rows = []
        for m in months_range:
            score = compute_monthly_promo_score(promo_df, m)
            summary_rows.append({
                "Місяць": m.strftime("%b %Y"),
                "К-сть акцій": score["n_promos"],
                "ТОВАР МІСЯЦЯ": "✅" if score["has_anchor"] else "—",
                "Глибока механіка": "✅" if score["has_deep"] else "—",
                "Інтенсивність": score["intensity"],
                "Lift-фактор": score["lift_factor"],
            })
        summary_df = pd.DataFrame(summary_rows)

        # Highlight plan month
        def highlight_plan_month(row):
            if pm and row["Місяць"] == pm.strftime("%b %Y"):
                return ["background-color:#e8f5e9"] * len(row)
            return [""] * len(row)

        st.dataframe(
            summary_df.style
                .apply(highlight_plan_month, axis=1)
                .format({"Інтенсивність": "{:.1f}", "Lift-фактор": "{:.3f}"}),
            use_container_width=True,
            hide_index=True,
            height=380
        )

        if pm:
            score_pm = compute_monthly_promo_score(promo_df, pm)
            st.markdown(f"**Місяць плану ({pm.strftime('%B %Y')}):** "
                        f"lift-фактор = **{score_pm['lift_factor']:.3f}** "
                        f"({'ТОВАР МІСЯЦЯ ✅' if score_pm['has_anchor'] else '—'}, "
                        f"{'глибока механіка ✅' if score_pm['has_deep'] else '—'})")
            if score_pm["lift_factor"] > 1.0:
                st.info(
                    f"Промо-поправка для плану на {pm.strftime('%B %Y')}: "
                    f"×{score_pm['lift_factor']:.3f} (+{(score_pm['lift_factor']-1)*100:.0f}%). "
                    f"Натисніть «Розрахувати план» у боковій панелі — поправка застосується автоматично."
                )

        # Raw data view
        with st.expander("Переглянути всі акції"):
            st.dataframe(
                promo_df.sort_values("date_start").assign(
                    date_start=lambda x: x.date_start.dt.strftime("%d.%m.%Y"),
                    date_end=lambda x: x.date_end.dt.strftime("%d.%m.%Y")
                ).rename(columns={
                    "date_start": "Початок", "date_end": "Кінець",
                    "promo_type": "Тип", "name": "Назва"
                }),
                use_container_width=True,
                hide_index=True,
            )
    else:
        st.info("Завантажте промокалендар щоб побачити місячну аналітику та lift-фактори для плану.")
