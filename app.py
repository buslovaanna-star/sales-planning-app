import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import warnings
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

    if st.button("🤖 Розрахувати план", type="primary", use_container_width=True):
        if st.session_state.df_fact is not None:
            plan = compute_auto_plan(
                st.session_state.df_fact, st.session_state.store_starts,
                st.session_state.seasonality, plan_month,
                growth_old, growth_young, age_threshold,
                short_weight, short_window,
            )
            st.session_state.df_plan_auto   = plan
            st.session_state.df_plan_edited = plan.copy()
            st.success("✓ План розраховано")
        else:
            st.warning("Спочатку завантажте файл")


# ══════════════════════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════════════════════
tab1, tab2, tab3, tab4 = st.tabs(["📁 Дані", "✏️ Редагування плану", "📊 Аналітика", "💡 Рекомендації"])


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

        edit_df = pd.DataFrame({
            "Магазин": plan_auto.index,
            "Авто-план": plan_auto.values.round(2),
            "Ваш план": st.session_state.df_plan_edited.values.round(2),
        }).set_index("Магазин")

        edited = st.data_editor(
            edit_df, use_container_width=True,
            column_config={
                "Авто-план": st.column_config.NumberColumn("Авто-план", disabled=True, format="%.0f"),
                "Ваш план":  st.column_config.NumberColumn("Ваш план ✏️", format="%.0f", min_value=0),
            },
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

        for month in last_n:
            for store in df.columns:
                fact_v = df.loc[month, store] if month in df.index else 0
                if pd.isna(fact_v) or fact_v == 0:
                    continue
                excel_pv = excel_plans.get((store, month), 0)
                if excel_pv > 0:
                    pv = excel_pv
                    src_label = "Excel"
                else:
                    pv = fallback_plans.get(store, {}).get(month, 0)
                    src_label = "авто"
                if pv <= 0:
                    continue
                pct = fact_v / pv * 100
                exec_hist.setdefault(store, []).append(pct)
                exec_detail.setdefault(store, []).append(
                    (month, round(fact_v), round(pv), round(pct, 1), src_label)
                )

        # ── Compute trend direction for each store (deseasoned last 6 months) ──
        def store_trend_pct(store):
            """Monthly % change of deseasoned sales over last 6 months."""
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
            trend_pct_mo = store_trend_pct(store)   # monthly % change
            trend_label  = (f"📈 зростає +{trend_pct_mo:.1f}%/міс" if trend_pct_mo > 1
                            else f"📉 спадає {trend_pct_mo:.1f}%/міс" if trend_pct_mo < -1
                            else "➡️ стабільний")

            if avg > 115:
                gap = avg - 100
                trend_note = (f" Тренд {trend_label} — план варто підняти агресивніше."
                              if trend_pct_mo > 1 else
                              f" Тренд {trend_label} — перевірте чи ріст не сповільнився.")
                recs_low.append((name, avg, len(pcts), detail_rows, trend_pct_mo,
                    f"Середнє виконання {avg:.0f}% за {len(pcts)} міс. — план занижений на ≈{gap:.0f}%.{trend_note}"))

            elif avg < 85:
                gap = 100 - avg
                if age < 12:
                    recs_new.append((name, avg, len(pcts), detail_rows, trend_pct_mo,
                        f"Новий магазин ({age} міс.), середнє виконання {avg:.0f}%. "
                        f"Тренд: {trend_label}. Рекомендуємо план = поточний рівень × (1 + сезонний коеф.)."))
                else:
                    trend_note = (f" Тренд {trend_label} — спад може продовжитись, знизити план."
                                  if trend_pct_mo < -1 else
                                  f" Тренд {trend_label} — можливо разові провали, перевірте деталі.")
                    recs_high.append((name, avg, len(pcts), detail_rows, trend_pct_mo,
                        f"Середнє виконання {avg:.0f}% за {len(pcts)} міс. — план завищений на ≈{gap:.0f}%.{trend_note}"))

            else:
                # Plan is ±15% — but check if trend diverges significantly from plan
                if abs(trend_pct_mo) > 3 and age >= 12:
                    direction = "зростає" if trend_pct_mo > 0 else "спадає"
                    action    = "підвищити" if trend_pct_mo > 0 else "знизити"
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
