import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import json
from datetime import datetime, date
import warnings
warnings.filterwarnings('ignore')

st.set_page_config(
    page_title="Планування продажів",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── CSS ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
.metric-card {
    background: #f8f9fa; border-radius: 10px; padding: 16px 20px;
    border-left: 4px solid #4CAF50;
}
.metric-card.warn { border-left-color: #FF9800; }
.metric-card.danger { border-left-color: #F44336; }
.section-header {
    font-size: 1.1rem; font-weight: 600; color: #1a1a2e;
    border-bottom: 2px solid #e0e0e0; padding-bottom: 6px; margin-bottom: 12px;
}
</style>
""", unsafe_allow_html=True)

# ── Session state ─────────────────────────────────────────────────────────────
if "df_fact" not in st.session_state:
    st.session_state.df_fact = None
if "df_plan_auto" not in st.session_state:
    st.session_state.df_plan_auto = None
if "df_plan_edited" not in st.session_state:
    st.session_state.df_plan_edited = None
if "seasonality" not in st.session_state:
    st.session_state.seasonality = None
if "store_starts" not in st.session_state:
    st.session_state.store_starts = {}
if "plan_month" not in st.session_state:
    st.session_state.plan_month = None

# ── Helpers ───────────────────────────────────────────────────────────────────

def load_excel(file) -> dict:
    """Reads ДАНІ ВНОСИТИ and Коеф. Сезонності sheets."""
    xl = pd.read_excel(file, sheet_name=None, header=None)
    return xl


def parse_fact_sheet(xl: dict) -> tuple[pd.DataFrame, dict, pd.Series]:
    """Parse ДАНІ ВНОСИТИ → long-format fact DataFrame."""
    raw = xl.get("ДАНІ ВНОСИТИ")
    if raw is None:
        st.error("Не знайдено лист 'ДАНІ ВНОСИТИ'")
        return None, {}, None

    stores = list(raw.iloc[0, 2:].dropna())
    date_rows = raw.iloc[3:, 0].reset_index(drop=True)
    data_block = raw.iloc[3:, 2:2+len(stores)].reset_index(drop=True)
    data_block.columns = stores

    df = data_block.copy()
    df.insert(0, "date", pd.to_datetime(date_rows))
    df = df[df["date"].notna()].copy()
    df = df.set_index("date")
    df = df.apply(pd.to_numeric, errors="coerce")

    # Store start dates (first non-null row)
    store_starts = {}
    for s in stores:
        nonzero = df[s].replace(0, np.nan).dropna()
        if len(nonzero) > 0:
            store_starts[s] = nonzero.index[0]
        else:
            store_starts[s] = None

    # Parse seasonality
    seas_raw = xl.get("Коеф. Сезонності")
    seasonality = pd.Series(dtype=float)
    if seas_raw is not None:
        for _, row in seas_raw.iterrows():
            if str(row.iloc[13]).strip() == "номер місяця":
                continue
            try:
                month_num = int(float(row.iloc[13]))
                coef = float(row.iloc[14])
                seasonality[month_num] = coef
            except (ValueError, TypeError):
                pass

    return df, store_starts, seasonality


def compute_auto_plan(
    df: pd.DataFrame,
    store_starts: dict,
    seasonality: pd.Series,
    target_month: pd.Timestamp,
    growth_old: float = 0.0,
    growth_young: float = 0.15,
    age_threshold_months: int = 12,
) -> pd.Series:
    """
    For each store, compute a plan value for target_month.

    Algorithm mirrors the Excel logic:
    - Stores older than age_threshold_months: take avg of same month last N years
      (weighted by seasonality), apply growth_old
    - Stores younger: take last known value * seasonal adjustment * growth_young
    """
    plan = {}
    m = target_month.month
    seas_coef = seasonality.get(m, 1.0)

    for store in df.columns:
        start = store_starts.get(store)
        if start is None:
            plan[store] = 0.0
            continue

        history = df[store].replace(0, np.nan).dropna()
        if len(history) == 0:
            plan[store] = 0.0
            continue

        age_months = (target_month.year - start.year) * 12 + (target_month.month - start.month)

        if age_months >= age_threshold_months:
            # Mature store: use same-month history from recent years
            same_month = history[history.index.month == m]
            if len(same_month) == 0:
                # Fallback: use 12-month trailing average * seasonality
                recent = history.iloc[-12:]
                base = recent.mean() * seas_coef
            else:
                base = same_month.mean()
            plan[store] = round(base * (1 + growth_old), 2)
        else:
            # Young store: last known value * seasonal adjustment
            recent = history.iloc[-min(3, len(history)):]
            avg_recent = recent.mean()
            # Seasonal ratio: target month vs average month
            avg_seas = seasonality.mean() if len(seasonality) > 0 else 1.0
            seas_ratio = seas_coef / avg_seas if avg_seas > 0 else 1.0
            plan[store] = round(avg_recent * seas_ratio * (1 + growth_young), 2)

    return pd.Series(plan)


def compute_execution(fact: pd.Series, plan: pd.Series) -> pd.DataFrame:
    """Returns store-level execution table."""
    df = pd.DataFrame({"Факт": fact, "План": plan})
    df["Виконання %"] = np.where(
        df["План"] > 0,
        (df["Факт"] / df["План"] * 100).round(1),
        np.nan
    )
    df["Відхилення"] = (df["Факт"] - df["План"]).round(2)
    return df


def flag_status(pct):
    if pd.isna(pct):
        return "—"
    if pct >= 100:
        return "✅"
    if pct >= 90:
        return "⚠️"
    return "🔴"


def to_excel_download(df_fact, df_plan_auto, df_plan_edited, plan_month):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        if df_fact is not None:
            df_fact.reset_index().to_excel(writer, sheet_name="Факт", index=False)
        if df_plan_auto is not None:
            pd.DataFrame({
                "Магазин": df_plan_auto.index,
                "Авто-план": df_plan_auto.values,
                "Місяць": str(plan_month)[:7] if plan_month else ""
            }).to_excel(writer, sheet_name="Авто-план", index=False)
        if df_plan_edited is not None:
            pd.DataFrame({
                "Магазин": df_plan_edited.index,
                "Відредагований план": df_plan_edited.values,
                "Місяць": str(plan_month)[:7] if plan_month else ""
            }).to_excel(writer, sheet_name="Відред-план", index=False)
    buf.seek(0)
    return buf


# ═══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.title("📊 Планування продажів")
    st.markdown("---")

    uploaded = st.file_uploader(
        "Завантажте Excel-файл",
        type=["xlsx"],
        help="Файл повинен містити лист 'ДАНІ ВНОСИТИ'"
    )

    if uploaded:
        xl = load_excel(uploaded)
        df_fact, store_starts, seasonality = parse_fact_sheet(xl)
        if df_fact is not None:
            st.session_state.df_fact = df_fact
            st.session_state.store_starts = store_starts
            st.session_state.seasonality = seasonality
            st.success(f"✓ Завантажено {len(df_fact.columns)} магазинів")

    st.markdown("---")
    st.markdown("**Параметри плану**")

    # Month picker
    next_month = pd.Timestamp.today().to_period("M").to_timestamp() + pd.offsets.MonthBegin(1)
    month_str = st.text_input(
        "Місяць плану (РРРР-ММ)",
        value=next_month.strftime("%Y-%m"),
        help="Наприклад: 2026-07"
    )
    try:
        plan_month = pd.Timestamp(month_str + "-01")
    except Exception:
        plan_month = next_month
    st.session_state.plan_month = plan_month

    growth_old = st.slider(
        "Приріст, зрілі магазини (%)",
        min_value=-20, max_value=50, value=0, step=1
    ) / 100

    growth_young = st.slider(
        "Приріст, нові магазини (%)",
        min_value=-20, max_value=50, value=15, step=1
    ) / 100

    age_threshold = st.slider(
        "Поріг «молодий магазин» (міс.)",
        min_value=6, max_value=24, value=12, step=1
    )

    if st.button("🤖 Розрахувати план", type="primary", use_container_width=True):
        if st.session_state.df_fact is not None:
            plan = compute_auto_plan(
                st.session_state.df_fact,
                st.session_state.store_starts,
                st.session_state.seasonality,
                plan_month,
                growth_old=growth_old,
                growth_young=growth_young,
                age_threshold_months=age_threshold,
            )
            st.session_state.df_plan_auto = plan
            st.session_state.df_plan_edited = plan.copy()
            st.success("✓ План розраховано")
        else:
            st.warning("Спочатку завантажте файл")


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN TABS
# ═══════════════════════════════════════════════════════════════════════════════
tab1, tab2, tab3, tab4 = st.tabs([
    "📁 Дані", "✏️ Редагування плану", "📊 Аналітика", "💡 Рекомендації"
])


# ─────────────────────────────────────────────────────────────────────────────
# TAB 1: DATA
# ─────────────────────────────────────────────────────────────────────────────
with tab1:
    st.subheader("Завантажені дані — факт продажів по магазинах")

    if st.session_state.df_fact is None:
        st.info("👈 Завантажте Excel-файл у боковій панелі")
    else:
        df = st.session_state.df_fact
        stores = list(df.columns)

        col1, col2, col3 = st.columns(3)
        col1.metric("Магазинів", len(stores))
        col2.metric("Місяців даних", len(df))
        col3.metric(
            "Останній місяць з даними",
            df.replace(0, np.nan).dropna(how="all").index[-1].strftime("%B %Y")
        )

        st.markdown("---")
        st.markdown("**Факт продажів (USD з ПДВ) — таблиця**")

        # Filter controls
        c1, c2 = st.columns([3, 1])
        search = c1.text_input("Пошук магазину", placeholder="Введіть назву...")
        show_last = c2.selectbox("Показати місяців", [6, 12, 24, 99], index=1)

        filtered_stores = [s for s in stores if search.lower() in s.lower()] if search else stores
        disp = df[filtered_stores].iloc[-show_last:].copy()
        disp.index = disp.index.strftime("%b %Y")
        disp = disp.replace(0, np.nan)

        st.dataframe(
            disp.style.format("{:,.0f}", na_rep="—").highlight_null(color="#fff3cd"),
            use_container_width=True,
            height=400
        )

        # Quick chart
        st.markdown("**Тренд продажів — обраний магазин**")
        sel_store = st.selectbox("Магазин для графіку", filtered_stores, key="chart_store")
        chart_data = df[[sel_store]].replace(0, np.nan).dropna()
        st.line_chart(chart_data, use_container_width=True)


# ─────────────────────────────────────────────────────────────────────────────
# TAB 2: PLAN EDITING
# ─────────────────────────────────────────────────────────────────────────────
with tab2:
    st.subheader(f"Редагування плану на {st.session_state.plan_month.strftime('%B %Y') if st.session_state.plan_month else '—'}")

    if st.session_state.df_plan_auto is None:
        st.info("👈 Розрахуйте план у боковій панелі")
    else:
        plan_auto = st.session_state.df_plan_auto
        plan_edited = st.session_state.df_plan_edited.copy()

        st.markdown("**Глобальне коригування**")
        col_g1, col_g2, col_g3 = st.columns(3)
        global_pct = col_g1.number_input(
            "Приріст до авто-плану (%)", min_value=-50.0, max_value=100.0, value=0.0, step=0.5
        )
        if col_g2.button("Застосувати до всіх"):
            st.session_state.df_plan_edited = (plan_auto * (1 + global_pct / 100)).round(2)
            st.rerun()
        if col_g3.button("↩ Скинути до авто-плану"):
            st.session_state.df_plan_edited = plan_auto.copy()
            st.rerun()

        st.markdown("---")
        st.markdown("**Ручне редагування по кожному магазину**")
        st.caption("Змінюйте значення плану безпосередньо в таблиці")

        edit_df = pd.DataFrame({
            "Магазин": plan_auto.index,
            "Авто-план": plan_auto.values.round(2),
            "Ваш план": st.session_state.df_plan_edited.values.round(2),
        }).set_index("Магазин")

        edited = st.data_editor(
            edit_df,
            use_container_width=True,
            column_config={
                "Авто-план": st.column_config.NumberColumn("Авто-план", disabled=True, format="%.0f"),
                "Ваш план": st.column_config.NumberColumn("Ваш план ✏️", format="%.0f", min_value=0),
            },
            height=600,
        )
        st.session_state.df_plan_edited = pd.Series(
            edited["Ваш план"].values, index=edited.index
        )

        # Summary
        st.markdown("---")
        c1, c2, c3 = st.columns(3)
        c1.metric("Сума авто-плану", f"{plan_auto.sum():,.0f} USD")
        c2.metric("Сума вашого плану", f"{st.session_state.df_plan_edited.sum():,.0f} USD")
        delta = st.session_state.df_plan_edited.sum() - plan_auto.sum()
        c3.metric("Різниця", f"{delta:+,.0f} USD",
                  delta_color="normal" if delta >= 0 else "inverse")

        # Download
        st.markdown("---")
        buf = to_excel_download(
            st.session_state.df_fact,
            plan_auto,
            st.session_state.df_plan_edited,
            st.session_state.plan_month
        )
        st.download_button(
            "⬇️ Завантажити план у Excel",
            data=buf,
            file_name=f"план_{st.session_state.plan_month.strftime('%Y_%m')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


# ─────────────────────────────────────────────────────────────────────────────
# TAB 3: ANALYTICS
# ─────────────────────────────────────────────────────────────────────────────
with tab3:
    st.subheader("Аналітика виконання плану")

    if st.session_state.df_fact is None:
        st.info("👈 Завантажте Excel-файл у боковій панелі")
    elif st.session_state.df_plan_auto is None:
        st.info("👈 Спочатку розрахуйте план")
    else:
        df = st.session_state.df_fact
        plan = st.session_state.df_plan_edited
        plan_month = st.session_state.plan_month

        # Find months with actual data for comparison
        available_months = df.replace(0, np.nan).dropna(how="all").index.tolist()

        compare_months = [m for m in available_months if m < plan_month]
        if compare_months:
            compare_month = st.selectbox(
                "Порівняти факт з планом за місяць",
                options=compare_months[-12:],
                format_func=lambda x: x.strftime("%B %Y"),
                index=len(compare_months[-12:]) - 1
            )

            fact_month = df.loc[compare_month].replace(0, np.nan)

            # Estimate what plan WOULD have been for that month
            plan_for_compare = compute_auto_plan(
                df[df.index < compare_month],
                st.session_state.store_starts,
                st.session_state.seasonality,
                compare_month,
                growth_old=0.0,
                growth_young=0.15,
                age_threshold_months=12,
            )

            exec_df = compute_execution(fact_month, plan_for_compare)
            exec_df["Статус"] = exec_df["Виконання %"].apply(flag_status)
            exec_df_show = exec_df.dropna(subset=["Факт"]).sort_values("Виконання %", ascending=False)

            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Магазинів з даними", len(exec_df_show))
            pct_above = (exec_df_show["Виконання %"] >= 100).sum()
            col2.metric("Виконали план ✅", pct_above)
            col3.metric("90–100% ⚠️", ((exec_df_show["Виконання %"] >= 90) & (exec_df_show["Виконання %"] < 100)).sum())
            col4.metric("Нижче 90% 🔴", (exec_df_show["Виконання %"] < 90).sum())

            st.markdown("---")
            st.markdown("**Виконання плану по магазинах**")

            def color_execution(val):
                if pd.isna(val):
                    return ""
                if val >= 100:
                    return "background-color: #c8e6c9; color: #1b5e20"
                if val >= 90:
                    return "background-color: #fff9c4; color: #f57f17"
                return "background-color: #ffcdd2; color: #b71c1c"

            st.dataframe(
                exec_df_show[["Факт", "План", "Виконання %", "Відхилення", "Статус"]].style
                    .format({
                        "Факт": "{:,.0f}",
                        "План": "{:,.0f}",
                        "Виконання %": "{:.1f}%",
                        "Відхилення": "{:+,.0f}",
                    })
                    .map(color_execution, subset=["Виконання %"]),
                use_container_width=True,
                height=500,
            )

            st.markdown("---")
            st.markdown("**Топ-10 за виконанням**")
            c1, c2 = st.columns(2)
            with c1:
                st.caption("🏆 Найкращі")
                top10 = exec_df_show.head(10)[["Виконання %"]]
                top10.index = [s.replace("Магазин - ", "") for s in top10.index]
                st.bar_chart(top10)
            with c2:
                st.caption("⚠️ Зона уваги")
                bot10 = exec_df_show.tail(10)[["Виконання %"]].sort_values("Виконання %")
                bot10.index = [s.replace("Магазин - ", "") for s in bot10.index]
                st.bar_chart(bot10)
        else:
            st.info("Недостатньо даних для порівняння факт/план")


# ─────────────────────────────────────────────────────────────────────────────
# TAB 4: RECOMMENDATIONS
# ─────────────────────────────────────────────────────────────────────────────
with tab4:
    st.subheader("💡 Рекомендації по плануванню")

    if st.session_state.df_fact is None:
        st.info("👈 Завантажте Excel-файл у боковій панелі")
    elif st.session_state.df_plan_auto is None:
        st.info("👈 Спочатку розрахуйте план")
    else:
        df = st.session_state.df_fact

        # Compute execution over last 6 months
        available = df.replace(0, np.nan).dropna(how="all").index.tolist()
        last_n = available[-6:] if len(available) >= 6 else available
        seasonality = st.session_state.seasonality
        store_starts = st.session_state.store_starts

        execution_history = {}
        for month in last_n:
            plan_hist = compute_auto_plan(
                df[df.index < month],
                store_starts,
                seasonality,
                month,
                growth_old=0.0,
                growth_young=0.15,
            )
            fact_hist = df.loc[month].replace(0, np.nan)
            for store in df.columns:
                if pd.notna(fact_hist.get(store)) and fact_hist[store] > 0 and plan_hist.get(store, 0) > 0:
                    pct = fact_hist[store] / plan_hist[store] * 100
                    if store not in execution_history:
                        execution_history[store] = []
                    execution_history[store].append(pct)

        recs = []
        for store, pcts in execution_history.items():
            if len(pcts) < 2:
                continue
            avg_exec = np.mean(pcts)
            name = store.replace("Магазин - ", "")

            if avg_exec > 115:
                recs.append(("🔼 План занижений", name, avg_exec,
                    f"Середнє виконання {avg_exec:.0f}% за {len(pcts)} міс. — план хронічно занижений. Рекомендуємо підняти на {avg_exec-100:.0f}%."))
            elif avg_exec < 85:
                recs.append(("🔽 План завищений", name, avg_exec,
                    f"Середнє виконання {avg_exec:.0f}% за {len(pcts)} міс. — магазин стабільно не виконує план. Перегляньте базові показники."))

            # New/young store
            start = store_starts.get(store)
            if start:
                age = (pd.Timestamp.today().year - start.year) * 12 + (pd.Timestamp.today().month - start.month)
                if age < 12 and avg_exec < 90:
                    recs.append(("🆕 Новий магазин", name, avg_exec,
                        f"Магазин відкрився {age} міс. тому. Рекомендуємо консервативний план з приростом 5–10% від поточного рівня."))

        if recs:
            st.markdown(f"Знайдено **{len(recs)} рекомендацій** на основі останніх {len(last_n)} місяців")

            categories = sorted(set(r[0] for r in recs))
            for cat in categories:
                cat_recs = [r for r in recs if r[0] == cat]
                with st.expander(f"{cat} — {len(cat_recs)} магазин(и)", expanded=True):
                    for _, name, avg_exec, msg in cat_recs:
                        color = "#e8f5e9" if "занижений" in cat else "#fff3e0"
                        st.markdown(f"""
<div style="background:{color};border-radius:8px;padding:12px 16px;margin-bottom:8px">
  <strong>{name}</strong><br>
  <span style="font-size:0.9rem">{msg}</span>
</div>""", unsafe_allow_html=True)
        else:
            st.success("✅ Всі магазини виконують план в нормі (85–115%). Рекомендацій немає.")

        # Store age overview
        st.markdown("---")
        st.markdown("**Вік магазинів мережі**")
        age_data = []
        today = pd.Timestamp.today()
        for store, start in store_starts.items():
            if start:
                age = (today.year - start.year) * 12 + (today.month - start.month)
                age_data.append({
                    "Магазин": store.replace("Магазин - ", ""),
                    "Вік (міс.)": age,
                    "Категорія": "Зрілий (>12 міс.)" if age >= 12 else "Новий (≤12 міс.)"
                })
        age_df = pd.DataFrame(age_data).sort_values("Вік (міс.)", ascending=False)
        young = age_df[age_df["Категорія"].str.startswith("Новий")]
        mature = age_df[age_df["Категорія"].str.startswith("Зрілий")]
        c1, c2 = st.columns(2)
        c1.metric("Зрілих магазинів", len(mature))
        c2.metric("Нових магазинів (≤12 міс.)", len(young))
        st.dataframe(age_df, use_container_width=True, hide_index=True, height=300)
