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


def deseasonalise(series: pd.Series, seasonality: dict) -> pd.Series:
    if not seasonality:
        return series
    return series / series.index.map(lambda d: seasonality.get(d.month, 1.0))


def compute_plan_store(series: pd.Series, target: pd.Timestamp,
                       seasonality: dict, store_start,
                       growth_old: float, growth_young: float,
                       age_threshold: int) -> float:
    """
    Mature store (≥ age_threshold months):
        deseasonalise last 12 non-zero months → linear trend 1-step-ahead → reseasonalise
    Young store (< age_threshold):
        same but use growth_young instead of growth_old
    """
    if store_start is None:
        return 0.0

    age = (target.year - store_start.year) * 12 + (target.month - store_start.month)
    growth = growth_old if age >= age_threshold else growth_young

    hist = series.replace(0, np.nan).dropna()
    hist = hist[hist.index < target]
    if len(hist) == 0:
        return 0.0

    if seasonality:
        deseas = deseasonalise(hist, seasonality)
    else:
        deseas = hist.copy()

    last12 = deseas.iloc[-12:]
    n = len(last12)
    if n >= 3:
        x = np.arange(n, dtype=float)
        y = last12.values.astype(float)
        slope, intercept = np.polyfit(x, y, 1)
        projected = slope * n + intercept
    elif n > 0:
        projected = last12.mean()
    else:
        return 0.0

    projected = max(projected, 0)
    seas_coef = seasonality.get(target.month, 1.0) if seasonality else 1.0
    return round(projected * seas_coef * (1 + growth), 2)


def compute_auto_plan(df, store_starts, seasonality, target_month,
                      growth_old=0.0, growth_young=0.15, age_threshold=12):
    return pd.Series({
        store: compute_plan_store(
            df[store], target_month, seasonality,
            store_starts.get(store), growth_old, growth_young, age_threshold
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

    if st.button("🤖 Розрахувати план", type="primary", use_container_width=True):
        if st.session_state.df_fact is not None:
            plan = compute_auto_plan(
                st.session_state.df_fact, st.session_state.store_starts,
                st.session_state.seasonality, plan_month,
                growth_old, growth_young, age_threshold,
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

        # Build execution history
        exec_hist = {}
        for month in last_n:
            plan_h = compute_auto_plan(
                df[df.index < month], store_starts, seasonality, month,
                growth_old=0.0, growth_young=0.15, age_threshold=12,
            )
            for store in df.columns:
                fact_v = df.loc[month, store] if month in df.index else 0
                if pd.isna(fact_v) or fact_v == 0:
                    continue
                pv = plan_h.get(store, 0)
                if pv > 0:
                    exec_hist.setdefault(store, []).append(fact_v / pv * 100)

        recs_low, recs_high, recs_new = [], [], []
        today = pd.Timestamp.today()

        for store, pcts in exec_hist.items():
            if len(pcts) < 2:
                continue
            avg = np.mean(pcts)
            name = store.replace("Магазин - ", "")
            start = store_starts.get(store)
            age = ((today.year - start.year) * 12 + today.month - start.month) if start else 999

            if avg > 115:
                recs_low.append((name, avg, len(pcts),
                    f"Середнє виконання {avg:.0f}% за {len(pcts)} міс. — план хронічно занижений. "
                    f"Рекомендуємо підвищити на ≈{avg-100:.0f}%."))
            elif avg < 85:
                if age < 12:
                    recs_new.append((name, avg, len(pcts),
                        f"Новий магазин ({age} міс.), середнє виконання {avg:.0f}% — "
                        f"рекомендуємо консервативний план (+5–10% від поточного рівня)."))
                else:
                    recs_high.append((name, avg, len(pcts),
                        f"Середнє виконання {avg:.0f}% за {len(pcts)} міс. — план стабільно завищений. "
                        f"Рекомендуємо знизити на ≈{100-avg:.0f}%."))

        total = len(recs_low) + len(recs_high) + len(recs_new)
        st.markdown(f"Знайдено **{total} рекомендацій** на основі останніх **{len(last_n)} місяців**")

        if recs_low:
            with st.expander(f"🔼 План занижений — {len(recs_low)} магазин(и)", expanded=True):
                for name, avg, n, msg in sorted(recs_low, key=lambda x: -x[1]):
                    st.markdown(
                        f'<div class="rec-card rec-low">'
                        f'<strong>{name}</strong> — виконання {avg:.0f}%<br>'
                        f'<span style="font-size:0.9rem">{msg}</span></div>',
                        unsafe_allow_html=True
                    )

        if recs_high:
            with st.expander(f"🔽 План завищений — {len(recs_high)} магазин(и)", expanded=True):
                for name, avg, n, msg in sorted(recs_high, key=lambda x: x[1]):
                    st.markdown(
                        f'<div class="rec-card rec-high">'
                        f'<strong>{name}</strong> — виконання {avg:.0f}%<br>'
                        f'<span style="font-size:0.9rem">{msg}</span></div>',
                        unsafe_allow_html=True
                    )

        if recs_new:
            with st.expander(f"🆕 Нові магазини — {len(recs_new)} магазин(и)", expanded=True):
                for name, avg, n, msg in sorted(recs_new, key=lambda x: x[1]):
                    st.markdown(
                        f'<div class="rec-card rec-new">'
                        f'<strong>{name}</strong> — виконання {avg:.0f}%<br>'
                        f'<span style="font-size:0.9rem">{msg}</span></div>',
                        unsafe_allow_html=True
                    )

        if total == 0:
            st.success("✅ Всі магазини виконують план у нормі (85–115%). Рекомендацій немає.")

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
