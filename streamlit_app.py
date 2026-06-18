# -*- coding: utf-8 -*-
"""
Финансовый дашборд «Никольская» — управленческий учёт собственника.
Apple iOS стиль. Читает Google Таблицу через ПУБЛИЧНЫЙ CSV-экспорт (gviz/tq).
Секреты не нужны. Вкладки: Сводка · Динамика · Контроль · Дивиденды.
"""

import re
import math
import urllib.request
from urllib.parse import quote
from io import StringIO
from datetime import datetime, timedelta

import pandas as pd
import altair as alt
import streamlit as st

# ──────────────────────────────────────────────────────────────────────────
# КОНФИГУРАЦИЯ
# ──────────────────────────────────────────────────────────────────────────
SHEET_ID = "1vGcHv7MgiGUkfB9masV7szypjFoSrghd-jE9GgyTF3s"
HEADER_ROW = 3          # заголовки в строке 4 (индекс 3)
DATA_START = 4          # данные с строки 5 (индекс 4)
CACHE_TTL = 300

OWNERS = ["Дарья", "Елена"]
SHARES = {"Дарья": 0.5, "Елена": 0.5}

C = {
    "blue": "#007AFF", "green": "#34C759", "orange": "#FF9500",
    "red": "#FF3B30", "purple": "#AF52DE", "pink": "#FF2D55",
    "teal": "#00C7BE", "ink": "#1D1D1F", "grey": "#8E8E93",
    "bg": "#F5F5F7", "card": "#FFFFFF",
}

DIV_ICONS = {
    "отдел продаж": "📞", "продаж": "📞", "аренда тз": "🏢",
    "аренда салона": "💈", "салон": "💅", "кафе": "☕",
    "тренировк": "🏋️", "аренда": "🏢",
}

RU_MONTHS = ["", "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
             "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]

st.set_page_config(page_title="Никольская · Финансы", page_icon="📊", layout="wide")

# ──────────────────────────────────────────────────────────────────────────
# СТИЛЬ (Apple iOS)
# ──────────────────────────────────────────────────────────────────────────
st.markdown("""
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
  html, body, [class*="css"], .stApp { font-family: -apple-system, 'SF Pro Display', 'Inter', sans-serif; }
  .stApp { background: #F5F5F7; }
  #MainMenu, footer, header { visibility: hidden; }
  .block-container { padding-top: 1.4rem; max-width: 1180px; }

  .hero-title { font-size: 30px; font-weight: 800; color: #1D1D1F; letter-spacing: -0.5px; margin: 0; }
  .hero-sub   { font-size: 15px; color: #8E8E93; margin: 2px 0 0 0; }

  .card { background:#FFFFFF; border-radius:20px; padding:22px 24px;
      box-shadow:0 1px 3px rgba(0,0,0,0.04), 0 8px 24px rgba(0,0,0,0.05); margin-bottom:16px; }
  .kpi-label { font-size:14px; font-weight:600; color:#8E8E93; margin-bottom:6px; }
  .kpi-value { font-size:38px; font-weight:800; letter-spacing:-1px; line-height:1.05; }
  .kpi-foot  { font-size:13px; color:#8E8E93; margin-top:6px; }

  .row-name { font-size:16px; font-weight:600; color:#1D1D1F; }
  .row-num  { font-size:14px; color:#8E8E93; }
  .bar-track { background:#E9E9EB; border-radius:8px; height:10px; width:100%; overflow:hidden; }
  .bar-fill  { height:10px; border-radius:8px; }
  .pct-pill  { font-size:14px; font-weight:700; padding:2px 10px; border-radius:20px; }

  .avatar { width:42px; height:42px; border-radius:50%; display:inline-flex; align-items:center;
      justify-content:center; color:#fff; font-weight:700; font-size:16px; margin-right:10px; }

  .alert { border-radius:14px; padding:14px 18px; margin-bottom:10px; font-size:15px; font-weight:500; }
  .alert-red    { background:rgba(255,59,48,0.10);  color:#C2261C; }
  .alert-orange { background:rgba(255,149,0,0.12);  color:#B5670A; }
  .alert-green  { background:rgba(52,199,89,0.12);  color:#1F7A38; }

  .stTabs [data-baseweb="tab-list"] { gap:6px; background:#E9E9EB; padding:4px; border-radius:14px; }
  .stTabs [data-baseweb="tab"] { border-radius:10px; padding:8px 20px; font-weight:600; font-size:15px; color:#1D1D1F; background:transparent; }
  .stTabs [aria-selected="true"] { background:#FFFFFF !important; box-shadow:0 1px 4px rgba(0,0,0,0.08); }
</style>
""", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────
# ПАРСЕРЫ
# ──────────────────────────────────────────────────────────────────────────
def to_number(x):
    if x is None:
        return 0.0
    if isinstance(x, (int, float)) and not isinstance(x, bool):
        return 0.0 if (isinstance(x, float) and math.isnan(x)) else float(x)
    s = str(x).strip()
    if not s:
        return 0.0
    s = s.replace("\xa0", "").replace(" ", "").replace("₽", "").replace("руб", "").replace(",", ".")
    s = re.sub(r"[^0-9.\-]", "", s)
    if s in ("", "-", ".", "-."):
        return 0.0
    if s.count(".") > 1:
        p = s.split(".")
        s = "".join(p[:-1]) + "." + p[-1]
    try:
        return float(s)
    except ValueError:
        return 0.0


def to_date(x):
    if x is None or x == "":
        return None
    s = str(x).strip()
    if not s:
        return None
    for fmt in ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    try:
        d = pd.to_datetime(s, dayfirst=True, errors="coerce")
        return None if pd.isna(d) else d.date()
    except Exception:
        return None


def cell(row, idx):
    return row[idx] if idx is not None and idx < len(row) else ""


def col_idx(headers, candidates, default):
    norm = [str(h).strip().lower() for h in headers]
    for cand in candidates:
        cand = cand.lower()
        for i, h in enumerate(norm):
            if cand in h:
                return i
    return default


def money(v):
    return f"{v:,.0f} ₽".replace(",", " ")


def div_icon(name):
    n = str(name).strip().lower()
    for key, ico in DIV_ICONS.items():
        if key in n:
            return ico
    return "•"


# ──────────────────────────────────────────────────────────────────────────
# ЗАГРУЗКА (публичный CSV)
# ──────────────────────────────────────────────────────────────────────────
@st.cache_data(ttl=CACHE_TTL, show_spinner=False)
def load_sheet(sheet_name):
    """Возвращает (headers, rows). Пусто, если лист недоступен/пуст."""
    url = (f"https://docs.google.com/spreadsheets/d/{SHEET_ID}"
           f"/gviz/tq?tqx=out:csv&sheet={quote(sheet_name)}")
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=20) as resp:
            text = resp.read().decode("utf-8")
    except Exception:
        return [], []
    try:
        raw = pd.read_csv(StringIO(text), header=None, dtype=str, keep_default_na=False)
    except Exception:
        return [], []
    if len(raw) <= HEADER_ROW:
        return [], []
    headers = list(raw.iloc[HEADER_ROW])
    rows = [list(r) for _, r in raw.iloc[DATA_START:].iterrows()
            if any(str(c).strip() for c in r)]
    return headers, rows


def access_ok():
    """Проверяем, открыта ли таблица. Берём заведомо существующий лист."""
    url = (f"https://docs.google.com/spreadsheets/d/{SHEET_ID}"
           f"/gviz/tq?tqx=out:csv&sheet={quote('План на год')}")
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=20) as resp:
            return resp.status == 200
    except Exception:
        return False


# ─── Доходы ───
@st.cache_data(ttl=CACHE_TTL, show_spinner=False)
def get_income():
    recs = []
    for sheet, pay in [("Доходы нал", "Наличные"), ("Доходы безнал", "Безнал")]:
        headers, rows = load_sheet(sheet)
        if not rows:
            continue
        i_date = col_idx(headers, ["дата"], 0)
        i_div  = col_idx(headers, ["подразделен"], 2)
        i_sum  = col_idx(headers, ["сумма"], 3)
        i_ctr  = col_idx(headers, ["контрагент"], 4)
        i_type = col_idx(headers, ["тип операц", "тип"], 6)
        for r in rows:
            d = to_date(cell(r, i_date))
            if not d:
                continue
            op = str(cell(r, i_type)).strip().lower()
            is_loan = ("займ" in op) and ("получ" in op or op == "займ")
            recs.append({
                "date": d, "div": str(cell(r, i_div)).strip() or "Прочее",
                "sum": to_number(cell(r, i_sum)), "ctr": str(cell(r, i_ctr)).strip(),
                "pay": pay, "is_op": "займ" not in op, "is_loan_in": is_loan,
            })
    return pd.DataFrame(recs)


# ─── Расходы ───
@st.cache_data(ttl=CACHE_TTL, show_spinner=False)
def get_expense():
    recs = []
    for sheet, pay in [("Расходы нал", "Наличные"), ("Расходы безнал", "Безнал")]:
        headers, rows = load_sheet(sheet)
        if not rows:
            continue
        i_date = col_idx(headers, ["дата"], 0)
        i_dep  = col_idx(headers, ["департамент", "подразделен"], 2)
        i_art  = col_idx(headers, ["статья"], 3)
        i_sum  = col_idx(headers, ["сумма"], 4)
        i_ctr  = col_idx(headers, ["контрагент"], 5)
        i_type = col_idx(headers, ["тип операц", "тип"], 7)
        for r in rows:
            d = to_date(cell(r, i_date))
            if not d:
                continue
            op = str(cell(r, i_type)).strip().lower()
            is_loan = ("займ" in op) and ("возвр" in op)
            recs.append({
                "date": d, "dep": str(cell(r, i_dep)).strip(),
                "art": str(cell(r, i_art)).strip() or "Прочее",
                "sum": to_number(cell(r, i_sum)), "ctr": str(cell(r, i_ctr)).strip(),
                "pay": pay, "is_op": "займ" not in op, "is_loan_out": is_loan,
            })
    return pd.DataFrame(recs)


# ─── План ───
@st.cache_data(ttl=CACHE_TTL, show_spinner=False)
def get_plan(month):
    headers, rows = load_sheet("План на год")
    plan = {}
    if not rows:
        return plan
    col = month + 1
    for r in rows:
        typ = str(cell(r, 0)).strip().lower()
        name = str(cell(r, 1)).strip()
        if not name or "расход" in typ:
            continue
        plan[name] = plan.get(name, 0) + to_number(cell(r, col))
    return plan


# ─── Займы ───
@st.cache_data(ttl=CACHE_TTL, show_spinner=False)
def get_loans(inc_df, exp_df):
    bal = {}
    if not inc_df.empty:
        for _, r in inc_df[inc_df["is_loan_in"]].iterrows():
            bal[r["ctr"]] = bal.get(r["ctr"], 0) + r["sum"]
    if not exp_df.empty:
        for _, r in exp_df[exp_df["is_loan_out"]].iterrows():
            bal[r["ctr"]] = bal.get(r["ctr"], 0) - r["sum"]
    headers, rows = load_sheet("Архив займов")
    if rows:
        i_ctr  = col_idx(headers, ["контрагент"], 2)
        i_type = col_idx(headers, ["тип"], 3)
        i_sum  = col_idx(headers, ["сумма"], 4)
        for r in rows:
            ctr = str(cell(r, i_ctr)).strip()
            if not ctr:
                continue
            typ = str(cell(r, i_type)).strip().lower()
            s = to_number(cell(r, i_sum))
            if "получ" in typ:
                bal[ctr] = bal.get(ctr, 0) + s
            elif "возвр" in typ:
                bal[ctr] = bal.get(ctr, 0) - s
    return {k: v for k, v in bal.items() if abs(v) > 1}


# ─── Дивиденды (порядок колонок в двух листах разный → ищем по заголовку) ───
@st.cache_data(ttl=CACHE_TTL, show_spinner=False)
def get_dividends():
    recs = []
    for sheet in ["Дивиденды", "Архив дивидендов"]:
        headers, rows = load_sheet(sheet)
        if not rows:
            continue
        i_date = col_idx(headers, ["дата"], 0)
        i_who  = col_idx(headers, ["получател", "учредител", "кому", "имя"], 1)
        i_sum  = col_idx(headers, ["сумма"], 3)
        for r in rows:
            who = str(cell(r, i_who)).strip()
            s = to_number(cell(r, i_sum))
            if not who or s == 0:
                continue
            wl = who.lower()
            if "дарь" in wl:
                who = "Дарья"
            elif "елен" in wl:
                who = "Елена"
            recs.append({"date": to_date(cell(r, i_date)), "who": who, "sum": s})
    return pd.DataFrame(recs)


# ──────────────────────────────────────────────────────────────────────────
# UI-ХЕЛПЕРЫ
# ──────────────────────────────────────────────────────────────────────────
def kpi_card(label, value, color, foot=""):
    st.markdown(f"""
      <div class="card">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value" style="color:{color}">{value}</div>
        <div class="kpi-foot">{foot}</div>
      </div>""", unsafe_allow_html=True)


def plan_color(pct):
    if pct >= 90:
        return C["green"]
    if pct >= 40:
        return C["orange"]
    return C["red"]


def month_filter(df, year, month):
    if df.empty:
        return df
    return df[df["date"].apply(lambda d: d and d.year == year and d.month == month)]


# ──────────────────────────────────────────────────────────────────────────
# ПРИЛОЖЕНИЕ
# ──────────────────────────────────────────────────────────────────────────
def main():
    inc = get_income()
    exp = get_expense()

    if inc.empty and exp.empty:
        st.markdown('<div class="hero-title">📊 Никольская</div>', unsafe_allow_html=True)
        if not access_ok():
            st.error("Таблица пока недоступна (нет публичного доступа).")
            st.markdown(
                "В настройках доступа таблицы поставьте **«Всем, у кого есть ссылка» → Читатель** "
                "(не «Ограниченный» и не «только организация»), сохраните и обновите страницу.")
        else:
            st.warning("Доступ есть, но данные не распознаны. Проверьте, что заголовки в строке 4, "
                       "а данные — с строки 5.")
        return

    months = set()
    for df in (inc, exp):
        if not df.empty:
            for d in df["date"]:
                if d:
                    months.add((d.year, d.month))
    months = sorted(months, reverse=True)
    if not months:
        st.warning("В данных нет распознанных дат.")
        return

    labels = [f"{RU_MONTHS[m]} {y}" for (y, m) in months]
    pick = st.selectbox("Период", labels, index=0, label_visibility="collapsed")
    year, month = months[labels.index(pick)]

    st.markdown(f'<div class="hero-title">Никольская · {RU_MONTHS[month]} {year}</div>'
                f'<div class="hero-sub">Управленческий учёт собственника</div>',
                unsafe_allow_html=True)
    st.markdown('<div style="height:14px"></div>', unsafe_allow_html=True)

    inc_m = month_filter(inc, year, month)
    exp_m = month_filter(exp, year, month)

    tab1, tab2, tab3, tab4 = st.tabs(["Сводка", "Динамика", "Контроль", "Дивиденды"])

    # ═══ СВОДКА ═══
    with tab1:
        op_inc = inc_m[inc_m["is_op"]]["sum"].sum() if not inc_m.empty else 0
        op_exp = exp_m[exp_m["is_op"]]["sum"].sum() if not exp_m.empty else 0
        profit = op_inc - op_exp
        margin = (profit / op_inc * 100) if op_inc else 0
        nal = inc_m[(inc_m["is_op"]) & (inc_m["pay"] == "Наличные")]["sum"].sum() if not inc_m.empty else 0
        beznal = op_inc - nal

        c1, c2, c3 = st.columns(3)
        with c1:
            kpi_card("Прибыль", money(profit),
                     C["green"] if profit >= 0 else C["red"], f"маржа {margin:.1f}%")
        with c2:
            kpi_card("Доход операционный", money(op_inc), C["blue"],
                     f"нал {money(nal)} · безнал {money(beznal)}")
        with c3:
            kpi_card("Расход операционный", money(op_exp), C["orange"], "без займов")

        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown('<div class="kpi-label" style="font-size:16px;color:#1D1D1F;'
                    'font-weight:700;margin-bottom:14px">План / факт по подразделениям</div>',
                    unsafe_allow_html=True)
        plan = get_plan(month)
        fact = (inc_m[inc_m["is_op"]].groupby("div")["sum"].sum().sort_values(ascending=False)
                if not inc_m.empty else pd.Series(dtype=float))
        names = list(dict.fromkeys(list(fact.index) + list(plan.keys())))
        for name in names:
            f = float(fact.get(name, 0))
            p = float(plan.get(name, 0))
            pct = (f / p * 100) if p else (100 if f > 0 else 0)
            col = plan_color(pct)
            st.markdown(f"""
              <div style="margin-bottom:16px">
                <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px">
                  <span class="row-name">{div_icon(name)} {name}</span>
                  <span class="pct-pill" style="background:{col}1A;color:{col}">{pct:.0f}%</span>
                </div>
                <div class="bar-track"><div class="bar-fill" style="width:{min(pct,100)}%;background:{col}"></div></div>
                <div class="row-num" style="margin-top:5px">{money(f)} из {money(p)}</div>
              </div>""", unsafe_allow_html=True)
        if not names:
            st.caption("Нет данных по подразделениям за период.")
        st.markdown('</div>', unsafe_allow_html=True)

    # ═══ ДИНАМИКА ═══
    with tab2:
        if inc_m.empty:
            st.info("Нет доходов за выбранный месяц.")
        else:
            op = inc_m[inc_m["is_op"]].copy()
            top5 = op.groupby("div")["sum"].sum().sort_values(ascending=False).head(5).index.tolist()
            cdf = op[op["div"].isin(top5)].groupby(["date", "div"])["sum"].sum().reset_index()
            cdf["date"] = pd.to_datetime(cdf["date"])
            palette = [C["blue"], C["green"], C["orange"], C["purple"], C["pink"]]

            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown('<div class="kpi-label" style="font-size:16px;color:#1D1D1F;'
                        'font-weight:700;margin-bottom:8px">Доходы по дням · топ-5 подразделений</div>',
                        unsafe_allow_html=True)
            chart = alt.Chart(cdf).mark_line(point=True, strokeWidth=2, interpolate="monotone").encode(
                x=alt.X("date:T", title=None, axis=alt.Axis(format="%d.%m")),
                y=alt.Y("sum:Q", title=None, axis=alt.Axis(format="~s")),
                color=alt.Color("div:N", title=None,
                                scale=alt.Scale(domain=top5, range=palette),
                                legend=alt.Legend(orient="bottom")),
                tooltip=[alt.Tooltip("date:T", title="Дата", format="%d.%m"),
                         alt.Tooltip("div:N", title="Подразделение"),
                         alt.Tooltip("sum:Q", title="Сумма", format=",.0f")],
            ).properties(height=320).configure_view(strokeWidth=0).configure_axis(
                grid=True, gridColor="#E9E9EB", labelColor="#8E8E93")
            st.altair_chart(chart, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

            pm = month - 1 or 12
            py = year if month > 1 else year - 1
            prev = month_filter(inc, py, pm)
            prev_op = prev[prev["is_op"]]["sum"].sum() if not prev.empty else 0
            cur_op = op["sum"].sum()
            delta = cur_op - prev_op
            dpct = (delta / prev_op * 100) if prev_op else 0
            col = C["green"] if delta >= 0 else C["red"]
            arrow = "▲" if delta >= 0 else "▼"
            kpi_card(f"Доход за {RU_MONTHS[month]} vs {RU_MONTHS[pm]}", money(cur_op), col,
                     f"{arrow} {money(abs(delta))} ({dpct:+.0f}%) · было {money(prev_op)}")

    # ═══ КОНТРОЛЬ ═══
    with tab3:
        loans = get_loans(inc, exp)
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown('<div class="kpi-label" style="font-size:16px;color:#1D1D1F;'
                    'font-weight:700;margin-bottom:14px">Займы · баланс по контрагентам</div>',
                    unsafe_allow_html=True)
        if not loans:
            st.caption("Открытых займов нет.")
        else:
            for ctr, bal in sorted(loans.items(), key=lambda x: -abs(x[1])):
                if bal > 0:
                    txt, col = f"клуб должен {money(bal)}", C["red"]
                else:
                    txt, col = f"переплата {money(abs(bal))} (проверьте архив)", C["orange"]
                st.markdown(f"""
                  <div style="display:flex;justify-content:space-between;align-items:center;
                       padding:12px 0;border-bottom:1px solid #F0F0F0">
                    <span class="row-name">{ctr}</span>
                    <span class="pct-pill" style="background:{col}1A;color:{col}">{txt}</span>
                  </div>""", unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="kpi-label" style="font-size:16px;color:#1D1D1F;'
                    'font-weight:700;margin:8px 0 10px 0">Алерты</div>', unsafe_allow_html=True)
        op_inc = inc_m[inc_m["is_op"]]["sum"].sum() if not inc_m.empty else 0
        op_exp = exp_m[exp_m["is_op"]]["sum"].sum() if not exp_m.empty else 0
        profit = op_inc - op_exp
        alerts = []
        if profit < 0:
            alerts.append(("red", f"Убыток за месяц: {money(profit)}"))
        elif op_inc and profit / op_inc < 0.05:
            alerts.append(("orange", f"Низкая маржа: {profit/op_inc*100:.1f}% — прибыль {money(profit)}"))
        # крупный нал без документа
        if not inc_m.empty:
            i_doc = None
            big_cash = inc_m[(inc_m["pay"] == "Наличные") & (inc_m["sum"] >= 100000)]
            if len(big_cash) > 0:
                alerts.append(("orange", f"Крупные наличные операции (≥100к): {len(big_cash)} шт — проверьте первичку"))
        for ctr, bal in loans.items():
            if bal < 0:
                alerts.append(("orange", f"{ctr}: возврат превышает приход на {money(abs(bal))}. "
                                         f"Вероятно, не внесён приход в «Архив займов»."))
        if not alerts:
            alerts.append(("green", "Критичных проблем не обнаружено."))
        for kind, txt in alerts:
            st.markdown(f'<div class="alert alert-{kind}">{txt}</div>', unsafe_allow_html=True)

    # ═══ ДИВИДЕНДЫ ═══
    with tab4:
        div = get_dividends()
        st.markdown('<div class="kpi-label" style="font-size:16px;color:#1D1D1F;'
                    'font-weight:700;margin-bottom:10px">Выплаты учредителям · накопительно</div>',
                    unsafe_allow_html=True)
        totals = {o: 0.0 for o in OWNERS}
        month_pay = {o: 0.0 for o in OWNERS}
        if not div.empty:
            for o in OWNERS:
                totals[o] = div[div["who"] == o]["sum"].sum()
                mm = div[(div["who"] == o) &
                         (div["date"].apply(lambda d: d and d.year == year and d.month == month))]
                month_pay[o] = mm["sum"].sum()
        total_all = sum(totals.values())
        colors = {"Дарья": C["purple"], "Елена": C["pink"]}
        cols = st.columns(2)
        for i, o in enumerate(OWNERS):
            fair = total_all * SHARES[o]
            with cols[i]:
                st.markdown(f"""
                  <div class="card">
                    <div style="display:flex;align-items:center;margin-bottom:8px">
                      <span class="avatar" style="background:{colors[o]}">{o[0]}</span>
                      <span class="row-name">{o} · доля {int(SHARES[o]*100)}%</span>
                    </div>
                    <div class="kpi-value" style="color:{colors[o]};font-size:32px">{money(totals[o])}</div>
                    <div class="kpi-foot">за {RU_MONTHS[month]}: {money(month_pay[o])} · справедливо: {money(fair)}</div>
                  </div>""", unsafe_allow_html=True)

        d_dar, d_el = totals["Дарья"], totals["Елена"]
        gap = abs(d_dar - d_el) / 2
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown('<div class="kpi-label" style="font-size:16px;color:#1D1D1F;'
                    'font-weight:700;margin-bottom:8px">Взаиморасчёт между учредителями</div>',
                    unsafe_allow_html=True)
        if gap < 1:
            st.markdown(f'<div class="alert alert-green">Выплаты сбалансированы 50/50. '
                        f'Всего выплачено {money(total_all)}.</div>', unsafe_allow_html=True)
        else:
            more = "Дарья" if d_dar > d_el else "Елена"
            less = "Елена" if d_dar > d_el else "Дарья"
            st.markdown(f'<div class="alert alert-orange">{more} получила больше. '
                        f'Чтобы выровнять до 50/50, <b>{more} должна {less}: {money(gap)}</b>.<br>'
                        f'Всего выплачено {money(total_all)}.</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
        if div.empty:
            st.caption("Листы «Дивиденды» / «Архив дивидендов» пока пусты.")

    st.markdown(f'<div style="text-align:center;color:#C7C7CC;font-size:12px;margin-top:24px">'
                f'Обновлено {datetime.now().strftime("%d.%m.%Y %H:%M")} · '
                f'кэш {CACHE_TTL//60} мин</div>', unsafe_allow_html=True)


if __name__ == "__main__":
    main()
