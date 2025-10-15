# -*- coding: utf-8 -*-
from __future__ import annotations

import os, re, unicodedata
from dataclasses import dataclass
from io import BytesIO
from typing import Dict, List, Optional, Set, Tuple

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st

# ========= Константы / настройки =========
EPS = 1e-9
TEMPLATE_PATH = "data/templates/Сосновское_clean.xlsx"
MIN_POINTS = 8
SHARED_WATERCUT_THR = 0.02

st.set_page_config(
    layout="wide",
    initial_sidebar_state="auto",
    page_title="Автодиагностика скважин",
    page_icon="🛢️",
)

# ---------- Вспомогательные ----------
def _norm(s):
    if not isinstance(s, str): return str(s)
    s = unicodedata.normalize("NFKC", s).replace("\u00A0", " ").replace("\xa0", " ")
    return re.sub(r"\s+", " ", s.strip())

def _to_num(x, fill=None):
    s = pd.to_numeric(x, errors="coerce")
    return s.fillna(fill) if fill is not None else s

def _bytes_of_upload(uploaded_file) -> bytes:
    # нужно для корректного кэширования по содержимому
    uploaded_file.seek(0)
    b = uploaded_file.read()
    uploaded_file.seek(0)
    return b

# ---------- Шаблон ----------
@st.cache_data
def read_template_df() -> pd.DataFrame:
    try:
        if os.path.exists(TEMPLATE_PATH):
            return pd.read_excel(TEMPLATE_PATH)
    except Exception:
        pass
    return pd.DataFrame()

def download_template():
    tpl = read_template_df()
    st.write("**Скачать шаблон исходных данных:**")
    out = BytesIO()
    tpl.to_excel(out, index=False, engine="openpyxl")
    out.seek(0)
    st.download_button(
        "Шаблон (XLSX)",
        data=out,
        file_name="template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ---------- Подготовка данных ----------
def enforce_monotonic_per_well_fast(df: pd.DataFrame) -> pd.DataFrame:
    # t_num монотонен за счёт cummax + маленький шаг
    g = df.groupby("well", sort=False)["t_num"]
    t = g.cummax()
    # добавим очень маленький прираст, чтобы равные превращались в строго возрастающие
    # за счёт порядкового номера внутри группы
    idx_in_grp = df.groupby("well", sort=False).cumcount().to_numpy()
    df = df.copy()
    df["t_num"] = t.to_numpy() + idx_in_grp * EPS
    return df

@st.cache_data
def prepare_data(file_bytes: bytes) -> pd.DataFrame:
    df = pd.read_excel(BytesIO(file_bytes))
    df.columns = [_norm(c) for c in df.columns]

    # автодетект имён, близко к вашей логике
    # допустим: H+I -> well; X -> дебит жидкости, AB -> обводнённость %
    def col(letter):
        # безопасно получить колонку по индексу «A..Z»
        letter = letter.strip().upper()
        acc = 0
        for ch in letter:
            acc = acc*26 + (ord(ch) - 64)
        i = acc - 1
        return df.columns[i] if 0 <= i < len(df.columns) else None

    def s(letter):
        c = col(letter)
        return df[c] if c in df.columns else None

    sH, sI, sX, sAB, sBT, sBS, sBR, sAJ = s("H"), s("I"), s("X"), s("AB"), s("BT"), s("BS"), s("BR"), s("AJ")

    # well
    well = (sH.astype(str).fillna("") if sH is not None else "").astype(str)
    if sI is not None:
        well = (well + " " + sI.astype(str).fillna("")).str.strip()
    df["well"] = well if isinstance(well, pd.Series) else pd.Series("", index=df.index, dtype=str)

    # дебиты периода
    X_vals = _to_num(sX, fill=0.0)
    AB_vals = _to_num(sAB, fill=0.0)
    df["qo_period"] = X_vals * (100.0 - AB_vals) / 100.0
    df["qw_period"] = X_vals * AB_vals / 100.0
    df["qL_period"] = df["qo_period"] + df["qw_period"]

    # ВНФ = BT/BS, ВНФ' по накопленному времени (AJ)
    df["prod_days"] = _to_num(sAJ, fill=0.0)
    df["qo"] = np.where(df["prod_days"] > 0, df["qo_period"] / df["prod_days"], np.nan)
    df["qw"] = np.where(df["prod_days"] > 0, df["qw_period"] / df["prod_days"], np.nan)
    df["qL"] = np.where(df["prod_days"] > 0, df["qL_period"] / df["prod_days"], np.nan)

    # Накопленное время работы — берём AJ и наращиваем внутри «серий»
    if (sBR is not None) and (sAJ is not None):
        br = sBR.astype(str).fillna("")
        aj = _to_num(sAJ, fill=0.0)
        new_series = (df["well"] != df["well"].shift()) | (br != br.shift())
        grp = new_series.cumsum()
        df["t_num"] = aj.groupby([df["well"], grp], sort=False).cumsum()
    else:
        df["t_num"] = _to_num(sAJ, fill=0.0)

    df = df.dropna(subset=["well", "t_num"]).sort_values(["well", "t_num"]).reset_index(drop=True)
    df = enforce_monotonic_per_well_fast(df)
    return df

# ---------- Отбор скважин ----------
@st.cache_data
def select_eligible_wells(df: pd.DataFrame,
                          min_points: int,
                          watercut_thr: float) -> Set[str]:
    with np.errstate(divide="ignore", invalid="ignore"):
        fw = df["qw_period"] / df["qL_period"]
    ok = (df["qL_period"] > 0) & (fw > watercut_thr) & (df["prod_days"] > 0)
    # посчитаем число валидных точек после порога по скважине
    cnt = ok.groupby(df["well"], sort=False).sum()
    return set(cnt.index[cnt >= min_points])

# ---------- MG ----------
@dataclass
class MGFlags:
    y_early_mean: Optional[float] = None
    slope_first_third: Optional[float] = None
    waviness_std: Optional[float] = None
    possible_behind_casing: bool = False
    possible_channeling: bool = False
    possible_mixed_causes: bool = False

@st.cache_data
def compute_mg(df: pd.DataFrame,
               eligible_wells: Tuple[str, ...],
               watercut_thr: float,
               min_points: int) -> pd.DataFrame:
    d = df[df["well"].isin(eligible_wells)].copy()
    with np.errstate(divide="ignore", invalid="ignore"):
        d["fw"] = d["qw_period"] / d["qL_period"]
    d = d.replace([np.inf, -np.inf], np.nan)

    out = []
    for w, g in d.groupby("well", sort=False):
        g = g.sort_values("t_num")
        # отсекаем до первого fw > thr
        idx = np.flatnonzero(g["fw"].to_numpy() > watercut_thr)
        if idx.size == 0: 
            continue
        g2 = g.iloc[idx[0]:].copy()
        g2["Qo_cum"] = g2["qo_period"].cumsum()
        g2["Qw_cum"] = g2["qw_period"].cumsum()
        g2["Qt_cum"] = g2["Qo_cum"] + g2["Qw_cum"]
        if (len(g2) < min_points) or (float(g2["Qt_cum"].iloc[-1]) <= 0):
            continue

        Qt_T = float(g2["Qt_cum"].iloc[-1])
        X = (g2["Qt_cum"] / Qt_T).to_numpy()
        X = np.maximum.accumulate(X) + np.arange(len(X)) * EPS
        g2["MG_X"] = X
        with np.errstate(divide="ignore", invalid="ignore"):
            g2["MG_Y"] = g2["Qo_cum"] / g2["Qt_cum"]

        flags = MGFlags()
        early = g2["MG_X"] <= 0.2
        if early.sum() >= 3:
            flags.y_early_mean = float(np.nanmean(g2.loc[early, "MG_Y"]))
            flags.possible_behind_casing = (flags.y_early_mean is not None) and (flags.y_early_mean >= 0.99)

        first_third = g2[g2["MG_X"] <= 0.33]
        if len(first_third) >= 3:
            try:
                k, _ = np.polyfit(first_third["MG_X"], first_third["MG_Y"], 1)
                flags.slope_first_third = float(k)
                flags.possible_channeling = (k < -0.8)
            except np.linalg.LinAlgError:
                pass

        if len(g2) >= 5:
            with np.errstate(invalid="ignore"):
                dy = np.gradient(g2["MG_Y"].to_numpy(), g2["MG_X"].to_numpy())
            flags.waviness_std = float(np.nanstd(dy))
            flags.possible_mixed_causes = flags.waviness_std > 1.0

        for k, v in vars(flags).items():
            g2[f"MG_diag_{k}"] = v

        out.append(g2)

    return pd.concat(out, axis=0).reset_index(drop=True) if out else pd.DataFrame()

# ---------- Chan ----------
@dataclass
class ChanFlags:
    slope_logWOR_logt: Optional[float] = None
    mean_derivative: Optional[float] = None
    std_derivative: Optional[float] = None
    possible_coning: bool = False
    possible_near_wellbore: bool = False
    possible_multilayer_channeling: bool = False

@st.cache_data
def compute_chan(df: pd.DataFrame,
                 eligible_wells: Tuple[str, ...],
                 min_points: int) -> pd.DataFrame:
    d = df[df["well"].isin(eligible_wells)].copy()
    out = []
    for w, g in d.groupby("well", sort=False):
        g = g.sort_values("t_num").copy()
        with np.errstate(divide="ignore", invalid="ignore"):
            g["WOR"] = g["qw"] / g["qo"]
        g = g.replace([np.inf, -np.inf], np.nan)
        g = g[(g["qo"] > 0) & (g["WOR"] > 0)].dropna(subset=["WOR"])
        if len(g) < min_points:
            continue

        g["t_pos"] = g["t_num"] - g["t_num"].min() + EPS
        with np.errstate(invalid="ignore"):
            g["dWOR_dt"] = np.gradient(g["WOR"].to_numpy(), g["t_pos"].to_numpy())
        g["dWOR_dt_pos"] = np.where(g["dWOR_dt"] > 0, g["dWOR_dt"], np.nan)

        mask = (g["WOR"] > 0) & (g["t_pos"] > 0)
        a = np.nan
        if mask.sum() >= 3:
            x = np.log(g.loc[mask, "t_pos"].to_numpy())
            y = np.log(g.loc[mask, "WOR"].to_numpy())
            try:
                a, _ = np.polyfit(x, y, 1)
            except np.linalg.LinAlgError:
                pass

        flags = ChanFlags()
        flags.slope_logWOR_logt = float(a)
        flags.mean_derivative = float(np.nanmean(g["dWOR_dt"]))
        flags.std_derivative = float(np.nanstd(g["dWOR_dt"]))
        if not np.isnan(a):
            flags.possible_coning = a > 0.5 and flags.mean_derivative > 0
            flags.possible_near_wellbore = a > 1.0 and flags.mean_derivative > 0
            flags.possible_multilayer_channeling = a > 0 and flags.std_derivative > 0.1

        for k, v in vars(flags).items():
            g[f"chan_diag_{k}"] = v
        out.append(g)

    return pd.concat(out, axis=0).reset_index(drop=True) if out else pd.DataFrame()

# ---------- Диагнозы ----------
def diag_mg(g: pd.DataFrame) -> Dict[str, str]:
    if g.empty: return {"mg_text":"нет данных MG","mg_detail":""}
    r = g.iloc[-1]
    parts = []
    if r.get("MG_diag_possible_behind_casing"): parts.append("возможны заколонные перетоки (ранний нефтеотбор Y≈1)")
    if r.get("MG_diag_possible_channeling"):    parts.append("признаки каналирования (крутой спад Y в первой трети)")
    if r.get("MG_diag_possible_mixed_causes"):  parts.append("смешанные причины (высокая волнистость dY/dX)")
    if not parts: parts.append("ближе к равномерному обводнению")
    y = r.get("MG_diag_y_early_mean", np.nan)
    k = r.get("MG_diag_slope_first_third", np.nan)
    w = r.get("MG_diag_waviness_std", np.nan)
    return {"mg_text":"; ".join(parts), "mg_detail":f"MG метрики: y_early≈{y:.2f}; наклон≈{k:.2f}; волнистость≈{w:.2f}"}

def diag_chan(g: pd.DataFrame) -> Dict[str, str]:
    if g.empty: return {"chan_text":"нет данных Chan","chan_detail":""}
    r = g.iloc[-1]
    parts = []
    if r.get("chan_diag_possible_multilayer_channeling"): parts.append("многослойное каналирование (рост WOR и дисперсии производной)")
    if r.get("chan_diag_possible_near_wellbore"):         parts.append("приствольные проблемы/ранний канал (очень высокий наклон)")
    if r.get("chan_diag_possible_coning"):                parts.append("возможен конинг (наклон > 0.5 при положительной производной)")
    if not parts: parts.append("нет выраженных признаков проблемного притока воды")
    a = r.get("chan_diag_slope_logWOR_logt", np.nan)
    m = r.get("chan_diag_mean_derivative", np.nan)
    s = r.get("chan_diag_std_derivative", np.nan)
    return {"chan_text":"; ".join(parts), "chan_detail":f"Chan метрики: наклон≈{a:.2f}; средн. dWOR/dt≈{m:.2e}; std≈{s:.2e}"}

# ---------- Экспорт ----------
@st.cache_data
def export_xlsx(mg_df: pd.DataFrame,
                chan_df: pd.DataFrame,
                diagnosis_df: pd.DataFrame,
                include_charts: bool) -> bytes:
    out = BytesIO()
    engine = "xlsxwriter" if include_charts else "openpyxl"
    with pd.ExcelWriter(out, engine=engine) as writer:
        diagnosis_df.to_excel(writer, sheet_name="Summary", index=False)
        mg_df.to_excel(writer, sheet_name="MG", index=False)
        chan_df.to_excel(writer, sheet_name="Chan", index=False)

        if include_charts and hasattr(writer, "book") and not mg_df.empty:
            wb = writer.book
            ws_mg = wb.add_worksheet("MG_plots"); writer.sheets["MG_plots"] = ws_mg
            row = 0
            for w, g in mg_df.groupby("well", sort=False):
                g0 = g.reset_index(drop=True)
                g0.to_excel(writer, sheet_name="MG_plots", index=False, startrow=row)
                n = len(g0)
                if n >= 3 and "MG_X" in g0 and "MG_Y" in g0:
                    cx = g0.columns.get_loc("MG_X")+1
                    cy = g0.columns.get_loc("MG_Y")+1
                    chart = wb.add_chart({'type':'scatter'})
                    chart.add_series({
                        'name': f'{w}',
                        'categories': ['MG_plots', row+1, cx, row+n, cx],
                        'values':     ['MG_plots', row+1, cy, row+n, cy],
                        'marker': {'type':'circle','size':4},
                    })
                    chart.set_title({'name': f'MG — {w}'})
                    chart.set_x_axis({'name':'X'}); chart.set_y_axis({'name':'Y'})
                    ws_mg.insert_chart(row, g0.shape[1]+2, chart)
                row += n + 4
    out.seek(0)
    return out.getvalue()

# ---------- UI ----------
st.markdown(
    """
### Поскважинный автодиагноз (Chan & Меркулова–Гинзбург)

**Что делает приложение:** считает признаки механизма обводнения по двум подходам и формирует диагнозы, графики и единый Excel.

**Фильтрация и одинаковый состав скважин:**
- MG использует кумулятивные объёмы и старт после достижения порога обводнённости `fw > {thr:.2f}`; нужно ≥{n} точек *после* порога.  
- Chan — лог–лог анализ `WOR(t)` и `dWOR/dt`; внутри скважины чистятся только точки, **список скважин общий** (как для MG).
""".format(thr=SHARED_WATERCUT_THR, n=MIN_POINTS)
)

with st.sidebar:
    st.subheader("Параметры")
    water_thr = st.number_input("Порог обводнённости fw для допуска (общий)", 0.0, 1.0, SHARED_WATERCUT_THR, 0.01)
    min_pts = st.number_input("Мин. число точек после порога (общий)", 3, 200, MIN_POINTS, 1)
    max_wells_to_plot = st.slider("Сколько скважин рисовать (для скорости)", 1, 50, 10)
    include_charts_in_excel = st.checkbox("Встраивать графики в Excel (медленнее)", value=False)
    st.caption("Совет: оставьте графики в Excel выключенными — файл сформируется значительно быстрее.")

download_template()

uploaded = st.file_uploader("Загрузите файл XLSX/XLS", type=["xlsx","xls"])
if not uploaded:
    st.info("Загрузите шаблон с вашими данными, чтобы запустить расчёт.")
    st.stop()

file_bytes = _bytes_of_upload(uploaded)

# «Быстрый UI»: ничего тяжёлого до нажатия кнопки
if st.button("▶ Запустить расчёт"):
    # 1) Подготовка
    with st.spinner("Подготовка данных..."):
        df = prepare_data(file_bytes)

    # 2) Отбор скважин
    eligible = select_eligible_wells(df, min_points=min_pts, watercut_thr=water_thr)
    if not eligible:
        st.warning("Не найдено скважин, удовлетворяющих критериям допуска. Проверьте данные/порог.")
        st.stop()

    wells_tuple = tuple(sorted(eligible))

    # 3) Расчёты (кэшируются по содержимому файла и параметрам)
    prog = st.progress(0, text="Расчёт MG...")
    mg_df = compute_mg(df, wells_tuple, water_thr, min_pts); prog.progress(50, text="Расчёт Chan...")
    chan_df = compute_chan(df, wells_tuple, min_pts);        prog.progress(100, text="Готово")

    # 4) Диагнозы
    rows = []
    for w in wells_tuple:
        mg_g = mg_df[mg_df["well"] == w]
        ch_g = chan_df[chan_df["well"] == w]
        rows.append({"well": w, **diag_mg(mg_g), **diag_chan(ch_g)})
    diagnosis_df = pd.DataFrame(rows)

    st.success(f"Готово: скважин в анализе — {len(wells_tuple)}")

    # 5) Отображение (ограничиваем число графиков)
    st.subheader("Сводная таблица диагнозов")
    st.dataframe(diagnosis_df, use_container_width=True)

    st.subheader("Графики (ограничен списоком для скорости)")
    to_show = wells_tuple[:max_wells_to_plot]
    cols = st.columns(2)
    for i, w in enumerate(to_show):
        with st.expander(f"Скважина {w} — графики"):
            mg_g = mg_df[mg_df["well"] == w]
            ch_g = chan_df[chan_df["well"] == w]

            with cols[i % 2]:
                if not mg_g.empty:
                    fig, ax = plt.subplots()
                    ax.scatter(mg_g["MG_X"], mg_g["MG_Y"], s=10)
                    ax.grid(True, alpha=0.3)
                    ax.set_xlabel("X = Qt_cum/Qt_cum(T)")
                    ax.set_ylabel("Y = Qo_cum/Qt_cum")
                    ax.set_title(f"MG — {w}")
                    st.pyplot(fig)
                else:
                    st.info("Нет данных MG")

                if not ch_g.empty:
                    fig2, ax2 = plt.subplots()
                    ax2.plot(ch_g["t_pos"], ch_g["WOR"], "o", markersize=3, label="WOR")
                    ax2.plot(ch_g["t_pos"], ch_g["dWOR_dt_pos"], "--", label="|dWOR/dt|")
                    ax2.set_xscale("log"); ax2.set_yscale("log")
                    ax2.grid(True, which="both", alpha=0.3); ax2.legend()
                    ax2.set_xlabel("t_pos (дни)"); ax2.set_ylabel("WOR, |dWOR/dt|")
                    ax2.set_title(f"Chan — {w}")
                    st.pyplot(fig2)
                else:
                    st.info("Нет данных Chan")

    # 6) Экспорт (по умолчанию без графиков — быстро)
    st.subheader("Скачать результаты")
    xlsx_bytes = export_xlsx(mg_df, chan_df, diagnosis_df, include_charts=include_charts_in_excel)
    st.download_button(
        "📥 Единый Excel (Summary, MG, Chan{charts})".format(charts=", +графики" if include_charts_in_excel else ""),
        data=xlsx_bytes,
        file_name="Autodiagnostics_results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Нажмите «Запустить расчёт» — интерфейс уже готов, данные будут кэшироваться для повторных запусков.")
