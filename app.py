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
TEMPLATE_PATH = "data/templates/Сосновское_clean.xlsx"   # обновите своим шаблоном с колонкой "Пласт"
MIN_POINTS_DEFAULT = 8
SHARED_WATERCUT_THR_DEFAULT = 0.02

st.set_page_config(
    layout="wide",
    initial_sidebar_state="auto",
    page_title="Автодиагностика скважин",
    page_icon="🛢️",
)

# ---------- Утилиты ----------
def _norm(s):
    if not isinstance(s, str): return str(s)
    s = unicodedata.normalize("NFKC", s).replace("\u00A0", " ").replace("\xa0", " ")
    return re.sub(r"\s+", " ", s.strip())

def _to_num(x, fill=None):
    s = pd.to_numeric(x, errors="coerce")
    return s.fillna(fill) if fill is not None else s

def _bytes_of_upload(uploaded_file) -> bytes:
    uploaded_file.seek(0)
    b = uploaded_file.read()
    uploaded_file.seek(0)
    return b

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
    st.write("**Скачать шаблон исходных данных (с колонкой _Пласт_):**")
    out = BytesIO()
    tpl.to_excel(out, index=False, engine="openpyxl")
    out.seek(0)
    st.download_button(
        "Шаблон (XLSX)",
        data=out,
        file_name="template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ---------- Подготовка данных (скважина+пласт = отдельная сущность) ----------
def _build_well_and_layer(df: pd.DataFrame) -> pd.DataFrame:
    """Формируем well, layer, well_id = 'well | layer'. Колонка 'Пласт' обязательна в новом шаблоне."""
    cols = {c.lower(): c for c in df.columns}
    # имена скважины (как раньше: H + I) можно детектить из шаблона, но поддержим и явные:
    well_col = None
    for cand in ["скважина", "скв", "скв.", "well", "id", "номер"]:
        if cand in cols:
            well_col = cols[cand]; break
    if well_col is None:
        # fallback: попробуем старую логику H+I (если есть)
        def col_by_letter(letter: str) -> Optional[str]:
            letter = letter.strip().upper()
            acc = 0
            for ch in letter:
                acc = acc*26 + (ord(ch) - 64)
            i = acc - 1
            return df.columns[i] if 0 <= i < len(df.columns) else None
        cH, cI = col_by_letter("H"), col_by_letter("I")
        if cH:
            well = df[cH].astype(str).fillna("")
            if cI:
                well = (well + " " + df[cI].astype(str).fillna("")).str.strip()
        else:
            # возьмём первый столбец если совсем ничего не нашли
            well = df.iloc[:, 0].astype(str).fillna("")
    else:
        well = df[well_col].astype(str).fillna("")

    # Пласт (обязательно по новому шаблону)
    layer_col = None
    for cand in ["пласт", "пл", "layer", "horizon", "formation"]:
        if cand in cols:
            layer_col = cols[cand]; break
    if layer_col is None:
        # если вдруг нет — создадим заглушку, но предупредим
        st.warning("В шаблоне отсутствует колонка «Пласт». Все точки будут считаться одним пластом.")
        layer = pd.Series("UNK", index=df.index, dtype=str)
    else:
        layer = df[layer_col].astype(str).fillna("").replace({"": "UNK"})

    out = df.copy()
    out["well"] = well.str.strip()
    out["layer"] = layer.str.strip()
    out["well_id"] = (out["well"] + " | " + out["layer"]).str.strip()
    return out

def enforce_monotonic_per_entity(df: pd.DataFrame) -> pd.DataFrame:
    # t_num строго возрастает внутри каждой сущности (well_id)
    g = df.groupby("well_id", sort=False)["t_num"]
    t = g.cummax()
    idx_in_grp = df.groupby("well_id", sort=False).cumcount().to_numpy()
    df = df.copy()
    df["t_num"] = t.to_numpy() + idx_in_grp * EPS
    return df

@st.cache_data
def prepare_data(file_bytes: bytes) -> pd.DataFrame:
    df = pd.read_excel(BytesIO(file_bytes))
    df.columns = [_norm(c) for c in df.columns]
    df = _build_well_and_layer(df)

    # попытка найти числовые поля по шаблону (как и ранее)
    cols = {c.lower(): c for c in df.columns}
    def pick(*names):
        for n in names:
            if n in cols: return cols[n]
        return None

    # жидкость периода и обводненность %
    cX  = pick("дебит жидкости, м3/сут", "жидкость", "x")  # в новом шаблоне может быть по-другому — подправьте имя
    cAB = pick("обводненность, %", "обводненность %", "ab")
    # дни добычи (для пересчёта в сутки)
    cAJ = pick("число дней добычи нефти, сут", "дни добычи", "prod_days", "aj")
    # маркеры серий (как раньше BR), не обязателен
    cBR = pick("серия", "смена режима", "br")

    # если X/AB нет в явном виде — поддержим «X, AB» из прежней таблицы
    if cX is None and "X" in df.columns:  cX = "X"
    if cAB is None and "AB" in df.columns: cAB = "AB"
    # prod_days
    if cAJ is None and "AJ" in df.columns: cAJ = "AJ"

    X_vals  = _to_num(df[cX], 0.0) if cX  else pd.Series(0.0, index=df.index)
    AB_vals = _to_num(df[cAB], 0.0) if cAB else pd.Series(0.0, index=df.index)

    df["qo_period"] = X_vals * (100.0 - AB_vals) / 100.0
    df["qw_period"] = X_vals * AB_vals / 100.0
    df["qL_period"] = df["qo_period"] + df["qw_period"]

    prod_days = _to_num(df[cAJ], 1.0) if cAJ else pd.Series(1.0, index=df.index)
    df["prod_days"] = prod_days

    with np.errstate(divide="ignore", invalid="ignore"):
        df["qo"] = np.where(prod_days > 0, df["qo_period"] / prod_days, np.nan)
        df["qw"] = np.where(prod_days > 0, df["qw_period"] / prod_days, np.nan)
        df["qL"] = np.where(prod_days > 0, df["qL_period"] / prod_days, np.nan)

    # накопленное время (по сериям внутри каждой пары (well, layer))
    if cAJ:
        if cBR:
            br = df[cBR].astype(str).fillna("")
            new_series = (df["well_id"] != df["well_id"].shift()) | (br != br.shift())
            grp = new_series.cumsum()
            df["t_num"] = prod_days.groupby([df["well_id"], grp], sort=False).cumsum()
        else:
            # просто суммируем во времени в рамках well_id
            df["t_num"] = prod_days.groupby(df["well_id"], sort=False).cumsum()
    else:
        df["t_num"] = np.arange(len(df), dtype=float)

    df = df.dropna(subset=["well_id", "t_num"]).sort_values(["well_id", "t_num"]).reset_index(drop=True)
    df = enforce_monotonic_per_entity(df)
    return df

# ---------- Отбор сущностей ----------
@st.cache_data
def select_eligible_entities(df: pd.DataFrame,
                             min_points: int,
                             watercut_thr: float) -> Tuple[str, ...]:
    with np.errstate(divide="ignore", invalid="ignore"):
        fw = df["qw_period"] / df["qL_period"]
    ok = (df["qL_period"] > 0) & (fw > watercut_thr) & (df["prod_days"] > 0)
    cnt = ok.groupby(df["well_id"], sort=False).sum()
    return tuple(cnt.index[cnt >= min_points].sort_values())

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
               eligible_ids: Tuple[str, ...],
               watercut_thr: float,
               min_points: int) -> pd.DataFrame:
    d = df[df["well_id"].isin(eligible_ids)].copy()
    with np.errstate(divide="ignore", invalid="ignore"):
        d["fw"] = d["qw_period"] / d["qL_period"]
    d = d.replace([np.inf, -np.inf], np.nan)
    out = []

    for wid, g in d.groupby("well_id", sort=False):
        g = g.sort_values("t_num")
        idx = np.flatnonzero((g["fw"].to_numpy() > watercut_thr) & (g["qL_period"].to_numpy() > 0) & (g["prod_days"].to_numpy() > 0))
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
                 eligible_ids: Tuple[str, ...],
                 min_points: int) -> pd.DataFrame:
    d = df[df["well_id"].isin(eligible_ids)].copy()
    out = []
    for wid, g in d.groupby("well_id", sort=False):
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
                diagnosis_df: pd.DataFrame) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        diagnosis_df.to_excel(writer, sheet_name="Summary", index=False)
        mg_df.to_excel(writer, sheet_name="MG", index=False)
        chan_df.to_excel(writer, sheet_name="Chan", index=False)
    out.seek(0)
    return out.getvalue()

# ---------- UI ----------
st.markdown(
    """
### Поскважинный автодиагноз (Chan & Меркулова–Гинзбург) — **с поддержкой пласта**

- Теперь каждая **пара** «Скважина + Пласт» анализируется **как отдельная скважина**.  
- Все таблицы, графики и диагнозы строятся для `well_id = "Скважина | Пласт"`.
- Состав допущенных сущностей общий для MG/Chan: требуется ≥ N точек после порога обводнённости.
"""
)

with st.sidebar:
    st.subheader("Параметры допуска")
    water_thr = st.number_input("Порог обводнённости fw", 0.0, 1.0, SHARED_WATERCUT_THR_DEFAULT, 0.01)
    min_pts = st.number_input("Мин. число точек после порога", 3, 200, MIN_POINTS_DEFAULT, 1)
    max_entities_to_plot = st.slider("Сколько сущностей рисовать", 1, 50, 10)
    st.caption("Сущность = (Скважина, Пласт)")

download_template()

uploaded = st.file_uploader("Загрузите файл XLSX/XLS (с колонкой «Пласт»)", type=["xlsx","xls"])
if not uploaded:
    st.info("Загрузите шаблон с вашими данными, чтобы запустить расчёт.")
    st.stop()

file_bytes = _bytes_of_upload(uploaded)

if st.button("▶ Запустить расчёт"):
    with st.spinner("Подготовка данных..."):
        df = prepare_data(file_bytes)

    eligible_ids = select_eligible_entities(df, min_points=min_pts, watercut_thr=water_thr)
    if not eligible_ids:
        st.warning("Нет допущенных пар (Скважина, Пласт) — проверьте данные и порог.")
        st.stop()

    prog = st.progress(0, text="Расчёт MG…")
    mg_df = compute_mg(df, eligible_ids, water_thr, min_pts); prog.progress(50, text="Расчёт Chan…")
    chan_df = compute_chan(df, eligible_ids, min_pts);        prog.progress(100, text="Готово")

    # сводка по сущностям
    rows = []
    for wid in eligible_ids:
        mg_g = mg_df[mg_df["well_id"] == wid]
        ch_g = chan_df[chan_df["well_id"] == wid]
        rows.append({"well_id": wid, **diag_mg(mg_g), **diag_chan(ch_g)})
    diagnosis_df = pd.DataFrame(rows)

    st.success(f"Готово: сущностей (скважина|пласт) — {len(eligible_ids)}")
    st.subheader("Сводная таблица диагнозов")
    st.dataframe(diagnosis_df, use_container_width=True)

    st.subheader("Графики (ограничен списоком для скорости)")
    to_show = eligible_ids[:max_entities_to_plot]
    cols = st.columns(2)
    for i, wid in enumerate(to_show):
        with st.expander(f"{wid} — графики"):
            mg_g = mg_df[mg_df["well_id"] == wid]
            ch_g = chan_df[chan_df["well_id"] == wid]

            with cols[i % 2]:
                if not mg_g.empty:
                    fig, ax = plt.subplots()
                    ax.scatter(mg_g["MG_X"], mg_g["MG_Y"], s=10)
                    ax.grid(True, alpha=0.3)
                    ax.set_xlabel("X = Qt_cum/Qt_cum(T)")
                    ax.set_ylabel("Y = Qo_cum/Qt_cum")
                    ax.set_title(f"MG — {wid}")
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
                    ax2.set_title(f"Chan — {wid}")
                    st.pyplot(fig2)
                else:
                    st.info("Нет данных Chan")

    st.subheader("Скачать результаты")
    xlsx_bytes = export_xlsx(mg_df, chan_df, diagnosis_df)
    st.download_button(
        "📥 Единый Excel (Summary, MG, Chan) по (Скважина|Пласт)",
        data=xlsx_bytes,
        file_name="Autodiagnostics_results_by_layer.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Интерфейс готов. Нажмите «Запустить расчёт».")
