# -*- coding: utf-8 -*-
# Автодиагностика скважин: модульная версия
# Блоки: 0) Конфиг/импорты, 1) Утилиты, 2) Работа с данными (I/O),
# 3) Подготовка/валидация, 4) Расчёты (MG/Chan), 5) Формирование выводов,
# 6) Экспорт результатов, 7) Интерфейс (Streamlit UI).
# Логика анализа ведётся по виртуальному ключу well_id = "Скважина | Пласт".

from __future__ import annotations

# ---------------------------- 0) КОНФИГ/ИМПОРТЫ -----------------------------
import io, re, unicodedata
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st

EPS = 1e-9
MIN_POINTS_DEFAULT = 6
SHARED_WATERCUT_THR_DEFAULT = 0.02
APP_TITLE = "Поскважинный автодиагноз (Chan & Меркулова–Гинзбург) — учёт Скважины и Пласта"

st.set_page_config(layout="wide", page_title="Автодиагностика скважин", page_icon="🛢️")


# ------------------------------- 1) УТИЛИТЫ ---------------------------------
def _drop_unnamed(df: pd.DataFrame) -> pd.DataFrame:
    """Убирает служебные столбцы вида 'Unnamed:*' (последствия сохранений Excel)."""
    return df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]

def _bytes_of_upload(f) -> bytes:
    """Читает bytes из загруженного файла и возвращает указатель в начало."""
    f.seek(0); b = f.read(); f.seek(0); return b

def _norm_text(s: str) -> str:
    """Нормализация текста: убрать регистр/пробелы/знаки, заменить 'ё'."""
    s = str(s).replace("ё","е").replace("Ё","Е")
    s = unicodedata.normalize("NFKC", s).lower()
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"[^\w]+", "", s)
    return s

def _norm_cols(df: pd.DataFrame) -> Dict[str, str]:
    """Карта {нормализованное_имя: оригинал} для гибкого поиска колонок."""
    return {_norm_text(c): c for c in df.columns}

def _find_col(normmap: Dict[str, str], variants: List[str]) -> Optional[str]:
    """Ищет колонку по точному совпадению нормализованных имён или по подстроке."""
    for v in variants:
        if v in normmap: 
            return normmap[v]
    for k, orig in normmap.items():
        if any(v in k for v in variants):
            return orig
    return None

def _to_num_series(s: pd.Series, fill=None) -> pd.Series:
    """Переводит текстовые числа в float, аккуратно чистит пробелы/запятые."""
    txt = (
        s.astype(str)
         .str.replace("\u00A0", " ")
         .str.replace(" ", "", regex=False)
         .str.replace(",", ".", regex=False)
    )
    out = pd.to_numeric(txt, errors="coerce")
    return out.fillna(fill) if fill is not None else out

def _fmt(x, fmt="{:.2f}") -> str:
    """Безопасное форматирование чисел для текстовых подсказок."""
    try:
        if x is None or (isinstance(x, float) and np.isnan(x)):
            return "—"
        return fmt.format(float(x))
    except Exception:
        return "—"

def _safe_slice_by_well_id(df: pd.DataFrame, wid: str) -> pd.DataFrame:
    """Безопасный срез по well_id (не падает, если столбца нет/пусто)."""
    if df is None or df.empty or "well_id" not in df.columns:
        return pd.DataFrame()
    return df[df["well_id"] == wid]


# ------------------------- 2) РАБОТА С ДАННЫМИ (I/O) ------------------------
@st.cache_data
def read_user_file(file_bytes: bytes, filename: str) -> pd.DataFrame:
    """
    Универсальное чтение исходника:
    - .xlsx/.xls — как Excel;
    - .csv — авто-определение разделителя и перебор кодировок.
    """
    name = (filename or "").lower()
    if name.endswith((".xlsx", ".xls")):
        df = pd.read_excel(io.BytesIO(file_bytes))
        return _drop_unnamed(df)

    last_exc = None
    for enc in ("utf-8", "utf-8-sig", "cp1251", "latin1"):
        try:
            df = pd.read_csv(io.BytesIO(file_bytes), sep=None, engine="python",
                             encoding=enc, on_bad_lines="skip")
            return _drop_unnamed(df)
        except Exception as e:
            last_exc = e
    raise ValueError(f"Не удалось прочитать файл: {last_exc}")

def make_templates() -> Dict[str, pd.DataFrame]:
    """Мини-набор шаблонов (для кнопок скачивания)."""
    liq_wc = pd.DataFrame({
        "Скважина": ["113","113","115","115"],
        "Пласт":    ["Б6","Б6","Бш","Бш"],
        "Дата":     pd.to_datetime(["2024-01-01","2024-02-01","2024-01-01","2024-02-01"]),
        "Дни добычи": [31,29,31,29],
        "Жидкость м3/сут": [100,120,95,110],
        "Обводненность %": [10,22,5,18],
    })
    qo_qw = pd.DataFrame({
        "Скважина": ["A-01","A-01","A-02","A-02"],
        "Пласт":    ["Ю1","Ю1","Ю2","Ю2"],
        "Дата":     pd.to_datetime(["2024-03-01","2024-04-01","2024-03-01","2024-04-01"]),
        "Дни добычи": [31,30,31,30],
        "Дебит нефти, м3/сут": [80,70,60,50],
        "Дебит воды, м3/сут":  [20,40,10,25],
    })
    return {"liq_wc": liq_wc, "qo_qw": qo_qw}

def df_to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    bio = io.BytesIO()
    df.to_excel(bio, index=False, engine="openpyxl")
    bio.seek(0)
    return bio.getvalue()

def df_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")


# ------------------ 3) ПОДГОТОВКА/ВАЛИДАЦИЯ И ИДЕНТИФИКАТОРЫ ----------------
def ensure_well_layer(df: pd.DataFrame) -> pd.DataFrame:
    """
    Гарантирует наличие колонок 'Скважина' и 'Пласт'. Если нет — создаёт.
    Формирует виртуальный идентификатор well_id = 'Скважина | Пласт'.
    """
    nm = _norm_cols(df)
    well_col  = _find_col(nm, ["скважина","скв","well","id","номер"])
    layer_col = _find_col(nm, ["пласт","пл","layer","horizon","formation"])

    out = df.copy()
    if well_col is None:
        out["Скважина"] = [f"WELL_{i+1}" for i in range(len(out))]
        well_col = "Скважина"
    if layer_col is None:
        out["Пласт"] = "UNK"
        layer_col = "Пласт"

    out["Скважина"] = out[well_col].astype(str).fillna("").str.strip()
    out["Пласт"]    = out[layer_col].astype(str).fillna("").replace({"": "UNK"}).str.strip()
    out["well_id"]  = (out["Скважина"] + " | " + out["Пласт"]).str.strip()
    return out

def enforce_monotonic_time(df: pd.DataFrame) -> pd.DataFrame:
    """Делает время строго неубывающим внутри каждой сущности well_id."""
    g = df.groupby("well_id", sort=False)["t_num"]
    t = g.cummax()
    idx = df.groupby("well_id", sort=False).cumcount().to_numpy()
    out = df.copy()
    out["t_num"] = t.to_numpy() + idx * EPS
    return out

@st.cache_data
def prepare_data(file_bytes: bytes, filename: str) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """
    Основная унификация входов:
    - Переводим всё в периодные объёмы qo_period/qw_period/qL_period;
    - При наличии суточных значений — умножаем на 'Дни добычи' (если нет — берём 1).
    - t_num: накопленное время (по дате, накопленному времени или сумме дней).
    Возвращает (df, detector) — где detector описывает распознанные поля.
    """
    raw = read_user_file(file_bytes, filename)
    raw.columns = [str(c) for c in raw.columns]
    df = ensure_well_layer(raw)
    nm = _norm_cols(df)
    detector: Dict[str, str] = {}

    # ---- Поиск времени ----
    c_date = _find_col(nm, ["дата","месяц","period","monthyear"])
    c_tcum = _find_col(nm, ["накопленноевремиработы","накопленноевремени","накопленноевремяработы","taccum"])
    c_days = _find_col(nm, ["днидобычи","числоднейдобычинефтисут","proddays","aj","сут"])

    # ---- Поиск дебитов/обводнённости ----
    c_qoP   = _find_col(nm, ["добычанефтим3мес","qoperiod"])
    c_qwP   = _find_col(nm, ["добычаводым3мес","qwperiod"])
    c_qL_P  = _find_col(nm, ["добыкажидкостим3мес","жидкостьм3мес","qlperiod"])

    c_liq   = _find_col(nm, ["жидкостьм3сут","дебитжидкостим3сут","liquid","ql"])
    c_wc    = _find_col(nm, ["обводненность","обводненностьпроцент","watercut","wc","ab"])
    c_qo    = _find_col(nm, ["дебитнефтим3сут","qo","oilrate"])
    c_qw    = _find_col(nm, ["дебитводым3сут","qw","waterrate"])

    # ---- Ветвление по источникам ----
    if c_qoP and c_qwP:
        df["qo_period"] = _to_num_series(df[c_qoP], 0.0); detector["qo_period"] = c_qoP
        df["qw_period"] = _to_num_series(df[c_qwP], 0.0); detector["qw_period"] = c_qwP
        df["qL_period"] = df["qo_period"] + df["qw_period"]
        df["qo"] = np.nan; df["qw"] = np.nan; df["qL"] = np.nan

    elif c_qL_P and c_wc:
        qL = _to_num_series(df[c_qL_P], 0.0); detector["qL_period"] = c_qL_P
        wc = _to_num_series(df[c_wc],   0.0); detector["watercut"]  = c_wc
        df["qo_period"] = qL * (100.0 - wc) / 100.0
        df["qw_period"] = qL * wc / 100.0
        df["qL_period"] = qL
        df["qo"] = np.nan; df["qw"] = np.nan; df["qL"] = np.nan

    elif c_liq and c_wc:
        liq = _to_num_series(df[c_liq], 0.0); detector["liq"]      = c_liq
        wc  = _to_num_series(df[c_wc],  0.0); detector["watercut"] = c_wc
        df["qo"] = liq * (100.0 - wc) / 100.0
        df["qw"] = liq * wc / 100.0
        df["qL"] = liq
        days = _to_num_series(df[c_days], 1.0) if c_days else pd.Series(1.0, index=df.index)
        if c_days: detector["prod_days"] = c_days
        df["prod_days"] = days
        df["qo_period"] = df["qo"] * days
        df["qw_period"] = df["qw"] * days
        df["qL_period"] = df["qL"] * days

    elif c_qo and c_qw:
        df["qo"] = _to_num_series(df[c_qo], 0.0); detector["qo"] = c_qo
        df["qw"] = _to_num_series(df[c_qw], 0.0); detector["qw"] = c_qw
        df["qL"] = df["qo"] + df["qw"]
        days = _to_num_series(df[c_days], 1.0) if c_days else pd.Series(1.0, index=df.index)
        if c_days: detector["prod_days"] = c_days
        df["prod_days"] = days
        df["qo_period"] = df["qo"] * days
        df["qw_period"] = df["qw"] * days
        df["qL_period"] = df["qL"] * days

    else:
        # Фолбэк — нули (не валимся, интерфейс работает)
        days = _to_num_series(df[c_days], 1.0) if c_days else pd.Series(1.0, index=df.index)
        if c_days: detector["prod_days"] = c_days
        df["prod_days"] = days
        df["qo_period"] = pd.Series(0.0, index=df.index)
        df["qw_period"] = pd.Series(0.0, index=df.index)
        df["qL_period"] = pd.Series(0.0, index=df.index)
        df["qo"] = np.nan; df["qw"] = np.nan; df["qL"] = np.nan

    # ---- Время t_num ----
    if c_tcum:
        df["t_num"] = _to_num_series(df[c_tcum], 0.0); detector["t_num"] = c_tcum
    elif c_date:
        t = pd.to_datetime(df[c_date], errors="coerce"); detector["date"] = c_date
        df["t_num"] = (t - t.groupby(df["well_id"]).transform("min")).dt.days.astype(float).fillna(0.0)
    else:
        df["t_num"] = df.get("prod_days", pd.Series(1.0, index=df.index)).groupby(df["well_id"]).cumsum()
        detector["t_num"] = "cum(prod_days)"

    df = df.dropna(subset=["well_id", "t_num"]).sort_values(["well_id", "t_num"]).reset_index(drop=True)
    df = enforce_monotonic_time(df)
    return df, detector

@st.cache_data
def select_eligible_with_reasons(df: pd.DataFrame, min_points: int, watercut_thr: float) -> Tuple[Tuple[str, ...], Dict[str, str]]:
    """
    Выбирает сущности (well_id), где точек ПОСЛЕ достижения порога fw>w_thr — ≥ min_points.
    Возвращает (tuple(ids), reasons) — причины недопуска для остальных.
    """
    reasons: Dict[str, str] = {}
    if df.empty:
        return tuple(), reasons

    with np.errstate(divide="ignore", invalid="ignore"):
        fw = df["qw_period"] / df["qL_period"]
    cond = (df["qL_period"] > 0) & (fw > watercut_thr) & (df["prod_days"] > 0)

    eligible: List[str] = []
    for wid, g in df.groupby("well_id", sort=False):
        ok = int(cond[g.index].sum())
        if ok >= min_points:
            eligible.append(wid)
        else:
            parts = []
            if (g["qL_period"] > 0).sum() == 0: parts.append("нет положительных объёмов жидкости")
            if (g["prod_days"] > 0).sum() == 0: parts.append("нет положительных дней добычи")
            parts.append(f"точек после fw>{watercut_thr:.2f}: {ok} (нужно ≥{min_points})")
            reasons[wid] = "; ".join(parts)
    return tuple(eligible), reasons


# --------------------------- 4) РАСЧЁТНЫЕ МОДУЛИ ----------------------------
# В этом блоке находятся ТОЛЬКО математические функции. Их можно заменять/обновлять
# без каких-либо правок в UI и I/O.

@dataclass
class MGFlags:
    y_early_mean: Optional[float] = None
    slope_first_third: Optional[float] = None
    waviness_std: Optional[float] = None
    possible_behind_casing: bool = False
    possible_channeling: bool = False
    possible_mixed_causes: bool = False

@st.cache_data
def compute_mg(df: pd.DataFrame, ids: Tuple[str, ...], thr: float, minp: int) -> pd.DataFrame:
    """Методика Меркуловой–Гинзбург (MG) — расчёт X, Y и диагностических флагов."""
    if df.empty or not ids:
        return pd.DataFrame(columns=["well_id","MG_X","MG_Y"])

    d = df[df["well_id"].isin(ids)].copy()
    with np.errstate(divide="ignore", invalid="ignore"):
        d["fw"] = d["qw_period"] / d["qL_period"]
    d = d.replace([np.inf, -np.inf], np.nan)

    outs: List[pd.DataFrame] = []
    for wid, g in d.groupby("well_id", sort=False):
        g = g.sort_values("t_num")
        idx = np.flatnonzero((g["fw"].to_numpy() > thr) &
                             (g["qL_period"].to_numpy() > 0) &
                             (g["prod_days"].to_numpy() > 0))
        if idx.size == 0:
            continue

        g2 = g.iloc[idx[0]:].copy()
        g2["Qo_cum"] = g2["qo_period"].cumsum()
        g2["Qw_cum"] = g2["qw_period"].cumsum()
        g2["Qt_cum"] = g2["Qo_cum"] + g2["Qw_cum"]
        if len(g2) < minp or float(g2["Qt_cum"].iloc[-1]) <= 0:
            continue

        Qt = float(g2["Qt_cum"].iloc[-1])
        X  = (g2["Qt_cum"] / Qt).to_numpy()
        X  = np.maximum.accumulate(X) + np.arange(len(X)) * EPS
        g2["MG_X"] = X
        with np.errstate(divide="ignore", invalid="ignore"):
            g2["MG_Y"] = g2["Qo_cum"] / g2["Qt_cum"]

        flags = MGFlags()
        early = g2["MG_X"] <= 0.2
        if early.sum() >= 3:
            flags.y_early_mean = float(np.nanmean(g2.loc[early, "MG_Y"]))
            flags.possible_behind_casing = (flags.y_early_mean is not None) and (flags.y_early_mean >= 0.99)

        first = g2[g2["MG_X"] <= 0.33]
        if len(first) >= 3:
            try:
                k, _ = np.polyfit(first["MG_X"], first["MG_Y"], 1)
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
        outs.append(g2)

    return pd.concat(outs, axis=0).reset_index(drop=True) if outs else pd.DataFrame(columns=["well_id","MG_X","MG_Y"])


@dataclass
class ChanFlags:
    slope_logWOR_logt: Optional[float] = None
    mean_derivative: Optional[float] = None
    std_derivative: Optional[float] = None
    possible_coning: bool = False
    possible_near_wellbore: bool = False
    possible_multilayer_channeling: bool = False

@st.cache_data
def compute_chan(df: pd.DataFrame, ids: Tuple[str, ...], minp: int) -> pd.DataFrame:
    """Методика Chan: WOR(t), производная и наклон log(WOR)~log(t)."""
    if df.empty or not ids:
        return pd.DataFrame(columns=["well_id","t_pos","WOR","dWOR_dt","dWOR_dt_pos"])

    d = df[df["well_id"].isin(ids)].copy()
    outs: List[pd.DataFrame] = []
    for wid, g in d.groupby("well_id", sort=False):
        g = g.sort_values("t_num").copy()
        with np.errstate(divide="ignore", invalid="ignore"):
            g["WOR"] = g.get("qw") / g.get("qo")
        g = g.replace([np.inf, -np.inf], np.nan)
        g = g[(g["qo"] > 0) & (g["WOR"] > 0)].dropna(subset=["WOR"])
        if len(g) < minp:
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
        flags.mean_derivative   = float(np.nanmean(g["dWOR_dt"]))
        flags.std_derivative    = float(np.nanstd(g["dWOR_dt"]))
        if not np.isnan(a):
            flags.possible_coning = a > 0.5 and flags.mean_derivative > 0
            flags.possible_near_wellbore = a > 1.0 and flags.mean_derivative > 0
            flags.possible_multilayer_channeling = a > 0 and flags.std_derivative > 0.1

        for k, v in vars(flags).items():
            g[f"chan_diag_{k}"] = v
        outs.append(g)

    return pd.concat(outs, axis=0).reset_index(drop=True) if outs else pd.DataFrame(columns=["well_id","t_pos","WOR","dWOR_dt","dWOR_dt_pos"])


# ----------------------- 5) ФОРМИРОВАНИЕ ТЕКСТОВ ВЫВОДОВ --------------------
def diag_mg(g: pd.DataFrame) -> Dict[str, str]:
    """Текстовое описание по MG для последней точки сущности."""
    if g.empty:
        return {"mg_text": "нет данных MG", "mg_detail": ""}
    r = g.iloc[-1]; parts: List[str] = []
    if r.get("MG_diag_possible_behind_casing"): parts.append("возможны заколонные перетоки (ранний нефтеотбор Y≈1)")
    if r.get("MG_diag_possible_channeling"):    parts.append("признаки каналирования (крутой спад Y в первой трети)")
    if r.get("MG_diag_possible_mixed_causes"):  parts.append("смешанные причины (высокая волнистость dY/dX)")
    if not parts: parts.append("ближе к равномерному обводнению")
    return {
        "mg_text": "; ".join(parts),
        "mg_detail": f"MG: y_early≈{_fmt(r.get('MG_diag_y_early_mean'))}; "
                     f"k≈{_fmt(r.get('MG_diag_slope_first_third'))}; "
                     f"std(dY/dX)≈{_fmt(r.get('MG_diag_waviness_std'))}"
    }

def diag_chan(g: pd.DataFrame) -> Dict[str, str]:
    """Текстовое описание по Chan для последней точки сущности."""
    if g.empty:
        return {"chan_text": "нет данных Chan", "chan_detail": ""}
    r = g.iloc[-1]; parts: List[str] = []
    if r.get("chan_diag_possible_multilayer_channeling"): parts.append("многослойное каналирование (рост WOR и дисперсии)")
    if r.get("chan_diag_possible_near_wellbore"):         parts.append("приствольные проблемы/ранний канал")
    if r.get("chan_diag_possible_coning"):                parts.append("возможен конинг (наклон > 0.5)")
    if not parts: parts.append("нет выраженных признаков проблемного притока воды")
    return {
        "chan_text": "; ".join(parts),
        "chan_detail": f"Chan: slope≈{_fmt(r.get('chan_diag_slope_logWOR_logt'))}; "
                       f"mean dWOR/dt≈{_fmt(r.get('chan_diag_mean_derivative'), '{:.2e}')}; "
                       f"std≈{_fmt(r.get('chan_diag_std_derivative'), '{:.2e}')}"
    }


# ---------------------------- 6) ЭКСПОРТ РЕЗУЛЬТАТОВ ------------------------
@st.cache_data
def export_xlsx(mg_df: pd.DataFrame, chan_df: pd.DataFrame, diagnosis_df: pd.DataFrame) -> bytes:
    """Единый Excel с листами Summary, MG и Chan (устойчив к пустым данным)."""
    if mg_df is None or mg_df.empty:     mg_df = pd.DataFrame(columns=["well_id"])
    if chan_df is None or chan_df.empty: chan_df = pd.DataFrame(columns=["well_id"])
    if diagnosis_df is None or diagnosis_df.empty:
        diagnosis_df = pd.DataFrame(columns=["well_id","mg_text","mg_detail","chan_text","chan_detail","reason"])
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as w:
        diagnosis_df.to_excel(w, sheet_name="Summary", index=False)
        mg_df.to_excel(w, sheet_name="MG", index=False)
        chan_df.to_excel(w, sheet_name="Chan", index=False)
    bio.seek(0)
    return bio.getvalue()


# ------------------------------- 7) ИНТЕРФЕЙС -------------------------------
# Вся визуальная часть изолирована. Меняя расчётные модули, мы не трогаем UI.

st.markdown(f"# {APP_TITLE}")
st.markdown("""
- Входной файл должен содержать **Скважина** и **Пласт** (если нет — будут созданы автоматически).
- Ключ анализа: `well_id = "Скважина | Пласт"`.
- Поддерживаемые входы:  
  ① *Жидкость (сут)* + *Обводнённость, %*; ② *Qo (сут)* + *Qw (сут)*;  
  ③ *Qo(мес)* + *Qw(мес)*; ④ *Qж(мес)* + *Обводнённость, %*.  
  *Дни добычи* (если нет — берём 1).
""")

# Шаблоны для скачивания
templates = make_templates()
c1, c2, c3 = st.columns(3)
c1.download_button("🔽 Шаблон (Жидкость+Обводнённость) — XLSX",
                   data=df_to_xlsx_bytes(templates["liq_wc"]),
                   file_name="template_liq_wc.xlsx")
c2.download_button("🔽 Шаблон (Жидкость+Обводнённость) — CSV",
                   data=df_to_csv_bytes(templates["liq_wc"]),
                   file_name="template_liq_wc.csv")
c3.download_button("🔽 Шаблон (Qo/Qw) — XLSX",
                   data=df_to_xlsx_bytes(templates["qo_qw"]),
                   file_name="template_qo_qw.xlsx")
st.markdown("---")

# Панель параметров
with st.sidebar:
    st.subheader("Параметры допуска")
    water_thr = st.number_input("Порог обводнённости fw", 0.0, 1.0, SHARED_WATERCUT_THR_DEFAULT, 0.01)
    min_pts   = st.number_input("Мин. число точек после порога", 3, 200, MIN_POINTS_DEFAULT, 1)
    max_plot  = st.slider("Сколько сущностей рисовать", 1, 50, 10)

# Загрузка и запуск
uploaded = st.file_uploader("Загрузите файл (.xlsx / .xls / .csv)", type=["xlsx","xls","csv"])
if not uploaded:
    st.info("Загрузите файл и нажмите «Запустить расчёт».")
else:
    file_bytes = _bytes_of_upload(uploaded)
    if st.button("▶ Запустить расчёт"):
        # 1) Подготовка
        with st.spinner("Подготовка данных..."):
            df, detector = prepare_data(file_bytes, uploaded.name)

        with st.expander("🔍 Диагностика распознанных колонок"):
            st.json(detector)

        # 2) Допуск сущностей
        eligible_ids, reasons = select_eligible_with_reasons(df, min_pts, water_thr)
        # если никто не прошёл — анализируем все well_id без ошибок
        if len(eligible_ids) == 0 and "well_id" in df.columns and not df.empty:
            eligible_ids = tuple(pd.Index(df["well_id"].unique()).astype(str))

        # 3) Расчёты (изолированные модули)
        prog = st.progress(0, text="Расчёт MG…")
        mg_df   = compute_mg(df, eligible_ids, water_thr, min_pts); prog.progress(50, text="Расчёт Chan…")
        chan_df = compute_chan(df, eligible_ids, min_pts);          prog.progress(100, text="Готово")

        # 4) Сводка и диагнозы
        rows: List[Dict[str, str]] = []
        for wid in eligible_ids:
            mg_g = _safe_slice_by_well_id(mg_df, wid)
            ch_g = _safe_slice_by_well_id(chan_df, wid)
            row  = {"well_id": wid, **diag_mg(mg_g), **diag_chan(ch_g)}
            if wid in reasons: row["reason"] = reasons[wid]
            rows.append(row)
        summary = pd.DataFrame(rows)

        st.success(f"Готово. Сущностей (Скважина|Пласт): {len(eligible_ids)}")
        st.subheader("Сводная таблица диагнозов")
        st.dataframe(summary, use_container_width=True)

        # 5) Графики (ограниченно для скорости)
        st.subheader("Графики (ограниченный список)")
        for wid in eligible_ids[:max_plot]:
            with st.expander(f"{wid} — графики (MG и Chan)"):
                left, right = st.columns(2)
                mg_g = _safe_slice_by_well_id(mg_df, wid)
                ch_g = _safe_slice_by_well_id(chan_df, wid)

                # MG
                with left:
                    if not mg_g.empty and {"MG_X","MG_Y"}.issubset(mg_g.columns) and mg_g[["MG_X","MG_Y"]].dropna().shape[0] > 0:
                        fig, ax = plt.subplots()
                        ax.scatter(mg_g["MG_X"], mg_g["MG_Y"], s=10)
                        ax.grid(True, alpha=0.3)
                        ax.set_xlabel("X = Qt_cum/Qt_cum(T)")
                        ax.set_ylabel("Y = Qo_cum/Qt_cum")
                        ax.set_title(f"MG — {wid}")
                        st.pyplot(fig)
                    else:
                        st.info("Нет пригодных точек для MG-графика")

                # Chan
                with right:
                    needed = {"t_pos","WOR","dWOR_dt_pos"}
                    if not ch_g.empty and needed.issubset(ch_g.columns) and ch_g[list(needed)].dropna(how="all").shape[0] > 0:
                        fig2, ax2 = plt.subplots()
                        ax2.plot(ch_g["t_pos"], ch_g["WOR"], "o", markersize=3, label="WOR")
                        ax2.plot(ch_g["t_pos"], ch_g["dWOR_dt_pos"], "--", label="|dWOR/dt|")
                        ax2.set_xscale("log"); ax2.set_yscale("log")
                        ax2.grid(True, which="both", alpha=0.3); ax2.legend()
                        ax2.set_xlabel("t_pos (дни)"); ax2.set_ylabel("WOR, |dWOR/dt|")
                        ax2.set_title(f"Chan — {wid}")
                        st.pyplot(fig2)
                    else:
                        st.info("Нет пригодных точек для Chan-графика")

        # 6) Экспорт
        st.subheader("Скачать результаты")
        st.download_button("📥 Единый Excel (Summary, MG, Chan)",
                           data=export_xlsx(mg_df, chan_df, summary),
                           file_name="Autodiagnostics_results_by_layer.xlsx")
    else:
        st.info("Интерфейс готов. Нажмите «Запустить расчёт».")
