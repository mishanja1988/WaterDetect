# -*- coding: utf-8 -*-
# Streamlit app: Автодиагностика скважин (Chan & Меркулова–Гинзбург) по (Скважина|Пласт)

from __future__ import annotations

import io
import os
import re
import unicodedata
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st

# ========================
# Глобальные настройки
# ========================
EPS = 1e-9
MIN_POINTS_DEFAULT = 6                 # сделали мягче, чтобы чаще было "допущено"
SHARED_WATERCUT_THR_DEFAULT = 0.02     # fw-порог для MG
APP_TITLE = "Поскважинный автодиагноз (Chan & Меркулова–Гинзбург) — учёт Скважины и Пласта"

st.set_page_config(
    layout="wide",
    initial_sidebar_state="auto",
    page_title="Автодиагностика скважин",
    page_icon="🛢️",
)

# ========================
# Описание и кнопки-шаблоны
# ========================
DESCRIPTION_MD = f"""
# {APP_TITLE}

- Входной файл должен содержать **Скважина** и **Пласт** (если чего-то нет — приложение создаст столбец автоматически).
- Виртуальный ключ анализа: **`well_id = "Скважина | Пласт"`**. Все группировки, допуски, графики и экспорт ведутся **строго по нему**.
- Поддерживаются варианты исходных данных:
  1) *Жидкость, м3/сут* **+** *Обводнённость, %*  → алгоритм сам посчитает нефть/воду;
  2) *Дебит нефти, м3/сут* **+** *Дебит воды, м3/сут*  → будет вычислено всё остальное;
  3) *Добыча нефти м3/мес* **+** *Добыча воды м3/мес*  → используются как периодные объёмы.

**Совет.** Если у вас суточные дебиты — добавьте колонку *Дни добычи* (для каждого периода). Если её нет — по умолчанию считаем 1 день.
"""

st.markdown(DESCRIPTION_MD)

# ========================
# Утилиты
# ========================
def _norm(s):
    if not isinstance(s, str):
        return str(s)
    s = unicodedata.normalize("NFKC", s).replace("\u00A0", " ").replace("\xa0", " ")
    return re.sub(r"\s+", " ", s.strip())

def _to_num(x, fill=None):
    s = pd.to_numeric(x, errors="coerce")
    return s.fillna(fill) if fill is not None else s

def _drop_unnamed(df: pd.DataFrame) -> pd.DataFrame:
    return df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]

def _bytes_of_upload(uploaded_file) -> bytes:
    uploaded_file.seek(0)
    b = uploaded_file.read()
    uploaded_file.seek(0)
    return b

def _slice_by_well_id(df: pd.DataFrame, wid: str) -> pd.DataFrame:
    if df is None or df.empty or "well_id" not in df.columns:
        return pd.DataFrame()
    return df[df["well_id"] == wid]

# ========================
# Шаблоны и примеры
# ========================
def _template_df() -> pd.DataFrame:
    # Мини-шаблон с нужными колонками
    return pd.DataFrame({
        "Скважина": ["113", "113", "115", "115"],
        "Пласт": ["Б6", "Б6", "Бш", "Бш"],
        "Дата": pd.to_datetime(["2024-01-01","2024-02-01","2024-01-01","2024-02-01"]),
        "Дни добычи": [31, 29, 31, 29],
        "Жидкость, м3/сут": [100, 120, 95, 110],
        "Обводнённость, %": [10, 22, 5, 18],
    })

def _template_rate_df() -> pd.DataFrame:
    # Альтернативный шаблон с суточными дебитами нефти и воды
    return pd.DataFrame({
        "Скважина": ["A-01", "A-01", "A-02", "A-02"],
        "Пласт": ["Ю1", "Ю1", "Ю2", "Ю2"],
        "Дата": pd.to_datetime(["2024-03-01","2024-04-01","2024-03-01","2024-04-01"]),
        "Дни добычи": [31, 30, 31, 30],
        "Дебит нефти, м3/сут": [80, 70, 60, 50],
        "Дебит воды, м3/сут": [20, 40, 10, 25],
    })

def _bytes_xlsx(df: pd.DataFrame) -> bytes:
    bio = io.BytesIO()
    df.to_excel(bio, index=False, engine="openpyxl")
    bio.seek(0)
    return bio.getvalue()

def _bytes_csv(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")

st.markdown("### 📥 Шаблоны для скачивания")
c1, c2, c3 = st.columns(3)
c1.download_button("🔽 Шаблон (Жидкость+Обводнённость) — XLSX", data=_bytes_xlsx(_template_df()), file_name="template_liq_wc.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
c2.download_button("🔽 Шаблон (Жидкость+Обводнённость) — CSV",  data=_bytes_csv(_template_df()),    file_name="template_liq_wc.csv",  mime="text/csv")
c3.download_button("🔽 Шаблон (Дебиты нефти/воды) — XLSX",     data=_bytes_xlsx(_template_rate_df()), file_name="template_qo_qw.xlsx",  mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.markdown("---")

# ========================
# Чтение пользовательского файла
# ========================
@st.cache_data
def read_user_file(file_bytes: bytes, filename: str) -> pd.DataFrame:
    name = (filename or "").lower()
    # Excel
    if name.endswith(".xlsx") or name.endswith(".xls"):
        df = pd.read_excel(io.BytesIO(file_bytes))
        return _drop_unnamed(df)

    # CSV
    last_exc = None
    for enc in ("utf-8", "utf-8-sig", "cp1251", "latin1"):
        try:
            df = pd.read_csv(io.BytesIO(file_bytes), sep=None, engine="python", encoding=enc, on_bad_lines="skip")
            return _drop_unnamed(df)
        except Exception as e:
            last_exc = e
    raise ValueError(f"Не удалось прочитать файл как CSV/Excel: {last_exc}")

# ========================
# Подготовка: Скважина/Пласт и well_id
# ========================
def _ensure_well_and_layer(df: pd.DataFrame) -> pd.DataFrame:
    cols_map = {c.lower(): c for c in df.columns}

    # --- Скважина ---
    well_col = None
    for cand in ["скважина", "скв", "скв.", "well", "id", "номер"]:
        if cand in cols_map:
            well_col = cols_map[cand]; break
    if well_col is None:
        df = df.copy()
        df["Скважина"] = [f"WELL_{i+1}" for i in range(len(df))]
        well_col = "Скважина"

    # --- Пласт ---
    layer_col = None
    for cand in ["пласт", "пл", "layer", "horizon", "formation"]:
        if cand in cols_map:
            layer_col = cols_map[cand]; break
    if layer_col is None:
        df = df.copy()
        df["Пласт"] = "UNK"
        layer_col = "Пласт"

    out = df.copy()
    out["Скважина"] = out[well_col].astype(str).fillna("").str.strip()
    out["Пласт"]    = out[layer_col].astype(str).fillna("").replace({"": "UNK"}).str.strip()
    out["well_id"]  = (out["Скважина"] + " | " + out["Пласт"]).str.strip()
    return out

def enforce_monotonic_per_entity(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby("well_id", sort=False)["t_num"]
    t = g.cummax()
    idx_in_grp = df.groupby("well_id", sort=False).cumcount().to_numpy()
    out = df.copy()
    out["t_num"] = t.to_numpy() + idx_in_grp * EPS
    return out

@st.cache_data
def prepare_data(file_bytes: bytes, filename: str) -> pd.DataFrame:
    raw = read_user_file(file_bytes, filename)
    raw.columns = [_norm(c) for c in raw.columns]
    df = _ensure_well_and_layer(raw)

    # Популярные названия
    cols = {c.lower(): c for c in df.columns}
    def pick(*names):
        for n in names:
            if n in cols: return cols[n]
        return None

    # Варианты исходных показателей
    c_liq = pick("жидкость, м3/сут", "дебит жидкости, м3/сут", "liquid")
    c_wc  = pick("обводнённость, %", "обводненность, %", "watercut %")
    c_qo  = pick("дебит нефти, м3/сут", "qo", "нефть, м3/сут")
    c_qw  = pick("дебит воды, м3/сут", "qw", "вода, м3/сут")

    c_qoP = pick("добыча нефти м3/мес", "qo_period")
    c_qwP = pick("добыча воды м3/мес",  "qw_period")
    c_qL  = pick("дебит жидкости, м3/сут", "жидкость, м3/сут", "qL")
    c_days = pick("дни добычи", "число дней добычи нефти, сут", "prod_days", "aj")
    c_date = pick("дата", "месяц", "period", "дата (месяц, год)")
    c_tcum = pick("накопленное время работы", "накопленно времени")

    # 1) Если есть периодные объёмы нефти/воды — используем их напрямую
    if c_qoP and c_qwP:
        df["qo_period"] = _to_num(df[c_qoP], 0.0)
        df["qw_period"] = _to_num(df[c_qwP], 0.0)
        df["qL_period"] = df["qo_period"] + df["qw_period"]

    # 2) Иначе если есть (Жидкость, % обводненности)
    elif c_liq and c_wc:
        liq = _to_num(df[c_liq], 0.0)
        wc  = _to_num(df[c_wc],  0.0)
        # считаем суточные дебиты
        df["qo"] = liq * (100.0 - wc) / 100.0
        df["qw"] = liq * wc / 100.0
        df["qL"] = liq
        days = _to_num(df[c_days], 1.0) if c_days else pd.Series(1.0, index=df.index)
        df["prod_days"] = days
        df["qo_period"] = df["qo"] * days
        df["qw_period"] = df["qw"] * days
        df["qL_period"] = df["qL"] * days

    # 3) Иначе если есть суточные дебиты нефти/воды
    elif c_qo and c_qw:
        df["qo"] = _to_num(df[c_qo], 0.0)
        df["qw"] = _to_num(df[c_qw], 0.0)
        df["qL"] = df["qo"] + df["qw"]
        days = _to_num(df[c_days], 1.0) if c_days else pd.Series(1.0, index=df.index)
        df["prod_days"] = days
        df["qo_period"] = df["qo"] * days
        df["qw_period"] = df["qw"] * days
        df["qL_period"] = df["qL"] * days

    else:
        # минимально жизнеспособный вариант: всё по нулям, чтобы не падать
        df["prod_days"] = _to_num(df[c_days], 1.0) if c_days else pd.Series(1.0, index=df.index)
        df["qo_period"] = pd.Series(0.0, index=df.index)
        df["qw_period"] = pd.Series(0.0, index=df.index)
        df["qL_period"] = pd.Series(0.0, index=df.index)
        df["qo"] = pd.Series(np.nan, index=df.index)
        df["qw"] = pd.Series(np.nan, index=df.index)
        df["qL"] = pd.Series(np.nan, index=df.index)

    # время
    if c_tcum:
        df["t_num"] = _to_num(df[c_tcum], 0.0)
    elif c_date:
        t = pd.to_datetime(df[c_date], errors="coerce")
        df["t_num"] = (t - t.groupby(df["well_id"]).transform("min")).dt.days.astype(float)
        df["t_num"] = df["t_num"].fillna(0.0)
    else:
        # если нет ни даты, ни накопленного времени — используем кумулятив по дням
        df["t_num"] = df["prod_days"].groupby(df["well_id"]).cumsum()

    df = df.dropna(subset=["well_id", "t_num"]).sort_values(["well_id", "t_num"]).reset_index(drop=True)
    df = enfor
