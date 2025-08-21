# app.py — Автодиагностика скважин (единый файл)
# ------------------------------------------------
# Запуск: streamlit run app.py
# Требуется: streamlit, pandas, numpy, matplotlib, openpyxl, xlsxwriter, Pillow

from __future__ import annotations

import io
import os
import re
import sys
import unicodedata
from dataclasses import dataclass
from io import BytesIO
from typing import Dict, List, Optional

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st

# =========================
# Глобальные настройки
# =========================
EPS = 1e-9
# Путь к шаблону из вложения (как просили)
TEMPLATE_PATH = "data/templates/Сосновское_clean.xlsx"

st.set_page_config(
    layout="wide",
    initial_sidebar_state="auto",
    page_title="Автодиагностика скважин",
    page_icon="🛢️",
)

st.write("### Поскважинный автодиагноз нефтяных скважин по механизму обводнения")
st.markdown(
    """
**Суть работы:** проведение расчетно-аналитического способов механизма обводнения скважин с использованием методики Чена (Chan) и Меркуловой–Гинзбурга (MG) по нефтяным скважинам на основе пользовательских исходных данных.

**Что необходимо сделать:**  
1. Скачать шаблон исходных данных;  
2. Заполнить шаблон своими данными;  
3. Подгрузить Ваш шаблон в окно подгрузки данных;  
4. Получить результат — текстовый и визуальный автодиагноз по каждой скважине;  
5. Скачать единый файл с результатами (Summary + MG + Chan с графиками).
"""
)

# =========================
# Утилиты
# =========================
def excel_letter_to_index(letter: str) -> int:
    """A->0, B->1, ..., Z->25, AA->26, AB->27, ..."""
    letter = letter.strip().upper()
    acc = 0
    for ch in letter:
        if not ("A" <= ch <= "Z"):
            raise ValueError(f"Неверная буква столбца Excel: {letter}")
        acc = acc * 26 + (ord(ch) - ord("A") + 1)
    return acc - 1


def col_by_letter(df: pd.DataFrame, letter: str) -> Optional[str]:
    """Вернуть имя столбца по букве Excel с учётом текущего порядка колонок."""
    idx = excel_letter_to_index(letter)
    return df.columns[idx] if 0 <= idx < len(df.columns) else None


def normalize_header(s: str) -> str:
    if not isinstance(s, str):
        return s
    s = unicodedata.normalize("NFKC", s).replace("\u00A0", " ").replace("\xa0", " ")
    s = re.sub(r"\s+", " ", s.strip())
    return s


def save_df_to_excel(df: pd.DataFrame, ind: bool = False) -> BytesIO:
    output = BytesIO()
    df.to_excel(output, index=ind, engine="openpyxl")
    output.seek(0)
    return output


@st.cache_data
def read_template_as_examples() -> tuple[pd.DataFrame, pd.DataFrame]:
    """Читаем шаблон для скачивания пользователем (в CSV/XLSX)."""
    if os.path.exists(TEMPLATE_PATH):
        df_template = pd.read_excel(TEMPLATE_PATH)
    else:
        st.warning(f"Шаблон не найден по пути {TEMPLATE_PATH}. Использую пустой DataFrame.")
        df_template = pd.DataFrame()
    # Возвращаем 2 копии — для CSV и XLSX кнопок
    return df_template, df_template


def upload_examples() -> None:
    """Кнопки скачивания шаблона."""
    example_csv, example_excel = read_template_as_examples()
    st.write("**Скачать шаблон таблицы для расчётов:**")
    col1, col2, *_ = st.columns(9)
    btn_csv = col1.download_button(
        label="Скачать шаблон в .csv",
        data=example_csv.to_csv(index=False),
        file_name="template_from_attachment.csv",
        mime="text/csv",
    )
    btn_xlsx = col2.download_button(
        label="Скачать шаблон в .xlsx",
        data=save_df_to_excel(example_excel),
        file_name="template_from_attachment.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    if btn_csv or btn_xlsx:
        st.success("Шаблон успешно сохранён.")


# =========================
# Подготовка данных под MG/Chan
# =========================
def enforce_monotonic_per_well(dfin: pd.DataFrame) -> pd.DataFrame:
    out = []
    for w, g in dfin.groupby("well", sort=False):
        t = g["t_num"].to_numpy(dtype=float)
        for i in range(1, t.size):
            if t[i] <= t[i - 1]:
                t[i] = t[i - 1] + EPS
        g = g.copy()
        g["t_num"] = t
        out.append(g)
    return pd.concat(out, axis=0).reset_index(drop=True)


def compute_cum_work_time(group: pd.DataFrame, col_BR: str, col_AJ: str) -> pd.Series:
    """
    Накопленное время работы в группе Well_calc:
    если BR[i] == BR[i-1] => AJ[i] + cum[i-1], иначе AJ[i]
    Считаем по порядку строк (как в файле).
    """
    br = group[col_BR].astype(str).fillna("")
    aj = pd.to_numeric(group[col_AJ], errors="coerce").fillna(0.0).to_numpy()
    out = np.zeros(len(group), dtype=float)
    for i in range(len(group)):
        if i == 0:
            out[i] = aj[i]
        else:
            out[i] = aj[i] + out[i - 1] if br.iloc[i] == br.iloc[i - 1] else aj[i]
    return pd.Series(out, index=group.index, name="Накопленное время работы")


def data_preparation(init_data: pd.DataFrame) -> pd.DataFrame:
    """
    Адаптация входного Excel к расчётам MG/Chan согласно новым требованиям.
    Используются фиксированные буквы колонок:
      H=Скважина, I=Объект, X=Жидкость (м3/мес), AB=Обводнённость (%),
      BT и BS для ВНФ = BT/BS, BR=период (ключ), AJ=дни работы.
    """
    dfn = init_data.copy()
    dfn.columns = [normalize_header(c) for c in dfn.columns]

    # Карта букв
    cH = col_by_letter(dfn, "H")   # Скважина
    cI = col_by_letter(dfn, "I")   # Объект
    cX = col_by_letter(dfn, "X")   # Жидкость, м3/мес
    cAB = col_by_letter(dfn, "AB") # Обводнённость, %
    cBT = col_by_letter(dfn, "BT") # Числитель для ВНФ
    cBS = col_by_letter(dfn, "BS") # Знаменатель для ВНФ
    cBR = col_by_letter(dfn, "BR") # Ключ периода (месяц/год)
    cAJ = col_by_letter(dfn, "AJ") # Дни работы в периоде

    # Well_calc = H + " " + I
    dfn["Well_calc"] = (
        dfn[cH].astype(str).str.strip().fillna("") + " " + dfn[cI].astype(str).str.strip().fillna("")
        if cH and cI
        else dfn.get(cH or cI, pd.Series([""] * len(dfn)))
    )
    dfn["well"] = dfn["Well_calc"]

    # Производные столбцы
    # Добыча нефти м3/мес = X * (100 - AB) / 100
    if cX and cAB:
        X_vals = pd.to_numeric(dfn[cX], errors="coerce")
        AB_vals = pd.to_numeric(dfn[cAB], errors="coerce")
        dfn["Добыча нефти м3/мес"] = X_vals * (100.0 - AB_vals) / 100.0
        dfn["Добыча воды м3/мес"] = X_vals * AB_vals / 100.0
    else:
        dfn["Добыча нефти м3/мес"] = np.nan
        dfn["Добыча воды м3/мес"] = np.nan

    # ВНФ = BT/BS
    if cBT and cBS:
        dfn["ВНФ"] = pd.to_numeric(dfn[cBT], errors="coerce") / pd.to_numeric(dfn[cBS], errors="coerce")
    else:
        dfn["ВНФ"] = np.nan

    # Накопленное время работы (по Well_calc)
    if cBR and cAJ:
        dfn["Накопленное время работы"] = 0.0
        for w, g in dfn.groupby("Well_calc", sort=False):
            dfn.loc[g.index, "Накопленное время работы"] = compute_cum_work_time(g, cBR, cAJ)
    else:
        dfn["Накопленное время работы"] = np.arange(len(dfn), dtype=float)

    # ВНФ' — производная по «Накопленное время работы»
    try:
        t = pd.to_numeric(dfn["Накопленное время работы"], errors="coerce").to_numpy()
        y = pd.to_numeric(dfn["ВНФ"], errors="coerce").to_numpy()
        grad = np.full(len(dfn), np.nan)
        for w, g in dfn.groupby("Well_calc", sort=False):
            idx = g.index.to_numpy()
            tt = t[idx]
            yy = y[idx]
            with np.errstate(invalid="ignore"):
                grad[idx] = np.gradient(yy, tt)
        dfn["ВНФ'"] = grad
    except Exception:
        dfn["ВНФ'"] = np.nan

    # Периодные объёмы и суточные дебиты
    dfn["qo_period"] = pd.to_numeric(dfn["Добыча нефти м3/мес"], errors="coerce").fillna(0.0)
    dfn["qw_period"] = pd.to_numeric(dfn["Добыча воды м3/мес"], errors="coerce").fillna(0.0)
    dfn["qL_period"] = dfn["qo_period"] + dfn["qw_period"]

    if cAJ:
        dfn["prod_days"] = pd.to_numeric(dfn[cAJ], errors="coerce").fillna(0.0)
    else:
        dfn["prod_days"] = np.nan

    dfn["qo"] = np.where(dfn["prod_days"] > 0, dfn["qo_period"] / dfn["prod_days"], np.nan)
    dfn["qw"] = np.where(dfn["prod_days"] > 0, dfn["qw_period"] / dfn["prod_days"], np.nan)
    dfn["qL"] = np.where(dfn["prod_days"] > 0, dfn["qL_period"] / dfn["prod_days"], np.nan)

    # Время расчётов
    dfn["t_num"] = pd.to_numeric(dfn["Накопленное время работы"], errors="coerce").fillna(0.0)

    # Сортировка и монотонность времени
    dfn = dfn.dropna(subset=["well", "t_num"]).sort_values(["well", "t_num"]).reset_index(drop=True)
    dfn = enforce_monotonic_per_well(dfn)
    return dfn


# =========================
# MG
# =========================
@dataclass
class MGFlags:
    y_early_mean: Optional[float] = None
    slope_first_third: Optional[float] = None
    waviness_std: Optional[float] = None
    possible_behind_casing: bool = False
    possible_channeling: bool = False
    possible_mixed_causes: bool = False


def compute_mg_full(df_in: pd.DataFrame, watercut_thr: float = 0.02, min_points: int = 8) -> pd.DataFrame:
    d = df_in.copy()
    with np.errstate(divide="ignore", invalid="ignore"):
        d["fw"] = np.where(d["qL_period"] > 0, d["qw_period"] / d["qL_period"], np.nan)

    frames = []
    for w, g in d.groupby("well", sort=False):
        g = g.sort_values("t_num").copy()
        idx = g.index[g["fw"] > watercut_thr]
        if len(idx) == 0 or len(g) < min_points:
            continue
        g2 = g.loc[idx[0]:].copy()

        g2["Qo_cum"] = g2["qo_period"].cumsum()
        g2["Qw_cum"] = g2["qw_period"].cumsum()
        g2["Qt_cum"] = g2["Qo_cum"] + g2["Qw_cum"]

        Qt_T = float(g2["Qt_cum"].iloc[-1])
        if Qt_T <= 0 or len(g2) < min_points:
            continue

        X = (g2["Qt_cum"] / Qt_T).to_numpy()
        for i in range(1, X.size):
            if X[i] <= X[i - 1]:
                X[i] = X[i - 1] + EPS
        g2["MG_X"] = X
        with np.errstate(invalid="ignore", divide="ignore"):
            g2["MG_Y"] = np.where(g2["Qt_cum"] > 0, g2["Qo_cum"] / g2["Qt_cum"], np.nan)

        flags = MGFlags()

        # Ранние точки
        early_mask = g2["MG_X"] <= 0.2
        if early_mask.sum() >= 3:
            flags.y_early_mean = float(np.nanmean(g2.loc[early_mask, "MG_Y"]))
            flags.possible_behind_casing = (flags.y_early_mean is not None) and (flags.y_early_mean >= 0.99)

        # Наклон в первой трети
        first_third = g2[g2["MG_X"] <= 0.33]
        if len(first_third) >= 3:
            x = first_third["MG_X"].to_numpy()
            y = first_third["MG_Y"].to_numpy()
            A = np.vstack([x, np.ones_like(x)]).T
            try:
                k, b = np.linalg.lstsq(A, y, rcond=None)[0]
                flags.slope_first_third = float(k)
                flags.possible_channeling = (k < -0.8)
            except Exception:
                pass

        # Волнистость
        if len(g2) >= 5:
            with np.errstate(invalid="ignore"):
                dy = np.gradient(g2["MG_Y"].to_numpy(), g2["MG_X"].to_numpy())
            flags.waviness_std = float(np.nanstd(dy))
            flags.possible_mixed_causes = flags.waviness_std > 1.0

        for key, val in {
            "MG_diag_y_early_mean": flags.y_early_mean,
            "MG_diag_slope_first_third": flags.slope_first_third,
            "MG_diag_waviness_std": flags.waviness_std,
            "MG_flag_behind_casing": flags.possible_behind_casing,
            "MG_flag_channeling": flags.possible_channeling,
            "MG_flag_mixed": flags.possible_mixed_causes,
        }.items():
            g2[key] = val

        frames.append(g2.assign(well=w))

    return pd.concat(frames, axis=0).reset_index(drop=True) if frames else pd.DataFrame()


# =========================
# Chan
# =========================
@dataclass
class ChanFlags:
    slope_logWOR_logt: Optional[float] = None
    mean_derivative: Optional[float] = None
    std_derivative: Optional[float] = None
    possible_coning: bool = False
    possible_near_wellbore: bool = False
    possible_multilayer_channeling: bool = False


def compute_chan_full(df_in: pd.DataFrame, min_points: int = 10) -> pd.DataFrame:
    frames = []
    for w, g in df_in.groupby("well", sort=False):
        g = g.sort_values("t_num").copy()
        with np.errstate(divide="ignore", invalid="ignore"):
            g["WOR"] = g["qw"] / g["qo"]
        g = g.replace([np.inf, -np.inf], np.nan)
        g = g[(g["qo"] > 0) & (g["WOR"] > 0)].dropna(subset=["WOR"])
        if len(g) < min_points:
            continue

        with np.errstate(invalid="ignore"):
            g["t_pos"] = g["t_num"] - g["t_num"].min() + EPS
            g["dWOR_dt"] = np.gradient(g["WOR"].to_numpy(), g["t_pos"].to_numpy())

        # Наклон в log-log
        mask = (g["WOR"] > 0) & (g["t_pos"] > 0)
        x = np.log(g.loc[mask, "t_pos"].to_numpy())
        y = np.log(g.loc[mask, "WOR"].to_numpy())
        if len(x) >= 3:
            A = np.vstack([x, np.ones_like(x)]).T
            try:
                a, b = np.linalg.lstsq(A, y, rcond=None)[0]
            except Exception:
                a = np.nan
        else:
            a = np.nan

        mean_deriv = float(np.nanmean(g["dWOR_dt"])) if len(g) else np.nan
        std_deriv = float(np.nanstd(g["dWOR_dt"])) if len(g) else np.nan

        g["well"] = w
        g["chan_slope_logWOR_logt"] = float(a) if a == a else np.nan
        g["chan_mean_dWOR_dt"] = mean_deriv
        g["chan_std_dWOR_dt"] = std_deriv
        g["chan_flag_coning"] = (a > 0.5 and mean_deriv > 0) if a == a else False
        g["chan_flag_near_wellbore"] = (a > 1.0 and mean_deriv > 0) if a == a else False
        g["chan_flag_multilayer_channeling"] = (a > 0 and std_deriv > 0.1) if a == a else False

        # Для графика log–log: положительная производная
        g["dWOR_dt_pos"] = np.where(g["dWOR_dt"] > 0, g["dWOR_dt"], np.nan)

        frames.append(g)

    return pd.concat(frames, axis=0).reset_index(drop=True) if frames else pd.DataFrame()


# =========================
# Текстовые диагнозы
# =========================
def diagnose_mg_group(g: pd.DataFrame) -> Dict[str, str]:
    y_early = g["MG_diag_y_early_mean"].dropna().iloc[-1] if "MG_diag_y_early_mean" in g and g["MG_diag_y_early_mean"].notna().any() else np.nan
    slope = g["MG_diag_slope_first_third"].dropna().iloc[-1] if "MG_diag_slope_first_third" in g and g["MG_diag_slope_first_third"].notna().any() else np.nan
    wav = g["MG_diag_waviness_std"].dropna().iloc[-1] if "MG_diag_waviness_std" in g and g["MG_diag_waviness_std"].notna().any() else np.nan
    f_bc = bool(g["MG_flag_behind_casing"].dropna().iloc[-1]) if "MG_flag_behind_casing" in g and g["MG_flag_behind_casing"].notna().any() else False
    f_ch = bool(g["MG_flag_channeling"].dropna().iloc[-1]) if "MG_flag_channeling" in g and g["MG_flag_channeling"].notna().any() else False
    f_mix = bool(g["MG_flag_mixed"].dropna().iloc[-1]) if "MG_flag_mixed" in g and g["MG_flag_mixed"].notna().any() else False

    parts: List[str] = []
    if f_bc:
        parts.append("возможны заколонные перетоки (ранний нефтеотбор Y≈1)")
    if f_ch:
        parts.append("признаки каналирования (крутой спад Y в первой трети)")
    if f_mix:
        parts.append("смешанные причины (высокая волнистость dY/dX)")
    if not parts:
        parts.append("характеристика ближе к равномерному обводнению")

    detail = f"MG метрики: y_early≈{y_early:.2f}; наклон≈{slope:.2f}; волнистость≈{wav:.2f}"
    return {"mg_text": "; ".join(parts), "mg_detail": detail}


def diagnose_chan_group(g: pd.DataFrame) -> Dict[str, str]:
    slope = g["chan_slope_logWOR_logt"].dropna().iloc[-1] if "chan_slope_logWOR_logt" in g and g["chan_slope_logWOR_logt"].notna().any() else np.nan
    mean_d = g["chan_mean_dWOR_dt"].dropna().iloc[-1] if "chan_mean_dWOR_dt" in g and g["chan_mean_dWOR_dt"].notna().any() else np.nan
    std_d = g["chan_std_dWOR_dt"].dropna().iloc[-1] if "chan_std_dWOR_dt" in g and g["chan_std_dWOR_dt"].notna().any() else np.nan
    f_cone = bool(g["chan_flag_coning"].dropna().iloc[-1]) if "chan_flag_coning" in g and g["chan_flag_coning"].notna().any() else False
    f_near = bool(g["chan_flag_near_wellbore"].dropna().iloc[-1]) if "chan_flag_near_wellbore" in g and g["chan_flag_near_wellbore"].notna().any() else False
    f_multi = bool(g["chan_flag_multilayer_channeling"].dropna().iloc[-1]) if "chan_flag_multilayer_channeling" in g and g["chan_flag_multilayer_channeling"].notna().any() else False

    parts: List[str] = []
    if f_multi:
        parts.append("многослойное каналирование (рост WOR и дисперсии производной)")
    if f_near:
        parts.append("приствольные проблемы/ранний канал (очень высокий наклон)")
    if f_cone:
        parts.append("возможен конинг (наклон > 0.5 при положительной производной)")
    if not parts:
        parts.append("нет выраженных признаков проблемного притока воды")

    detail = f"Chan метрики: наклон≈{slope:.2f}; средн. dWOR/dt≈{mean_d:.2e}; std≈{std_d:.2e}"
    return {"chan_text": "; ".join(parts), "chan_detail": detail}


# =========================
# Экспорт «всё в один Excel»
# =========================
def _render_plot_image(kind: str, g: pd.DataFrame, well: str) -> BytesIO:
    """Сгенерировать matplotlib-график в PNG (в памяти)."""
    buf = BytesIO()
    if kind == "MG":
        fig, ax = plt.subplots(figsize=(6, 4))
        ax.scatter(g["MG_X"], g["MG_Y"], s=16, label="MG: Y(X)")
        ax.set_title(f"MG — скважина {well}")
        ax.set_xlabel("X = Qt_cum / Qt_cum(T)")
        ax.set_ylabel("Y = Qo_cum / Qt_cum")
        ax.grid(True, alpha=0.3)
        ax.legend(loc="best")
    else:  # Chan
        fig, ax = plt.subplots(figsize=(6, 4))
        m_wor = (g["t_pos"] > 0) & (g["WOR"] > 0)
        m_der = (g["t_pos"] > 0) & (g["dWOR_dt_pos"] > 0)
        ax.plot(
            g.loc[m_wor, "t_pos"],
            g.loc[m_wor, "WOR"],
            marker="o",
            linestyle="none",
            markersize=4,
            label="WOR",
        )
        ax.plot(
            g.loc[m_der, "t_pos"],
            g.loc[m_der, "dWOR_dt_pos"],
            linestyle="--",
            label="|dWOR/dt|",
        )
        ax.set_xscale("log")
        ax.set_yscale("log")
        ax.set_xlabel("t_pos (дни)")
        ax.set_ylabel("WOR, |dWOR/dt|")
        ax.grid(True, which="both", alpha=0.3)
        ax.legend(loc="best")
        ax.set_title(f"Chan — скважина {well} (log–log)")
    fig.tight_layout()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


def export_all_results_single_file(
    mg_df: pd.DataFrame, chan_df: pd.DataFrame, diagnosis_df: pd.DataFrame
) -> BytesIO:
    """
    Итоговый XLSX:
      - Summary: diagnosis_df
      - MG: блоки по скважинам (таблица + график справа)
      - Chan: блоки по скважинам (таблица + график справа)
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # --- Summary ---
        if diagnosis_df is not None and not diagnosis_df.empty:
            diagnosis_df.to_excel(writer, sheet_name="Summary", index=False)
            writer.sheets["Summary"].set_column(0, diagnosis_df.shape[1] - 1, 18)
        else:
            ws_sum = writer.book.add_worksheet("Summary")
            writer.sheets["Summary"] = ws_sum
            ws_sum.write(0, 0, "Нет сводных диагнозов")

        # --- MG ---
        ws_mg = writer.book.add_worksheet("MG")
        writer.sheets["MG"] = ws_mg
        cur_row = 0
        if mg_df is not None and not mg_df.empty:
            for well, g in mg_df.groupby("well", sort=False):
                title = f"Скважина {well} — MG"
                g_reset = g.reset_index(drop=True)
                g_reset.to_excel(writer, sheet_name="MG", index=False, startrow=cur_row + 1, startcol=0)
                ws_mg.write(cur_row, 0, title)
                ws_mg.set_column(0, min(8, g_reset.shape[1] - 1), 14)
                if g_reset.shape[1] > 9:
                    ws_mg.set_column(9, g_reset.shape[1] - 1, 16)
                img = _render_plot_image("MG", g, well)
                ws_mg.insert_image(cur_row + 1, 9, f"MG_{well}.png", {"image_data": img})
                cur_row = cur_row + 1 + len(g_reset) + 4
        else:
            ws_mg.write(0, 0, "Нет данных MG")

        # --- Chan ---
        ws_ch = writer.book.add_worksheet("Chan")
        writer.sheets["Chan"] = ws_ch
        cur_row = 0
        if chan_df is not None and not chan_df.empty:
            for well, g in chan_df.groupby("well", sort=False):
                title = f"Скважина {well} — Chan"
                g_reset = g.reset_index(drop=True)
                g_reset.to_excel(writer, sheet_name="Chan", index=False, startrow=cur_row + 1, startcol=0)
                ws_ch.write(cur_row, 0, title)
                ws_ch.set_column(0, min(8, g_reset.shape[1] - 1), 14)
                if g_reset.shape[1] > 9:
                    ws_ch.set_column(9, g_reset.shape[1] - 1, 16)
                img = _render_plot_image("Chan", g, well)
                ws_ch.insert_image(cur_row + 1, 9, f"Chan_{well}.png", {"image_data": img})
                cur_row = cur_row + 1 + len(g_reset) + 4
        else:
            ws_ch.write(0, 0, "Нет данных Chan")

    output.seek(0)
    return output


# =========================
# Основной UI/поток
# =========================
def main() -> None:
    # Кнопки скачивания шаблона
    upload_examples()

    # Загрузка данных
    uploaded_file = st.file_uploader(label="**Загрузите данные для расчёта**", accept_multiple_files=False)
    if uploaded_file is None:
        st.info("Пожалуйста, загрузите файл в формате .csv, .txt, .xls, .xlsx")
        return

    if uploaded_file.name.lower().endswith((".txt", ".csv")):
        df_raw = pd.read_csv(uploaded_file)
    elif uploaded_file.name.lower().endswith((".xls", ".xlsx")):
        df_raw = pd.read_excel(uploaded_file)
    else:
        st.error("Неверный формат данных. Загрузите .csv, .txt, .xls, .xlsx")
        return

    # Подготовка
    df = data_preparation(df_raw)

    # MG
    mg_df = compute_mg_full(df)
    st.text(f"[OK] MG рассчитан: строк {len(mg_df)}; скважин {mg_df['well'].nunique() if not mg_df.empty else 0}")

    # Chan
    chan_df = compute_chan_full(df)
    st.text(f"[OK] Chan рассчитан: строк {len(chan_df)}; скважин {chan_df['well'].nunique() if not chan_df.empty else 0}")

    # Вывод по скважинам + сбор сводки
    rows: List[Dict[str, str]] = []
    wells_mg = set(mg_df["well"].unique() if not mg_df.empty else [])
    wells_ch = set(chan_df["well"].unique() if not chan_df.empty else [])
    all_wells = sorted(list(wells_mg.union(wells_ch)))

    for w in all_wells:
        mg_g = mg_df[mg_df["well"] == w] if not mg_df.empty else pd.DataFrame()
        ch_g = chan_df[chan_df["well"] == w] if not chan_df.empty else pd.DataFrame()

        mg_diag = diagnose_mg_group(mg_g) if not mg_g.empty else {"mg_text": "нет данных MG", "mg_detail": ""}
        ch_diag = diagnose_chan_group(ch_g) if not ch_g.empty else {"chan_text": "нет данных Chan", "chan_detail": ""}

        st.markdown(f'<h2 style="color: darkred;">Скважина {w}</h2>', unsafe_allow_html=True)
        st.text(f"  MG:   {mg_diag['mg_text']}")
        if mg_diag["mg_detail"]:
            st.text(f"        {mg_diag['mg_detail']}")
        st.text(f"  Chan: {ch_diag['chan_text']}")
        if ch_diag["chan_detail"]:
            st.text(f"        {ch_diag['chan_detail']}")

        rows.append({"well": w, **mg_diag, **ch_diag})

        # --- График MG ---
        st.markdown(f"##### MG-график (Y vs X) — скважина {w}")
        st.text("Кривая показывает долю накопленной нефти (Y) от накопленной жидкости при увеличении доли накопленной жидкости (X).")
        if not mg_g.empty:
            fig_mg, ax_mg = plt.subplots(figsize=(7, 4))
            ax_mg.scatter(mg_g["MG_X"], mg_g["MG_Y"], label="MG: Y(X)", s=16)
            ax_mg.set_title(f"MG — скважина {w}")
            ax_mg.set_xlabel("X = Qt_cum / Qt_cum(T)")
            ax_mg.set_ylabel("Y = Qo_cum / Qt_cum")
            ax_mg.grid(True, alpha=0.3)
            ax_mg.legend(loc="best")
            st.pyplot(fig_mg, use_container_width=False)
        else:
            st.text(f"  [!] Нет данных MG для скважины {w}")

        # --- График Chan: одна ось, log–log ---
        st.markdown(f"##### Chan-график (WOR и |dWOR/dt|) — скважина {w} (log–log)")
        st.text("Обе кривые на одном графике; оси X и Y — логарифмические. Для производной отображаются только положительные значения.")
        if not ch_g.empty:
            fig_chan, ax = plt.subplots(figsize=(7, 4))
            m_wor = (ch_g["t_pos"] > 0) & (ch_g["WOR"] > 0)
            m_der = (ch_g["t_pos"] > 0) & (ch_g["dWOR_dt_pos"] > 0)
            ax.plot(ch_g.loc[m_wor, "t_pos"], ch_g.loc[m_wor, "WOR"], marker="o", linestyle="none", markersize=4, label="WOR")
            ax.plot(ch_g.loc[m_der, "t_pos"], ch_g.loc[m_der, "dWOR_dt_pos"], linestyle="--", label="|dWOR/dt|")
            ax.set_xscale("log")
            ax.set_yscale("log")
            ax.set_xlabel("t_pos (дни)")
            ax.set_ylabel("WOR, |dWOR/dt|")
            ax.grid(True, which="both", alpha=0.3)
            ax.legend(loc="best")
            ax.set_title(f"Chan — скважина {w} (log–log)")
            st.pyplot(fig_chan, use_container_width=False)
        else:
            st.text(f"  [!] Нет данных Chan для скважины {w}")

    diagnosis_df = pd.DataFrame(rows).sort_values("well").reset_index(drop=True)
    if not diagnosis_df.empty:
        st.markdown(f'<h2 style="color: darkred;">СВОДНАЯ ТАБЛИЦА ДИАГНОЗОВ</h2>', unsafe_allow_html=True)
        st.table(diagnosis_df)
    else:
        st.text("\n[!] Не сформировано ни одного диагноза (возможно, после фильтрации мало валидных точек).")

    # ЕДИНЫЙ EXCEL (Summary + MG + Chan) с графиками
    result_bytes = export_all_results_single_file(mg_df, chan_df, diagnosis_df)
    st.download_button(
        label="Скачать единый файл результатов (Summary + MG + Chan)",
        data=result_bytes,
        file_name="Autodiagnostics_results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# =========================
# Точка входа
# =========================
if __name__ == "__main__":
    main()
