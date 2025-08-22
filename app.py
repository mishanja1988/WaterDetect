from __future__ import annotations

import os
import re
import unicodedata
from dataclasses import dataclass
from io import BytesIO
from typing import Dict, List, Optional

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st
# openpyxl используется для чтения и в резервном механизме экспорта
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows


# =========================
# Глобальные настройки
# =========================
EPS = 1e-9
TEMPLATE_PATH = "data/templates/Сосновское_clean.xlsx"

st.set_page_config(
    layout="wide",
    initial_sidebar_state="auto",
    page_title="Автодиагностика скважин",
    page_icon="🛢️",
)

st.write("### Поскважинный автодиагноз нефтяных скважин по механизму обводнения")

DESCRIPTION_MD = """
**Суть работы:** проведение расчётно-аналитического автодиагноза механизма обводнения по методикам Чена (Chan) и Меркуловой–Гинзбурга (MG) на основе пользовательских исходных данных.

**Что необходимо сделать:**
1. Скачать шаблон исходных данных;
2. Заполнить шаблон своими данными;
3. Загрузить файл в окно подгрузки;
4. Получить текстовый и визуальный диагноз по каждой скважине;
5. Скачать единый Excel c таблицами и **интерактивными графиками**.

**Добавленные столбцы:**
* `Well_calc = H + " " + I`
* `Добыча нефти м3/мес = X * (100 - AB) / 100`
* `Добыча воды м3/мес = X * AB / 100`
* `ВНФ = BT / BS`
* `Накопленное время работы = ЕСЛИ(BR[i]==BR[i-1]; AJ[i] + cum[i-1]; AJ[i])`
* `ВНФ'` — производная по «Накопленному времени»
"""
st.markdown(DESCRIPTION_MD)


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

def series_by_letter(df: pd.DataFrame, letter: str) -> Optional[pd.Series]:
    """Безопасно вернуть серию по букве (или None)."""
    col = col_by_letter(df, letter)
    return df.get(col)

def normalize_header(s: str) -> str:
    if not isinstance(s, str):
        return str(s)
    s = unicodedata.normalize("NFKC", s).replace("\u00A0", " ").replace("\xa0", " ")
    return re.sub(r"\s+", " ", s.strip())

def to_num_or_nan(ser: Optional[pd.Series], df: pd.DataFrame, fill: Optional[float] = None) -> pd.Series:
    """Вернуть числовую серию длиной df. Если ser=None — NaN (или fill, если указан)."""
    if isinstance(ser, pd.Series):
        out = pd.to_numeric(ser, errors="coerce")
    else:
        out = pd.Series(np.nan, index=df.index, dtype=float)
    if fill is not None:
        out = out.fillna(fill)
    return out

def read_template_df() -> pd.DataFrame:
    try:
        if os.path.exists(TEMPLATE_PATH):
            return pd.read_excel(TEMPLATE_PATH)
    except Exception:
        pass
    return pd.DataFrame()

def upload_examples() -> None:
    tpl = read_template_df()
    st.write("**Скачать шаблон исходных данных:**")
    
    out_excel = BytesIO()
    tpl.to_excel(out_excel, index=False, engine="openpyxl")
    out_excel.seek(0)
    
    st.download_button(
        "Скачать шаблон (XLSX)",
        data=out_excel,
        file_name="template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# =========================
# Подготовка данных под MG/Chan
# =========================
def enforce_monotonic_per_well(dfin: pd.DataFrame) -> pd.DataFrame:
    """Обеспечивает строгое возрастание времени для каждой скважины."""
    # groupby().apply() может быть медленным, но здесь важна корректность
    return dfin.groupby("well", sort=False, group_keys=False).apply(
        lambda g: g.assign(t_num=g["t_num"].cummax().add(np.arange(len(g)) * EPS))
    )

def data_preparation(init_data: pd.DataFrame) -> pd.DataFrame:
    df = init_data.copy()
    df.columns = [normalize_header(c) for c in df.columns]

    sH  = series_by_letter(df, "H")
    sI  = series_by_letter(df, "I")
    sX  = series_by_letter(df, "X")
    sAB = series_by_letter(df, "AB")
    sBT = series_by_letter(df, "BT")
    sBS = series_by_letter(df, "BS")
    sBR = series_by_letter(df, "BR")
    sAJ = series_by_letter(df, "AJ")

    # Well_calc
    well_series = pd.Series("", index=df.index, dtype=str)
    if sH is not None:
        well_series += sH.astype(str).str.strip().fillna("")
    if sI is not None:
        well_series = well_series.str.strip() + " " + sI.astype(str).str.strip().fillna("")
    df["well"] = well_series.str.strip()

    # Производные столбцы
    X_vals  = to_num_or_nan(sX,  df)
    AB_vals = to_num_or_nan(sAB, df)
    df["Добыча нефти м3/мес"] = X_vals * (100.0 - AB_vals) / 100.0
    df["Добыча воды м3/мес"]  = X_vals * AB_vals / 100.0

    BT_vals = to_num_or_nan(sBT, df)
    BS_vals = to_num_or_nan(sBS, df)
    with np.errstate(divide="ignore", invalid="ignore"):
        df["ВНФ"] = BT_vals / BS_vals

    # Накопленное время работы (ОПТИМИЗИРОВАНО)
    if sBR is not None and sAJ is not None:
        br_series = sBR.astype(str).fillna("")
        aj_series = pd.to_numeric(sAJ, errors="coerce").fillna(0.0)
        # Определяем группы, где значение BR не меняется
        new_period_marker = (df['well'] != df['well'].shift()) | (br_series != br_series.shift())
        period_group = new_period_marker.cumsum()
        # Считаем кумулятивную сумму внутри каждой скважины и группы периодов
        df["Накопленное время работы"] = aj_series.groupby([df['well'], period_group]).cumsum()
    else:
        df["Накопленное время работы"] = 0.0

    # ВНФ' (производная)
    df = df.sort_values(["well", "Накопленное время работы"]).reset_index(drop=True)
    t_all = pd.to_numeric(df["Накопленное время работы"], errors="coerce")
    y_all = pd.to_numeric(df["ВНФ"], errors="coerce")
    
    # Расчёт градиента внутри каждой группы скважин
    df["ВНФ'"] = df.groupby("well", sort=False).apply(
        lambda g: pd.Series(np.gradient(g["ВНФ"], g["Накопленное время работы"]), index=g.index)
    ).reset_index(level=0, drop=True)
    
    # Объёмы периода и суточные дебиты
    df["qo_period"] = pd.to_numeric(df["Добыча нефти м3/мес"], errors="coerce").fillna(0.0)
    df["qw_period"] = pd.to_numeric(df["Добыча воды м3/мес"],  errors="coerce").fillna(0.0)
    df["qL_period"] = df["qo_period"] + df["qw_period"]

    df["prod_days"] = to_num_or_nan(sAJ, df, fill=0.0)
    with np.errstate(divide="ignore", invalid="ignore"):
        df["qo"] = df["qo_period"] / df["prod_days"]
        df["qw"] = df["qw_period"] / df["prod_days"]
        df["qL"] = df["qL_period"] / df["prod_days"]

    df["t_num"] = df["Накопленное время работы"]

    # Порядок и монотонность
    df = df.dropna(subset=["well", "t_num"]).sort_values(["well", "t_num"]).reset_index(drop=True)
    df = enforce_monotonic_per_well(df)
    return df

# =========================
# MG (без изменений)
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
        d["fw"] = d["qw_period"] / d["qL_period"]
        d["fw"] = d["fw"].replace([np.inf, -np.inf], np.nan)

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
        
        X = g2["Qt_cum"] / Qt_T
        X_mono = X.cummax().add(np.arange(len(X)) * EPS)
        g2["MG_X"] = X_mono
        with np.errstate(invalid="ignore", divide="ignore"):
            g2["MG_Y"] = g2["Qo_cum"] / g2["Qt_cum"]

        flags = MGFlags()
        early_mask = g2["MG_X"] <= 0.2
        if early_mask.sum() >= 3:
            flags.y_early_mean = float(np.nanmean(g2.loc[early_mask, "MG_Y"]))
            flags.possible_behind_casing = (flags.y_early_mean is not None) and (flags.y_early_mean >= 0.99)

        first_third = g2[g2["MG_X"] <= 0.33]
        if len(first_third) >= 3:
            try:
                k, _ = np.polyfit(first_third["MG_X"], first_third["MG_Y"], 1)
                flags.slope_first_third = float(k)
                flags.possible_channeling = (k < -0.8)
            except np.linalg.LinAlgError: pass
        
        if len(g2) >= 5:
            with np.errstate(invalid="ignore"):
                dy = np.gradient(g2["MG_Y"], g2["MG_X"])
            flags.waviness_std = float(np.nanstd(dy))
            flags.possible_mixed_causes = flags.waviness_std > 1.0

        for key, val in vars(flags).items():
            g2[f"MG_diag_{key}"] = val
        frames.append(g2)

    return pd.concat(frames, axis=0).reset_index(drop=True) if frames else pd.DataFrame()

# =========================
# Chan (без изменений)
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
            g["dWOR_dt"] = np.gradient(g["WOR"], g["t_pos"])
        
        mask = (g["WOR"] > 0) & (g["t_pos"] > 0)
        a = np.nan
        if mask.sum() >= 3:
            x = np.log(g.loc[mask, "t_pos"])
            y = np.log(g.loc[mask, "WOR"])
            try:
                a, _ = np.polyfit(x, y, 1)
            except np.linalg.LinAlgError: pass
        
        flags = ChanFlags()
        flags.slope_logWOR_logt = float(a)
        flags.mean_derivative = float(np.nanmean(g["dWOR_dt"]))
        flags.std_derivative = float(np.nanstd(g["dWOR_dt"]))
        
        if not np.isnan(a):
            flags.possible_coning = a > 0.5 and flags.mean_derivative > 0
            flags.possible_near_wellbore = a > 1.0 and flags.mean_derivative > 0
            flags.possible_multilayer_channeling = a > 0 and flags.std_derivative > 0.1

        for key, val in vars(flags).items():
            g[f"chan_diag_{key}"] = val
        g["dWOR_dt_pos"] = np.where(g["dWOR_dt"] > 0, g["dWOR_dt"], np.nan)
        frames.append(g)

    return pd.concat(frames, axis=0).reset_index(drop=True) if frames else pd.DataFrame()


# =========================
# Текстовые диагнозы (адаптировано под новые имена столбцов)
# =========================
def diagnose_mg_group(g: pd.DataFrame) -> Dict[str, str]:
    if g.empty: return {"mg_text": "нет данных MG", "mg_detail": ""}
    last_row = g.iloc[-1]
    y_early = last_row.get("MG_diag_y_early_mean", np.nan)
    slope = last_row.get("MG_diag_slope_first_third", np.nan)
    wav = last_row.get("MG_diag_waviness_std", np.nan)

    parts: List[str] = []
    if last_row.get("MG_diag_possible_behind_casing"): parts.append("возможны заколонные перетоки (ранний нефтеотбор Y≈1)")
    if last_row.get("MG_diag_possible_channeling"): parts.append("признаки каналирования (крутой спад Y в первой трети)")
    if last_row.get("MG_diag_possible_mixed_causes"): parts.append("смешанные причины (высокая волнистость dY/dX)")
    if not parts: parts.append("характеристика ближе к равномерному обводнению")
    
    detail = f"MG метрики: y_early≈{y_early:.2f}; наклон≈{slope:.2f}; волнистость≈{wav:.2f}"
    return {"mg_text": "; ".join(parts), "mg_detail": detail}


def diagnose_chan_group(g: pd.DataFrame) -> Dict[str, str]:
    if g.empty: return {"chan_text": "нет данных Chan", "chan_detail": ""}
    last_row = g.iloc[-1]
    slope = last_row.get("chan_diag_slope_logWOR_logt", np.nan)
    mean_d = last_row.get("chan_diag_mean_derivative", np.nan)
    std_d = last_row.get("chan_diag_std_derivative", np.nan)

    parts: List[str] = []
    if last_row.get("chan_diag_possible_multilayer_channeling"): parts.append("многослойное каналирование (рост WOR и дисперсии производной)")
    if last_row.get("chan_diag_possible_near_wellbore"): parts.append("приствольные проблемы/ранний канал (очень высокий наклон)")
    if last_row.get("chan_diag_possible_coning"): parts.append("возможен конинг (наклон > 0.5 при положительной производной)")
    if not parts: parts.append("нет выраженных признаков проблемного притока воды")
    
    detail = f"Chan метрики: наклон≈{slope:.2f}; средн. dWOR/dt≈{mean_d:.2e}; std≈{std_d:.2e}"
    return {"chan_text": "; ".join(parts), "chan_detail": detail}


# =========================
# Экспорт с нативными Excel-графиками (ПЕРЕРАБОТАНО)
# =========================
def export_all_results_single_file(mg_df: pd.DataFrame, chan_df: pd.DataFrame, diagnosis_df: pd.DataFrame) -> BytesIO:
    """Создает XLSX с тремя листами и нативными Excel-графиками."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        
        # 1. Лист "Summary"
        diagnosis_df.to_excel(writer, sheet_name="Summary", index=False)
        writer.sheets["Summary"].autofit()
        
        # 2. Лист "MG" с данными и графиками
        ws_mg = workbook.add_worksheet("MG")
        writer.sheets["MG"] = ws_mg
        current_row = 0
        
        if mg_df is not None and not mg_df.empty:
            for well, g in mg_df.groupby("well", sort=False):
                ws_mg.write(current_row, 0, f"Скважина {well} — MG")
                current_row += 1
                
                g_reset = g.reset_index(drop=True)
                g_reset.to_excel(writer, sheet_name="MG", index=False, startrow=current_row)
                
                # Создание нативного графика
                chart = workbook.add_chart({'type': 'scatter'})
                
                num_points = len(g_reset)
                col_x_idx = g_reset.columns.get_loc("MG_X") + 1
                col_y_idx = g_reset.columns.get_loc("MG_Y") + 1
                
                chart.add_series({
                    'name':       f'Скважина {well}',
                    'categories': ['MG', current_row + 1, col_x_idx, current_row + num_points, col_x_idx],
                    'values':     ['MG', current_row + 1, col_y_idx, current_row + num_points, col_y_idx],
                    'marker':     {'type': 'circle', 'size': 5},
                })
                chart.set_title({'name': f'MG — Скважина {well}'})
                chart.set_x_axis({'name': 'X = Qt_cum / Qt_cum(T)'})
                chart.set_y_axis({'name': 'Y = Qo_cum / Qt_cum'})
                chart.set_legend({'position': 'none'})
                
                ws_mg.insert_chart(current_row, g_reset.shape[1] + 1, chart)
                current_row += len(g_reset) + 5
        else:
            ws_mg.write(0, 0, "Нет данных MG")
            
        # 3. Лист "Chan" с данными и графиками
        ws_ch = workbook.add_worksheet("Chan")
        writer.sheets["Chan"] = ws_ch
        current_row = 0

        if chan_df is not None and not chan_df.empty:
            for well, g in chan_df.groupby("well", sort=False):
                ws_ch.write(current_row, 0, f"Скважина {well} — Chan")
                current_row += 1

                g_reset = g.reset_index(drop=True)
                g_reset.to_excel(writer, sheet_name="Chan", index=False, startrow=current_row)

                chart = workbook.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})

                num_points = len(g_reset)
                col_t_idx = g_reset.columns.get_loc("t_pos") + 1
                col_wor_idx = g_reset.columns.get_loc("WOR") + 1
                col_dw_idx = g_reset.columns.get_loc("dWOR_dt_pos") + 1

                chart.add_series({
                    'name':       'WOR',
                    'categories': ['Chan', current_row + 1, col_t_idx, current_row + num_points, col_t_idx],
                    'values':     ['Chan', current_row + 1, col_wor_idx, current_row + num_points, col_wor_idx],
                    'marker':     {'type': 'circle', 'size': 5},
                    'line':       {'none': True},
                })
                chart.add_series({
                    'name':       '|dWOR/dt|',
                    'categories': ['Chan', current_row + 1, col_t_idx, current_row + num_points, col_t_idx],
                    'values':     ['Chan', current_row + 1, col_dw_idx, current_row + num_points, col_dw_idx],
                    'marker':     {'type': 'none'},
                    'line':       {'dash_type': 'dash'},
                })
                chart.set_title({'name': f'Chan — Скважина {well}'})
                chart.set_x_axis({'name': 't_pos (дни)', 'log_base': 10})
                chart.set_y_axis({'name': 'WOR, |dWOR/dt|', 'log_base': 10})
                
                ws_ch.insert_chart(current_row, g_reset.shape[1] + 1, chart)
                current_row += len(g_reset) + 5
        else:
            ws_ch.write(0, 0, "Нет данных Chan")

    output.seek(0)
    return output

# =========================
# Основной UI/поток
# =========================
def main() -> None:
    upload_examples()
    uploaded_file = st.file_uploader(label="**Загрузите XLSX/XLS файл для расчёта**", type=["xlsx", "xls"])
    
    if uploaded_file is None:
        st.info("Пожалуйста, загрузите файл, созданный на основе шаблона.")
        return

    try:
        with st.spinner("Чтение и обработка данных..."):
            df_raw = pd.read_excel(uploaded_file)
            df = data_preparation(df_raw)
        
        with st.spinner("Расчёт по методике Меркуловой-Гинзбурга..."):
            mg_df = compute_mg_full(df)
        st.success(f"✔️ MG: Расчёт выполнен для {mg_df['well'].nunique() if not mg_df.empty else 0} скважин.")
        
        with st.spinner("Расчёт по методике Chan..."):
            chan_df = compute_chan_full(df)
        st.success(f"✔️ Chan: Расчёт выполнен для {chan_df['well'].nunique() if not chan_df.empty else 0} скважин.")

        rows: List[Dict[str, str]] = []
        all_wells = sorted(list(set(mg_df["well"].unique()) | set(chan_df["well"].unique())))

        if not all_wells:
            st.warning("Не найдено скважин для анализа после обработки данных. Проверьте входной файл.")
            return

        for w in all_wells:
            mg_g = mg_df[mg_df["well"] == w]
            ch_g = chan_df[chan_df["well"] == w]
            
            mg_diag = diagnose_mg_group(mg_g)
            ch_diag = diagnose_chan_group(ch_g)
            rows.append({"well": w, **mg_diag, **ch_diag})

            with st.expander(f"Диагноз и графики для скважины: {w}"):
                st.markdown(f"#### 📜 Диагноз: {w}")
                col1, col2 = st.columns(2)
                col1.metric("Диагноз MG", mg_diag['mg_text'], help=mg_diag['mg_detail'])
                col2.metric("Диагноз Chan", ch_diag['chan_text'], help=ch_diag['chan_detail'])
                
                st.markdown(f"#### 📈 Графики: {w}")
                plot_col1, plot_col2 = st.columns(2)
                
                with plot_col1:
                    if not mg_g.empty:
                        fig_mg, ax_mg = plt.subplots()
                        ax_mg.scatter(mg_g["MG_X"], mg_g["MG_Y"], s=16)
                        ax_mg.set_title(f"MG — скважина {w}")
                        ax_mg.set_xlabel("X = Qt_cum / Qt_cum(T)")
                        ax_mg.set_ylabel("Y = Qo_cum / Qt_cum")
                        ax_mg.grid(True, alpha=0.3)
                        st.pyplot(fig_mg)
                    else:
                        st.info(f"Нет данных MG для {w}")

                with plot_col2:
                    if not ch_g.empty:
                        fig_chan, ax = plt.subplots()
                        ax.plot(ch_g["t_pos"], ch_g["WOR"], "o", markersize=4, label="WOR")
                        ax.plot(ch_g["t_pos"], ch_g["dWOR_dt_pos"], "--", label="|dWOR/dt|")
                        ax.set_xscale("log"); ax.set_yscale("log")
                        ax.set_xlabel("t_pos (дни)"); ax.set_ylabel("WOR, |dWOR/dt|")
                        ax.grid(True, which="both", alpha=0.3); ax.legend()
                        ax.set_title(f"Chan — скважина {w} (log–log)")
                        st.pyplot(fig_chan)
                    else:
                        st.info(f"Нет данных Chan для {w}")

        diagnosis_df = pd.DataFrame(rows).sort_values("well").reset_index(drop=True)
        if not diagnosis_df.empty:
            st.markdown("---")
            st.subheader("Сводная таблица диагнозов")
            st.dataframe(diagnosis_df)
        
        st.markdown("---")
        st.subheader("📥 Скачать результаты")
        result_bytes = export_all_results_single_file(mg_df, chan_df, diagnosis_df)
        st.download_button(
            label="Скачать единый Excel-файл (таблицы + интерактивные графики)",
            data=result_bytes,
            file_name="Autodiagnostics_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"Произошла ошибка при обработке файла: {e}")
        st.warning("Убедитесь, что структура файла соответствует шаблону.")

# =========================
# Точка входа
# =========================
if __name__ == "__main__":
    main()
