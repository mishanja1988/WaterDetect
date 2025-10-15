from IPython.display import display, Markdown
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st
from PIL import Image
import io, re, unicodedata, requests, sys
from dataclasses import dataclass
from typing import Optional, List, Dict
from io import BytesIO
import sys
import os

import plotly.graph_objects as go
from plotly.subplots import make_subplots

# Переменные
EPS = 1e-9      



st.set_page_config(
        layout='wide', 
        initial_sidebar_state='auto', 
        page_title='Автодиагностика скважин', 
        page_icon='image'
    )

st.write('### Поскважинный автодиагноз нефтяных скважин по механизму обводнения')
st.markdown(
        '''
        **Суть работы:** проведение расчетно-аналитического способов механизма обводнения скважин с использованием методики Чена (Chan) и Меркуловой–Гинзбурга (MG) по нефтяным скважинам на основе пользовательских исходных данных:\n\n
        **Что необходимо сделать:**  
            1. Загрузить шалон для заполнения исходных данных;  
            2. Заполнить шаблон своими данными;  
            3. Подгрузить Ваш шаблон в окно подгрузки данных;  
            4. Получить результат - текстовый и визуальный автодиагноз по каждой скважине;  
            5. Скачать результирующие таблицы для анализа.  
        '''
        )


# Степаненко ИБ - Функция чтения данных примеров загрузки для моделей:
@st.cache_data
def read_examples():
    example_csv = pd.read_csv('data/templates/df_raw.csv', skiprows=[1]).drop(columns='Unnamed: 0.1')
    example_excel = pd.read_excel('data/templates/df_raw.xlsx', skiprows=[1]).drop(columns='Unnamed: 0.1')
    return example_csv, example_excel

# Степаненко ИБ - Функций перекодировки эксель:
def save_df_to_excel(df, ind=False):
    output = BytesIO()
    df.to_excel(output, index=ind, engine='openpyxl')
    output.seek(0)
    return output

# Степаненко ИБ - функция кнопок выгрузки примеров загрузки для моделей
def upload_examples(): 
    global example_csv, example_excel
    st.write('**Скачать пример таблицы для подачи расчетов алгоритмов:**')
    col1, col2, col3, col4, col5, col6, col7, col8, col9,  = st.columns(9)
    button_txt = col1.download_button(
        label='Скачать таблицу в .csv', 
        data = example_csv.to_csv(index=False), 
        file_name='example_text.csv',
        mime='text/csv'
        ) 
    button_excel = col2.download_button(
        label="Скачать таблицу в .xlsx",
        data=save_df_to_excel(example_excel),
        file_name='example_excel.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
    )
    if button_txt or button_excel:
        st.success("Таблица примера успешно сохранена в загрузки")


# Степаненко ИБ - Вызов функций кнопок выгрузки результатов: 
example_csv, example_excel = read_examples()
upload_examples()




def normalize_header(s: str) -> str:
    if not isinstance(s, str):
        return s
    s = unicodedata.normalize("NFKC", s).replace("\u00A0"," ").replace("\xa0"," ")
    s = re.sub(r"\s+", " ", s.strip())
    return s


# Агупов МА - Функция соединения:
def enforce_monotonic_per_well(dfin: pd.DataFrame) -> pd.DataFrame:
    out = []
    for w, g in dfin.groupby("well", sort=False):
        t = g["t_num"].to_numpy(dtype=float)
        for i in range(1, t.size):
            if t[i] <= t[i-1]:
                t[i] = t[i-1] + EPS
        g = g.copy(); g["t_num"] = t
        out.append(g)
    return pd.concat(out, axis=0).reset_index(drop=True)


# Агупов МА - Функция основная по преобразованию данных:
def data_preparation(init_data):
    dfn = init_data.copy()
    dfn.columns = [normalize_header(c) for c in dfn.columns]
    COLUMN_MAP = {
        "well": "Unnamed: 0",
        "date_or_period": "Накопленно времени",
        "qo": "Дебит нефти, м3/сут",
        "qw": "Дебит воды м3/сут",
        "qL": "Дебит жидкости, м3/сут",
        "prod_days": "Число дней добычи нефти, сут"
    }
    df = dfn.rename(columns={v:k for k,v in COLUMN_MAP.items() if v in dfn.columns}).copy()
    # Автодетект well
    if "well" not in df.columns:
        lower_cols = {c.lower(): c for c in dfn.columns}
        hit = None
        for token in ["скважина","скв","скв.","well","id","№","номер"]:
            cand = [orig for lc, orig in lower_cols.items() if token in lc]
            if cand: hit = cand[0]; break
        if hit is None:
            hit = next((c for c in dfn.columns if str(c).lower().startswith("unnamed")), dfn.columns[0])
        df["well"] = dfn[hit]

    # Остальные поля — мягкий маппинг
    if "date_or_period" not in df.columns:
        for token in ["накопленно времени","дата (месяц, год)","период","месяц","дни работы"]:
            hit = [c for c in dfn.columns if token == c.lower()]
            if hit: df["date_or_period"] = dfn[hit[0]]; break

    if "qo" not in df.columns:
        for c in dfn.columns:
            if "неф" in c.lower():
                df["qo"] = pd.to_numeric(dfn[c], errors="coerce"); break
    if "qw" not in df.columns:
        for c in dfn.columns:
            if "вод" in c.lower():
                df["qw"] = pd.to_numeric(dfn[c], errors="coerce"); break
    if "qL" not in df.columns:
        hit = None
        for c in dfn.columns:
            if "жидк" in c.lower():
                hit = c; break
        if hit: df["qL"] = pd.to_numeric(dfn[hit], errors="coerce")
    if "prod_days" not in df.columns:
        for c in dfn.columns:
            if "число дней добычи" in c.lower() or "сут" in c.lower():
                ser = pd.to_numeric(dfn[c], errors="coerce")
                if ser.notna().sum() >= len(ser)*0.5:
                    df["prod_days"] = ser; break

    for c in ["qo","qw","qL","prod_days"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    # Время без агрегаций
    if "date_or_period" in df.columns:
        try:
            t = pd.to_datetime(df["date_or_period"])
            df["t_num"] = (t - t.groupby(df["well"]).transform("min")).dt.days.astype(float)
        except Exception:
            df["t_num"] = pd.to_numeric(df["date_or_period"], errors="coerce").astype(float)
    else:
        df["t_num"] = np.arange(len(df), dtype=float)

    # Жидкость и объёмы за период
    if "qL" not in df.columns:
        df["qL"] = df.get("qo", 0).fillna(0) + df.get("qw", 0).fillna(0)
    if "prod_days" in df.columns:
        df["qo_period"] = df.get("qo", 0) * df["prod_days"].fillna(1.0)
        df["qw_period"] = df.get("qw", 0) * df["prod_days"].fillna(1.0)
        df["qL_period"] = df["qL"] * df["prod_days"].fillna(1.0)
    else:
        df["qo_period"] = df.get("qo", 0)
        df["qw_period"] = df.get("qw", 0)
        df["qL_period"] = df["qL"]

    # Упорядочить и сделать монотонным время по скважине (без удаления строк)
    df = df.dropna(subset=["well","t_num"]).sort_values(["well","t_num"]).reset_index(drop=True)
    df = enforce_monotonic_per_well(df)
    return df


# Агупов МА - Функция Меркулова–Гинзбург (MG): расчёт и флаги
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
    d["fw"] = np.where(d["qL_period"]>0, d["qw_period"]/d["qL_period"], np.nan)
    frames = []
    for w, g in d.groupby("well", sort=False):
        g = g.sort_values("t_num").copy()
        idx = g.index[g["fw"] > watercut_thr]
        if len(idx)==0 or len(g)<min_points:
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
            if X[i] <= X[i-1]:
                X[i] = X[i-1] + EPS
        g2["MG_X"] = X
        g2["MG_Y"] = np.where(g2["Qt_cum"]>0, g2["Qo_cum"]/g2["Qt_cum"], np.nan)

        flags = MGFlags()
        early_mask = g2["MG_X"]<=0.2
        if early_mask.sum()>=3:
            flags.y_early_mean = float(np.nanmean(g2.loc[early_mask,"MG_Y"]))
            flags.possible_behind_casing = (flags.y_early_mean is not None) and (flags.y_early_mean >= 0.99)

        first_third = g2[g2["MG_X"]<=0.33]
        if len(first_third)>=3:
            x = first_third["MG_X"].to_numpy()
            y = first_third["MG_Y"].to_numpy()
            A = np.vstack([x, np.ones_like(x)]).T
            try:
                k, b = np.linalg.lstsq(A, y, rcond=None)[0]
                flags.slope_first_third = float(k)
                flags.possible_channeling = (k < -0.8)
            except Exception:
                pass

        if len(g2)>=5:
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



# Агупов МА - Функция Чена (Chan): WOR и производная, расчёт и флаги
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
        g["WOR"] = g["qw"] / g["qo"]
        g = g.replace([np.inf,-np.inf], np.nan)
        g = g[(g["qo"]>0) & (g["WOR"]>0)].dropna(subset=["WOR"])
        if len(g) < min_points:
            continue

        with np.errstate(invalid="ignore"):
            g["t_pos"] = g["t_num"] - g["t_num"].min() + EPS
            g["dWOR_dt"] = np.gradient(g["WOR"].to_numpy(), g["t_pos"].to_numpy())

        mask = (g["WOR"]>0) & (g["t_pos"]>0)
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
        std_deriv  = float(np.nanstd(g["dWOR_dt"])) if len(g) else np.nan

        g["well"] = w
        g["chan_slope_logWOR_logt"] = float(a) if a==a else np.nan
        g["chan_mean_dWOR_dt"] = mean_deriv
        g["chan_std_dWOR_dt"] = std_deriv
        g["chan_flag_coning"] = (a > 0.5 and mean_deriv > 0) if a==a else False
        g["chan_flag_near_wellbore"] = (a > 1.0 and mean_deriv > 0) if a==a else False
        g["chan_flag_multilayer_channeling"] = (a > 0 and std_deriv > 0.1) if a==a else False
        frames.append(g)
    return pd.concat(frames, axis=0).reset_index(drop=True) if frames else pd.DataFrame()


# Агупов МА - Функция автодиагноза по каждой скважине (Меркулова–Гинзбург (MG))
def diagnose_mg_group(g: pd.DataFrame) -> Dict[str, str]:
    y_early = g["MG_diag_y_early_mean"].dropna().iloc[-1] if "MG_diag_y_early_mean" in g and g["MG_diag_y_early_mean"].notna().any() else np.nan
    slope   = g["MG_diag_slope_first_third"].dropna().iloc[-1] if "MG_diag_slope_first_third" in g and g["MG_diag_slope_first_third"].notna().any() else np.nan
    wav     = g["MG_diag_waviness_std"].dropna().iloc[-1] if "MG_diag_waviness_std" in g and g["MG_diag_waviness_std"].notna().any() else np.nan
    f_bc = bool(g["MG_flag_behind_casing"].dropna().iloc[-1]) if "MG_flag_behind_casing" in g and g["MG_flag_behind_casing"].notna().any() else False
    f_ch = bool(g["MG_flag_channeling"].dropna().iloc[-1]) if "MG_flag_channeling" in g and g["MG_flag_channeling"].notna().any() else False
    f_mix= bool(g["MG_flag_mixed"].dropna().iloc[-1]) if "MG_flag_mixed" in g and g["MG_flag_mixed"].notna().any() else False

    parts: List[str] = []
    if f_bc:  parts.append("возможны заколонные перетоки (ранний нефтеотбор Y≈1)")
    if f_ch:  parts.append("признаки каналирования (крутой спад Y в первой трети)")
    if f_mix: parts.append("смешанные причины (высокая волнистость dY/dX)")
    if not parts: parts.append("характеристика ближе к равномерному обводнению")

    detail = f"MG метрики: y_early≈{y_early:.2f}; наклон≈{slope:.2f}; волнистость≈{wav:.2f}"
    return {"mg_text": "; ".join(parts), "mg_detail": detail}

# Агупов МА - Функция автодиагноза по каждой скважине (Чена (Chan))
def diagnose_chan_group(g: pd.DataFrame) -> Dict[str, str]:
    slope  = g["chan_slope_logWOR_logt"].dropna().iloc[-1] if "chan_slope_logWOR_logt" in g and g["chan_slope_logWOR_logt"].notna().any() else np.nan
    mean_d = g["chan_mean_dWOR_dt"].dropna().iloc[-1] if "chan_mean_dWOR_dt" in g and g["chan_mean_dWOR_dt"].notna().any() else np.nan
    std_d  = g["chan_std_dWOR_dt"].dropna().iloc[-1]  if "chan_std_dWOR_dt"  in g and g["chan_std_dWOR_dt"].notna().any()  else False
    f_cone = bool(g["chan_flag_coning"].dropna().iloc[-1]) if "chan_flag_coning" in g and g["chan_flag_coning"].notna().any() else False
    f_near = bool(g["chan_flag_near_wellbore"].dropna().iloc[-1]) if "chan_flag_near_wellbore" in g and g["chan_flag_near_wellbore"].notna().any() else False
    f_multi= bool(g["chan_flag_multilayer_channeling"].dropna().iloc[-1]) if "chan_flag_multilayer_channeling" in g and g["chan_flag_multilayer_channeling"].notna().any() else False

    parts: List[str] = []
    if f_multi: parts.append("многослойное каналирование (рост WOR и дисперсии производной)")
    if f_near:  parts.append("приствольные проблемы/ранний канал (очень высокий наклон)")
    if f_cone:  parts.append("возможен конинг (наклон > 0.5 при положительной производной)")
    if not parts: parts.append("нет выраженных признаков проблемного притока воды")

    detail = f"Chan метрики: наклон≈{slope:.2f}; средн. dWOR/dt≈{mean_d:.2e}; std≈{std_d:.2e}"
    return {"chan_text": "; ".join(parts), "chan_detail": detail}


def show():
    # 1. Подгрузка данных:
    uploaded_file = st.file_uploader(label='**Загрузите данные для расчета**', accept_multiple_files=False) 
    if uploaded_file is None:
        st.info("Пожалуйста, загрузите файл в формате  .csv, .txt, .xls, .xlsx")  
        return
    
    if '.txt' in uploaded_file.name or '.csv' in uploaded_file.name: 
        df_raw = pd.read_csv(uploaded_file)
    elif '.xls' in uploaded_file.name or '.xlsx' in uploaded_file.name: 
        df_raw = pd.read_excel(uploaded_file)# .drop(columns='Unnamed: 0.1')
    else:
        st.error('Неправильный формат данных, подгрузите данные в формате .csv, .txt, .xls, .xlsx')
        return

    
    # 2. Функция преобразования данных:
    df = data_preparation(df_raw)
    
    # 3. Функция гингзбурга:
    mg_df = compute_mg_full(df)
    st.text(f"[OK] MG рассчитан: строк {len(mg_df)}; скважин {mg_df['well'].nunique() if not mg_df.empty else 0}")
    
    # 4. Функция Чена:
    chan_df = compute_chan_full(df)
    st.text(f"[OK] Chan рассчитан: строк {len(chan_df)}; скважин {chan_df['well'].nunique() if not chan_df.empty else 0}")

    
    upload_result(mg_df, chan_df)
    
    # 5. Функция вывода результатов:
    rows = []
    all_wells = sorted(list(set(mg_df["well"].unique() if not mg_df.empty else []).union(
                    set(chan_df["well"].unique() if not chan_df.empty else []))))

    for w in all_wells:
        mg_g = mg_df[mg_df["well"]==w] if not mg_df.empty else pd.DataFrame()
        ch_g = chan_df[chan_df["well"]==w] if not chan_df.empty else pd.DataFrame()

        mg_diag = diagnose_mg_group(mg_g) if not mg_g.empty else {"mg_text":"нет данных MG","mg_detail":""}
        ch_diag = diagnose_chan_group(ch_g) if not ch_g.empty else {"chan_text":"нет данных Chan","chan_detail":""}

        st.markdown(f'<h2 style="color: darkred;">Скважина {w}</h2>', unsafe_allow_html=True) # st.markdown(f"## Скважина {w}:")
        st.text(f"  MG:   {mg_diag['mg_text']}")
        if mg_diag['mg_detail']: st.text(f"        {mg_diag['mg_detail']}")
        st.text(f"  Chan: {ch_diag['chan_text']}")
        if ch_diag['chan_detail']: st.text(f"        {ch_diag['chan_detail']}")

        rows.append({"well": w, **mg_diag, **ch_diag})

        # --- График MG ---
        st.markdown(f"##### MG-график (Y vs X) — скважина {w}")
        st.text("Кривая показывает долю накопленной нефти (Y) от накопленной жидкости при увеличении доли накопленной жидкости (X). Форма кривой позволяет судить о механизме обводнения.")
        
        if not mg_g.empty:
            fig_mg, ax_mg = plt.subplots(figsize=(6, 3))
            ax_mg.scatter(mg_g['MG_X'], mg_g['MG_Y'])
            ax_mg.set_title(f'MG — скважина {w}')
            ax_mg.set_xlabel('X = Qt_cum / Qt_cum(T)')
            ax_mg.set_ylabel('Y = Qo_cum / Qt_cum')
            ax_mg.grid(True, alpha=0.3)
            st.pyplot(fig_mg, use_container_width=False)
        else:
            st.text(f"  [!] Нет данных MG для скважины {w}")

    #     # --- График Chan ---
    #     st.markdown(f"##### Chan-график (WOR и dWOR/dt) — скважина {w}")
    #     st.text("WOR показывает относительный рост воды; производная dWOR/dt подсвечивает скорость изменений. По совместной динамике можно отличать конинг/каналирование/приствольные эффекты.")
    #     if not ch_g.empty:
    #         fig_chan, ax1 = plt.subplots(figsize=(6, 3))
    #         ax1.set_xlabel('t_pos (дни)')
    #         ax1.set_ylabel('WOR')
    #         ax1.scatter(ch_g['t_pos'], ch_g['WOR'], label='WOR')
    #         ax1.grid(True, alpha=0.3)
    #         ax2 = ax1.twinx()
    #         ax2.set_ylabel('dWOR/dt')
    #         ax2.plot(ch_g['t_pos'], ch_g['dWOR_dt'], label='dWOR/dt', linestyle='--')
    #         lines, labels = ax1.get_legend_handles_labels()
    #         lines2, labels2 = ax2.get_legend_handles_labels()
    #         ax2.legend(lines + lines2, labels + labels2, loc='upper left')
    #         plt.title(f'Chan — скважина {w}')
    #         st.pyplot(fig_chan, use_container_width=False)
    #     else:
    #         st.text(f"  [!] Нет данных Chan для скважины {w}")
                
    # diagnosis_df = pd.DataFrame(rows).sort_values("well").reset_index(drop=True)


        # --- График Chan ---
        st.markdown(f"##### Chan-график (WOR и dWOR/dt) — скважина {w}")
        st.text("WOR показывает относительный рост воды; производная dWOR/dt подсвечивает скорость изменений. По совместной динамике можно отличать конинг/каналирование/приствольные эффекты.")
        if not ch_g.empty:
            fig_chan, ax1 = plt.subplots(figsize=(6, 3))
            ax1.set_xlabel('t_pos (дни)')
            ax1.set_ylabel('WOR')
            ax1.scatter(ch_g['t_pos'], ch_g['WOR'], label='WOR')
            ax1.grid(True, alpha=0.3)
            ax1.set_xscale('log') # Set x-axis to logarithmic scale
            ax1.set_yscale('log') # Set y-axis to logarithmic scale
            ax2 = ax1.twinx()
            ax2.set_ylabel('dWOR/dt')
            ax2.plot(ch_g['t_pos'], ch_g['dWOR_dt'], label='dWOR/dt', linestyle='--')
            ax2.set_xscale('log') # Set x-axis to logarithmic scale
            # ax2.set_yscale('log') # Do not set y-axis to logarithmic scale for derivative
            lines, labels = ax1.get_legend_handles_labels()
            lines2, labels2 = ax2.get_legend_handles_labels()
            ax2.legend(lines + lines2, labels + labels2, loc='upper left')
            plt.title(f'Chan — скважина {w}')
            st.pyplot(fig_chan, use_container_width=False)
        else:
            st.text(f"  [!] Нет данных Chan для скважины {w}")
        
    diagnosis_df = pd.DataFrame(rows).sort_values("well").reset_index(drop=True)


    if not diagnosis_df.empty:
        st.markdown(f'<h2 style="color: darkred;">СВОДНАЯ ТАБЛИЦА ДИАГНОЗОВ</h2>', unsafe_allow_html=True)
        st.table(diagnosis_df)
    else:
        st.text("\n[!] Не сформировано ни одного диагноза (возможно, после фильтрации мало валидных точек).")





def upload_result(df_MG, df_Chan): 
    c1, c2, c3, c4, c5, c6, c7, c8, c9,  = st.columns(9)
    but_excel_MG = c1.download_button(
        label = "Скачать результаты Меркуловой–Гинзбург (MG)", 
        data = save_df_to_excel(df_MG, ind=True),
        file_name='MG_results.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
        )
    but_excel_Chan = c2.download_button(
        label = "Скачать результаты в Чена (Chan)", 
        data = save_df_to_excel(df_Chan, ind=True),
        file_name='Chan_results.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
        )
    if but_excel_MG or but_excel_Chan:
        st.success("Таблица результатов успешно сохранена в загрузки")
        
        

if __name__ == '__main__':

    show()
