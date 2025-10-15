# app.py — Автодиагностика скважин (всё в одном файле)

from __future__ import annotations

import os
import re
import unicodedata
from dataclasses import dataclass
from typing import Dict, List, Optional
from io import BytesIO


import numpy as np
import pandas as pd
import streamlit as st


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
5. Скачать единый Excel c таблицами и графиками.

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
# Основной UI/поток
# =========================
def main() -> None:
    # Кнопки скачивания шаблона
    upload_examples()

    # Загрузка данных
    uploaded_file = st.file_uploader(label="**Загрузите данные для расчёта**", accept_multiple_files=False)
    if uploaded_file is None:
        st.info("Поддерживаются .csv, .txt, .xls, .xlsx")
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

        # --- График Chan: одна ось, обе шкалы log ---
        st.markdown(f"##### Chan-график (WOR и |dWOR/dt|) — скважина {w} (log–log)")
        st.text("Обе кривые на одном графике; оси X и Y — логарифмические. Для производной отображаются только положительные значения.")
        if not ch_g.empty:
            fig_chan, ax = plt.subplots(figsize=(7, 4))
            m_wor = (ch_g["t_pos"] > 0) & (ch_g["WOR"] > 0)
            m_der = (ch_g["t_pos"] > 0) & (ch_g["dWOR_dt_pos"] > 0)
            ax.plot(ch_g.loc[m_wor, "t_pos"], ch_g.loc[m_wor, "WOR"], marker="o", linestyle="none", label="WOR", markersize=4)
            ax.plot(ch_g.loc[m_der, "t_pos"], ch_g.loc[m_der, "dWOR_dt_pos"], linestyle="--", label="|dWOR/dt|")
            ax.set_xscale("log"); ax.set_yscale("log")
            ax.set_xlabel("t_pos (дни)"); ax.set_ylabel("WOR, |dWOR/dt|")
            ax.grid(True, which="both", alpha=0.3); ax.legend(loc="best")
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
