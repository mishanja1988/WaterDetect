# -*- coding: utf-8 -*-
# Streamlit app: Автодиагностика скважин (Chan & Меркулова–Гинзбург) по (Скважина|Пласт)

from __future__ import annotations

import io, os, re, unicodedata
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
MIN_POINTS_DEFAULT = 6
SHARED_WATERCUT_THR_DEFAULT = 0.02
APP_TITLE = "Поскважинный автодиагноз (Chan & Меркулова–Гинзбург) — учёт Скважины и Пласта"

st.set_page_config(layout="wide", page_title="Автодиагностика скважин", page_icon="🛢️")

# ========================
# Описание + шаблоны
# ========================
st.markdown(f"""
# {APP_TITLE}

- Входной файл должен содержать **Скважина** и **Пласт**. Если чего-то нет — столбцы создаются автоматически.
- Ключ анализа: **`well_id = "Скважина | Пласт"`** — все расчёты, фильтры, графики и экспорт ведутся по нему.
- Поддерживаются входы:  
  ① *Жидкость (сут)* + *Обводнённость, %*; ② *Qo (сут)* + *Qw (сут)*; ③ *Qo(мес)* + *Qw(мес)*; ④ *Qж(мес)* + *Обводнённость, %*.  
  Дни добычи (*Дни добычи*) — опционально (если нет, берём 1).
""")

def _template_liq_wc() -> pd.DataFrame:
    return pd.DataFrame({
        "Скважина": ["113","113","115","115"],
        "Пласт": ["Б6","Б6","Бш","Бш"],
        "Дата": pd.to_datetime(["2024-01-01","2024-02-01","2024-01-01","2024-02-01"]),
        "Дни добычи": [31,29,31,29],
        "Жидкость м3/сут": [100,120,95,110],
        "Обводненность %": [10,22,5,18],
    })

def _template_qo_qw() -> pd.DataFrame:
    return pd.DataFrame({
        "Скважина": ["A-01","A-01","A-02","A-02"],
        "Пласт": ["Ю1","Ю1","Ю2","Ю2"],
        "Дата": pd.to_datetime(["2024-03-01","2024-04-01","2024-03-01","2024-04-01"]),
        "Дни добычи": [31,30,31,30],
        "Дебит нефти, м3/сут": [80,70,60,50],
        "Дебит воды, м3/сут": [20,40,10,25],
    })

def _bytes_xlsx(df: pd.DataFrame) -> bytes:
    bio = io.BytesIO(); df.to_excel(bio, index=False, engine="openpyxl"); bio.seek(0); return bio.getvalue()

def _bytes_csv(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")

st.markdown("### 📥 Шаблоны")
c1,c2,c3 = st.columns(3)
c1.download_button("🔽 Шаблон (Жидкость+Обводнённость) — XLSX", data=_bytes_xlsx(_template_liq_wc()), file_name="template_liq_wc.xlsx")
c2.download_button("🔽 Шаблон (Жидкость+Обводнённость) — CSV",  data=_bytes_csv(_template_liq_wc()),   file_name="template_liq_wc.csv")
c3.download_button("🔽 Шаблон (Qo/Qw) — XLSX", data=_bytes_xlsx(_template_qo_qw()), file_name="template_qo_qw.xlsx")

st.markdown("---")

# ========================
# Утилиты распознавания
# ========================
def _drop_unnamed(df: pd.DataFrame) -> pd.DataFrame:
    return df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]

def _norm_text(s: str) -> str:
    s = str(s)
    s = s.replace("ё", "е").replace("Ё","Е")
    s = unicodedata.normalize("NFKC", s)
    s = s.lower()
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"[^\w]+", "", s)  # убрать пунктуацию
    return s

def _norm_cols(df: pd.DataFrame) -> Dict[str,str]:
    # map normalized -> original
    return {_norm_text(c): c for c in df.columns}

def _find_col(normmap: Dict[str,str], variants: List[str]) -> Optional[str]:
    for v in variants:
        if v in normmap: return normmap[v]
    # частичное совпадение
    for key, orig in normmap.items():
        if any(v in key for v in variants):
            return orig
    return None

def _to_num_series(s: pd.Series, fill=None) -> pd.Series:
    # поддержка десятичной запятой и пробелов
    txt = s.astype(str).str.replace("\u00A0"," ").str.replace(" ","", regex=False).str.replace(",",".", regex=False)
    out = pd.to_numeric(txt, errors="coerce")
    return out.fillna(fill) if fill is not None else out

def _bytes_of_upload(f) -> bytes:
    f.seek(0); b=f.read(); f.seek(0); return b

# ========================
# Чтение файла
# ========================
@st.cache_data
def read_user_file(file_bytes: bytes, filename: str) -> pd.DataFrame:
    name = (filename or "").lower()
    if name.endswith((".xlsx",".xls")):
        df = pd.read_excel(io.BytesIO(file_bytes))
        return _drop_unnamed(df)
    last_exc=None
    for enc in ("utf-8","utf-8-sig","cp1251","latin1"):
        try:
            df = pd.read_csv(io.BytesIO(file_bytes), sep=None, engine="python", encoding=enc, on_bad_lines="skip")
            return _drop_unnamed(df)
        except Exception as e:
            last_exc=e
    raise ValueError(f"Не удалось прочитать файл: {last_exc}")

# ========================
# Подготовка: Скважина/Пласт → well_id и численные поля
# ========================
def _ensure_well_layer(df: pd.DataFrame) -> pd.DataFrame:
    nm = _norm_cols(df)
    well_col  = _find_col(nm, ["скважина","скв","well","id","номер"])
    layer_col = _find_col(nm, ["пласт","пл","layer","horizon","formation"])
    out = df.copy()
    if well_col is None:
        out["Скважина"] = [f"WELL_{i+1}" for i in range(len(out))]
        well_col="Скважина"
    if layer_col is None:
        out["Пласт"] = "UNK"
        layer_col="Пласт"
    out["Скважина"] = out[well_col].astype(str).fillna("").str.strip()
    out["Пласт"]    = out[layer_col].astype(str).fillna("").replace({"": "UNK"}).str.strip()
    out["well_id"]  = (out["Скважина"] + " | " + out["Пласт"]).str.strip()
    return out

def enforce_monotonic(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby("well_id", sort=False)["t_num"]
    t = g.cummax()
    idx = df.groupby("well_id", sort=False).cumcount().to_numpy()
    out = df.copy(); out["t_num"] = t.to_numpy() + idx * EPS; return out

@st.cache_data
def prepare_data(file_bytes: bytes, filename: str) -> Tuple[pd.DataFrame, Dict[str,str]]:
    raw = read_user_file(file_bytes, filename)
    raw.columns = [str(c) for c in raw.columns]
    df = _ensure_well_layer(raw)
    nm = _norm_cols(df)

    # детектор колонок
    det: Dict[str,str] = {}

    # даты/накопленное время
    c_date = _find_col(nm, ["дата","месяц","period","monthyear"])
    c_tcum = _find_col(nm, ["накопленноевремиработы","накопленноевремени","накопленноевремяработы","taccum"])

    # дни
    c_days = _find_col(nm, ["днидобычи","числоднейдобычинефтисут","proddays","aj","сут"])

    # периодные объёмы
    c_qoP = _find_col(nm, ["добычанефтим3мес","qoperiod","qo_mo","qo_mm","qo_q"])
    c_qwP = _find_col(nm, ["добычаводым3мес","qwperiod","qw_mo","qw_mm","qw_q"])
    c_qL_P = _find_col(nm, ["добыкажидкостим3мес","жидкостьм3мес","qlperiod","ql_mo","qL_mo"])

    # суточные
    c_liq = _find_col(nm, ["жидкостьм3сут","дебитжидкостим3сут","liquid","ql"])
    c_wc  = _find_col(nm, ["обводненность","обводненностьпроцент","watercut","wc","ab"])
    c_qo  = _find_col(nm, ["дебитнефтим3сут","qo","oilrate"])
    c_qw  = _find_col(nm, ["дебитводым3сут","qw","waterrate"])

    if c_qoP and c_qwP:
        df["qo_period"] = _to_num_series(df[c_qoP], 0.0); det["qo_period"]=c_qoP
        df["qw_period"] = _to_num_series(df[c_qwP], 0.0); det["qw_period"]=c_qwP
        df["qL_period"] = df["qo_period"] + df["qw_period"]
    elif c_qL_P and c_wc:
        qL = _to_num_series(df[c_qL_P], 0.0); det["qL_period"]=c_qL_P
        wc = _to_num_series(df[c_wc], 0.0);   det["watercut"]=c_wc
        df["qo_period"] = qL * (100.0 - wc)/100.0
        df["qw_period"] = qL * wc/100.0
        df["qL_period"] = qL
    elif c_liq and c_wc:
        liq = _to_num_series(df[c_liq], 0.0); det["liq"]=c_liq
        wc  = _to_num_series(df[c_wc], 0.0);  det["watercut"]=c_wc
        df["qo"] = liq * (100.0 - wc)/100.0
        df["qw"] = liq * wc/100.0
        df["qL"] = liq
        days = _to_num_series(df[c_days], 1.0) if c_days else pd.Series(1.0, index=df.index)
        if c_days: det["prod_days"]=c_days
        df["prod_days"] = days
        df["qo_period"] = df["qo"] * days
        df["qw_period"] = df["qw"] * days
        df["qL_period"] = df["qL"] * days
    elif c_qo and c_qw:
        df["qo"] = _to_num_series(df[c_qo], 0.0); det["qo"]=c_qo
        df["qw"] = _to_num_series(df[c_qw], 0.0); det["qw"]=c_qw
        df["qL"] = df["qo"] + df["qw"]
        days = _to_num_series(df[c_days], 1.0) if c_days else pd.Series(1.0, index=df.index)
        if c_days: det["prod_days"]=c_days
        df["prod_days"] = days
        df["qo_period"] = df["qo"] * days
        df["qw_period"] = df["qw"] * days
        df["qL_period"] = df["qL"] * days
    else:
        # минимально жизнеспособный фолбэк
        days = _to_num_series(df[c_days], 1.0) if c_days else pd.Series(1.0, index=df.index)
        if c_days: det["prod_days"]=c_days
        df["prod_days"] = days
        df["qo_period"] = pd.Series(0.0, index=df.index)
        df["qw_period"] = pd.Series(0.0, index=df.index)
        df["qL_period"] = pd.Series(0.0, index=df.index)

    # время
    if c_tcum:
        df["t_num"] = _to_num_series(df[c_tcum], 0.0); det["t_num"]=c_tcum
    elif c_date:
        t = pd.to_datetime(df[c_date], errors="coerce"); det["date"]=c_date
        df["t_num"] = (t - t.groupby(df["well_id"]).transform("min")).dt.days.astype(float).fillna(0.0)
    else:
        df["t_num"] = df["prod_days"].groupby(df["well_id"]).cumsum(); det["t_num"]="cum(prod_days)"

    df = df.dropna(subset=["well_id","t_num"]).sort_values(["well_id","t_num"]).reset_index(drop=True)
    df = enforce_monotonic(df)
    return df, det

# ========================
# Отбор сущностей + причины
# ========================
@st.cache_data
def select_eligible_with_reasons(df: pd.DataFrame, min_points: int, watercut_thr: float) -> Tuple[Tuple[str,...], Dict[str,str]]:
    reasons={}
    if df.empty: return tuple(), reasons
    with np.errstate(divide="ignore", invalid="ignore"):
        fw = df["qw_period"] / df["qL_period"]
    cond = (df["qL_period"]>0) & (fw>watercut_thr) & (df["prod_days"]>0)

    elig=[]
    for wid,g in df.groupby("well_id", sort=False):
        ok=int(cond[g.index].sum())
        if ok>=min_points: elig.append(wid)
        else:
            parts=[]
            if (g["qL_period"]>0).sum()==0: parts.append("нет положительных объёмов жидкости")
            if (g["prod_days"]>0).sum()==0: parts.append("нет положительных дней добычи")
            parts.append(f"точек после fw>{watercut_thr:.2f}: {ok} (нужно ≥{min_points})")
            reasons[wid]="; ".join(parts)
    return tuple(elig), reasons

# ========================
# MG/Chan
# ========================
@dataclass
class MGFlags:
    y_early_mean: Optional[float]=None
    slope_first_third: Optional[float]=None
    waviness_std: Optional[float]=None
    possible_behind_casing: bool=False
    possible_channeling: bool=False
    possible_mixed_causes: bool=False

@st.cache_data
def compute_mg(df: pd.DataFrame, ids: Tuple[str,...], thr: float, minp:int) -> pd.DataFrame:
    if df.empty or not ids: return pd.DataFrame(columns=["well_id","MG_X","MG_Y"])
    d=df[df["well_id"].isin(ids)].copy()
    with np.errstate(divide="ignore", invalid="ignore"):
        d["fw"]=d["qw_period"]/d["qL_period"]
    d=d.replace([np.inf,-np.inf], np.nan)
    outs=[]
    for wid,g in d.groupby("well_id", sort=False):
        g=g.sort_values("t_num")
        idx=np.flatnonzero((g["fw"].to_numpy()>thr)&(g["qL_period"].to_numpy()>0)&(g["prod_days"].to_numpy()>0))
        if idx.size==0: continue
        g2=g.iloc[idx[0]:].copy()
        g2["Qo_cum"]=g2["qo_period"].cumsum(); g2["Qw_cum"]=g2["qw_period"].cumsum()
        g2["Qt_cum"]=g2["Qo_cum"]+g2["Qw_cum"]
        if len(g2)<minp or float(g2["Qt_cum"].iloc[-1])<=0: continue
        Qt=float(g2["Qt_cum"].iloc[-1])
        X=(g2["Qt_cum"]/Qt).to_numpy()
        X=np.maximum.accumulate(X)+np.arange(len(X))*EPS
        g2["MG_X"]=X
        with np.errstate(divide="ignore", invalid="ignore"):
            g2["MG_Y"]=g2["Qo_cum"]/g2["Qt_cum"]

        flags=MGFlags()
        early=g2["MG_X"]<=0.2
        if early.sum()>=3:
            flags.y_early_mean=float(np.nanmean(g2.loc[early,"MG_Y"]))
            flags.possible_behind_casing=(flags.y_early_mean is not None) and (flags.y_early_mean>=0.99)
        first=g2[g2["MG_X"]<=0.33]
        if len(first)>=3:
            try:
                k,_=np.polyfit(first["MG_X"], first["MG_Y"], 1)
                flags.slope_first_third=float(k)
                flags.possible_channeling=(k<-0.8)
            except np.linalg.LinAlgError: pass
        if len(g2)>=5:
            with np.errstate(invalid="ignore"):
                dy=np.gradient(g2["MG_Y"].to_numpy(), g2["MG_X"].to_numpy())
            flags.waviness_std=float(np.nanstd(dy))
            flags.possible_mixed_causes=flags.waviness_std>1.0
        for k,v in vars(flags).items():
            g2[f"MG_diag_{k}"]=v
        outs.append(g2)
    return pd.concat(outs,axis=0).reset_index(drop=True) if outs else pd.DataFrame(columns=["well_id","MG_X","MG_Y"])

@dataclass
class ChanFlags:
    slope_logWOR_logt: Optional[float]=None
    mean_derivative: Optional[float]=None
    std_derivative: Optional[float]=None
    possible_coning: bool=False
    possible_near_wellbore: bool=False
    possible_multilayer_channeling: bool=False

@st.cache_data
def compute_chan(df: pd.DataFrame, ids: Tuple[str,...], minp:int) -> pd.DataFrame:
    if df.empty or not ids: return pd.DataFrame(columns=["well_id","t_pos","WOR","dWOR_dt","dWOR_dt_pos"])
    d=df[df["well_id"].isin(ids)].copy()
    outs=[]
    for wid,g in d.groupby("well_id", sort=False):
        g=g.sort_values("t_num").copy()
        with np.errstate(divide="ignore", invalid="ignore"):
            g["WOR"]=g["qw"]/g["qo"]
        g=g.replace([np.inf,-np.inf], np.nan)
        g=g[(g["qo"]>0)&(g["WOR"]>0)].dropna(subset=["WOR"])
        if len(g)<minp: continue
        g["t_pos"]=g["t_num"]-g["t_num"].min()+EPS
        with np.errstate(invalid="ignore"):
            g["dWOR_dt"]=np.gradient(g["WOR"].to_numpy(), g["t_pos"].to_numpy())
        g["dWOR_dt_pos"]=np.where(g["dWOR_dt"]>0, g["dWOR_dt"], np.nan)
        mask=(g["WOR"]>0)&(g["t_pos"]>0)
        a=np.nan
        if mask.sum()>=3:
            x=np.log(g.loc[mask,"t_pos"].to_numpy()); y=np.log(g.loc[mask,"WOR"].to_numpy())
            try: a,_=np.polyfit(x,y,1)
            except np.linalg.LinAlgError: pass
        flags=ChanFlags()
        flags.slope_logWOR_logt=float(a)
        flags.mean_derivative=float(np.nanmean(g["dWOR_dt"]))
        flags.std_derivative=float(np.nanstd(g["dWOR_dt"]))
        if not np.isnan(a):
            flags.possible_coning=(a>0.5 and flags.mean_derivative>0)
            flags.possible_near_wellbore=(a>1.0 and flags.mean_derivative>0)
            flags.possible_multilayer_channeling=(a>0 and flags.std_derivative>0.1)
        for k,v in vars(flags).items():
            g[f"chan_diag_{k}"]=v
        outs.append(g)
    return pd.concat(outs,axis=0).reset_index(drop=True) if outs else pd.DataFrame(columns=["well_id","t_pos","WOR","dWOR_dt","dWOR_dt_pos"])

# ========================
# Диагнозы
# ========================
def diag_mg(g: pd.DataFrame)->Dict[str,str]:
    if g.empty: return {"mg_text":"нет данных MG","mg_detail":""}
    r=g.iloc[-1]; parts=[]
    if r.get("MG_diag_possible_behind_casing"): parts.append("возможны заколонные перетоки (ранний нефтеотбор Y≈1)")
    if r.get("MG_diag_possible_channeling"):    parts.append("признаки каналирования (крутой спад Y в первой трети)")
    if r.get("MG_diag_possible_mixed_causes"):  parts.append("смешанные причины (высокая волнистость dY/dX)")
    if not parts: parts.append("ближе к равномерному обводнению")
    return {"mg_text":"; ".join(parts),
            "mg_detail":f"MG: y_early≈{r.get('MG_diag_y_early_mean',np.nan):.2f}; "
                        f"k≈{r.get('MG_diag_slope_first_third',np.nan):.2f}; "
                        f"std(dY/dX)≈{r.get('MG_diag_waviness_std',np.nan):.2f}"}

def diag_chan(g: pd.DataFrame)->Dict[str,str]:
    if g.empty: return {"chan_text":"нет данных Chan","chan_detail":""}
    r=g.iloc[-1]; parts=[]
    if r.get("chan_diag_possible_multilayer_channeling"): parts.append("многослойное каналирование (рост WOR и дисперсии)")
    if r.get("chan_diag_possible_near_wellbore"):         parts.append("приствольные проблемы/ранний канал")
    if r.get("chan_diag_possible_coning"):                parts.append("возможен конинг (наклон > 0.5)")
    if not parts: parts.append("нет выраженных признаков проблемного притока воды")
    return {"chan_text":"; ".join(parts),
            "chan_detail":f"Chan: slope≈{r.get('chan_diag_slope_logWOR_logt',np.nan):.2f}; "
                          f"mean dWOR/dt≈{r.get('chan_diag_mean_derivative',np.nan):.2e}; "
                          f"std≈{r.get('chan_diag_std_derivative',np.nan):.2e}"}

# ========================
# Экспорт
# ========================
@st.cache_data
def export_xlsx(mg_df, chan_df, diagnosis_df)->bytes:
    if mg_df is None or mg_df.empty: mg_df=pd.DataFrame(columns=["well_id"])
    if chan_df is None or chan_df.empty: chan_df=pd.DataFrame(columns=["well_id"])
    if diagnosis_df is None or diagnosis_df.empty:
        diagnosis_df=pd.DataFrame(columns=["well_id","mg_text","mg_detail","chan_text","chan_detail","reason"])
    bio=io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as w:
        diagnosis_df.to_excel(w, sheet_name="Summary", index=False)
        mg_df.to_excel(w, sheet_name="MG", index=False)
        chan_df.to_excel(w, sheet_name="Chan", index=False)
    bio.seek(0); return bio.getvalue()

# ========================
# UI
# ========================
with st.sidebar:
    st.subheader("Параметры допуска")
    water_thr = st.number_input("Порог обводнённости fw", 0.0, 1.0, SHARED_WATERCUT_THR_DEFAULT, 0.01)
    min_pts   = st.number_input("Мин. число точек после порога", 3, 200, MIN_POINTS_DEFAULT, 1)
    max_plot  = st.slider("Сколько сущностей рисовать", 1, 50, 10)

uploaded = st.file_uploader("Загрузите файл (.xlsx / .xls / .csv)", type=["xlsx","xls","csv"])

if not uploaded:
    st.info("Загрузите файл и нажмите «Запустить расчёт».")
else:
    b=_bytes_of_upload(uploaded)
    if st.button("▶ Запустить расчёт"):
        with st.spinner("Подготовка данных..."):
            df, detector = prepare_data(b, uploaded.name)

        # Диагностика распознанных полей
        with st.expander("🔍 Диагностика распознанных колонок"):
            st.json(detector)

        eligible_ids, reasons = select_eligible_with_reasons(df, min_pts, water_thr)
        if len(eligible_ids)==0 and "well_id" in df.columns and not df.empty:
            eligible_ids = tuple(pd.Index(df["well_id"].unique()).astype(str))

        prog = st.progress(0, text="Расчёт MG…")
        mg_df   = compute_mg(df, eligible_ids, water_thr, min_pts); prog.progress(50, text="Расчёт Chan…")
        chan_df = compute_chan(df, eligible_ids, min_pts);          prog.progress(100, text="Готово")

        rows=[]
        for wid in eligible_ids:
            mg_g = mg_df[mg_df["well_id"]==wid] if ("well_id" in mg_df.columns) else pd.DataFrame()
            ch_g = chan_df[chan_df["well_id"]==wid] if ("well_id" in chan_df.columns) else pd.DataFrame()
            row={"well_id":wid, **diag_mg(mg_g), **diag_chan(ch_g)}
            if wid in reasons: row["reason"]=reasons[wid]
            rows.append(row)
        summary = pd.DataFrame(rows)

        st.success(f"Готово. Сущностей (Скважина|Пласт): {len(eligible_ids)}")
        st.subheader("Сводная таблица диагнозов")
        st.dataframe(summary, use_container_width=True)

        st.subheader("Графики (ограниченный список)")
        show_ids = eligible_ids[:max_plot]
        cols = st.columns(2)
        for i,wid in enumerate(show_ids):
            with st.expander(f"{wid} — графики"):
                mg_g = mg_df[mg_df["well_id"]==wid] if ("well_id" in mg_df.columns) else pd.DataFrame()
                ch_g = chan_df[chan_df["well_id"]==wid] if ("well_id" in chan_df.columns) else pd.DataFrame()
                with cols[i%2]:
                    if not mg_g.empty:
                        fig,ax=plt.subplots(); ax.scatter(mg_g["MG_X"], mg_g["MG_Y"], s=10)
                        ax.grid(True, alpha=0.3); ax.set_xlabel("X = Qt_cum/Qt_cum(T)"); ax.set_ylabel("Y = Qo_cum/Qt_cum")
                        ax.set_title(f"MG — {wid}"); st.pyplot(fig)
                    else: st.info("Нет данных MG")
                    if not ch_g.empty:
                        fig2,ax2=plt.subplots()
                        ax2.plot(ch_g["t_pos"], ch_g["WOR"], "o", markersize=3, label="WOR")
                        ax2.plot(ch_g["t_pos"], ch_g["dWOR_dt_pos"], "--", label="|dWOR/dt|")
                        ax2.set_xscale("log"); ax2.set_yscale("log"); ax2.grid(True, which="both", alpha=0.3); ax2.legend()
                        ax2.set_xlabel("t_pos (дни)"); ax2.set_ylabel("WOR, |dWOR/dt|"); ax2.set_title(f"Chan — {wid}")
                        st.pyplot(fig2)
                    else: st.info("Нет данных Chan")

        st.subheader("Скачать результаты")
        st.download_button("📥 Единый Excel (Summary, MG, Chan)", data=export_xlsx(mg_df, chan_df, summary),
                           file_name="Autodiagnostics_results_by_layer.xlsx")
    else:
        st.info("Интерфейс готов. Нажмите «Запустить расчёт».")
