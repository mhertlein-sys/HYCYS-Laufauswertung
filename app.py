import streamlit as st
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import datetime
import pwlf
from PIL import Image
import os

# 1. APP-LAYOUT & LOGO
st.set_page_config(page_title="HYCYS - Laufauswertung", layout="wide")

# Kopfzeile: Titel links, Logo rechts
col_title, col_logo = st.columns([4, 1])
with col_title:
    st.title("HYCYS - Laufauswertung")
with col_logo:
    # Wir prüfen verschiedene mögliche Dateinamen für dein Logo
    logo_file = "image_822d59.png" # Dein aktuell hochgeladenes Bild
    if os.path.exists(logo_file):
        st.image(logo_file, use_container_width=True)
    elif os.path.exists("logo.png"):
        st.image("logo.png", use_container_width=True)

st.markdown("---")

# 2. SEITENLEISTE
st.sidebar.header("Basisdaten")
test_type = st.sidebar.selectbox("Testprotokoll", ["HYCYS Triathlon BLUE", "HYCYS Running BLUE"])
start_v = st.sidebar.number_input("Startgeschwindigkeit (m/s)", value=2.9, step=0.1)
anzahl = st.sidebar.number_input("Anzahl Stufen", value=8, min_value=3, max_value=20)
ausbelastung = st.sidebar.checkbox("Test bis zur Ausbelastung", value=True)

st.sidebar.header("Protokoll-Setup")
stufendauer = st.sidebar.number_input("Stufendauer (Minuten)", value=5.0, step=0.5)
pausendauer = st.sidebar.number_input("Pausendauer (Sekunden)", value=30, step=5)

# Körperfett-Menü (Einklappbar)
with st.sidebar.expander("Körperfettmessung (Parizkova 10-Falten)"):
    sf_names = ["Wange", "Kinn", "Achselfalte vorn", "10. Rippe", "Bauch (Nabel)", "Spina illiaca", "Oberschenkel", "Rücken", "Triceps", "Wade"]
    df_sf = pd.DataFrame({"Falte": sf_names, "M1": [0.0]*10, "M2": [0.0]*10, "M3": [0.0]*10})
    edited_sf = st.data_editor(df_sf, hide_index=True, use_container_width=True)

st.sidebar.header("Daten-Upload")
uploaded_file = st.sidebar.file_uploader("Spirometrie-Datei hochladen (.xlsx oder .csv)", type=["xlsx", "csv"])

st.sidebar.header("Stufentest Laktat")
default_lac = [1.03, 1.09, 1.19, 1.42, 2.36, 2.32, 6.05, 9.48]
default_hr = [123, 137, 145, 154, 164, 170, 177, 183]
speeds = [start_v + (i * 0.3) for i in range(anzahl)]
lac_values = [default_lac[i] if i < len(default_lac) else 0.0 for i in range(anzahl)]
hr_values = [default_hr[i] if i < len(default_hr) else 100 for i in range(anzahl)]
df_input = pd.DataFrame({"v (m/s)": np.round(speeds, 2), "Laktat": lac_values, "HF": hr_values})
edited_df = st.sidebar.data_editor(df_input, disabled=["v (m/s)"], hide_index=True, use_container_width=True)

st.sidebar.markdown("---")
start_button = st.sidebar.button("Auswertung starten")

# --- HILFSFUNKTIONEN ---
def parse_time_to_seconds(t):
    if pd.isnull(t): return None
    if isinstance(t, datetime.time): return t.hour * 3600 + t.minute * 60 + t.second
    if isinstance(t, str):
        try:
            parts = t.strip().split(':')
            if len(parts) == 3: return int(parts[0])*3600 + int(parts[1])*60 + float(parts[2])
            if len(parts) == 2: return int(parts[0])*60 + float(parts[1])
        except: pass
    if isinstance(t, (int, float)):
        if 0 < t < 1: return t * 86400 
        return float(t)
    return 0.0

def calculate_vlamax_for_v(v_thresh, rel_vo2max, speed_arr, vo2_steady_arr):
    if not (rel_vo2max > 0 and v_thresh > 0 and len(vo2_steady_arr) > 0): return None
    re_m_s_arr = [vo2 / v if (v > 0 and pd.notna(vo2) and vo2 > 0) else 0.0 for v, vo2 in zip(speed_arr, vo2_steady_arr)]
    re_at_thresh = np.interp(v_thresh, speed_arr, re_m_s_arr)
    rel_vo2_demand = re_at_thresh * v_thresh
    delta_vo2 = rel_vo2max - rel_vo2_demand
    if delta_vo2 <= 0 or rel_vo2_demand <= 0: return None
    ks1, ks2, la_eq, vol_dist = 0.0631, 1.3310, 0.02049, 0.4      
    vla_ox_max = (rel_vo2_demand * la_eq) / vol_dist
    adp = np.sqrt((ks1 * rel_vo2_demand) / delta_vo2)
    adp_3 = adp ** 3
    term2 = 1 + (ks2 / adp_3)
    return (vla_ox_max * term2) / 60

# 3. BERECHNUNG & ANZEIGE
if start_button:
    speed = edited_df["v (m/s)"].values
    lactate = edited_df["Laktat"].values
    hr = edited_df["HF"].values
    athlete_name, birthdate = "Unbekannt", "Unbekannt"
    weight, abs_vo2max, rel_vo2max, body_fat_pct = 0.0, 0.0, 0.0, 0.0
    vo2_steady_values = []
    spiro_df = pd.DataFrame()

    # Körperfett (Parizkova)
    sf_means = edited_sf[["M1", "M2", "M3"]].mean(axis=1)
    if sf_means.sum() > 0:
        body_fat_pct = (39.572 * np.log10(sf_means.sum())) - 61.25

    # Spiro-Daten einlesen
    if uploaded_file is not None:
        try:
            df_excel = pd.read_csv(uploaded_file, header=None, low_memory=False) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file, header=None)
            athlete_name = f"{df_excel.iloc[2, 1]} {df_excel.iloc[1, 1]}" 
            weight = float(df_excel.iloc[6, 1]) 
            raw_date = df_excel.iloc[7, 1] 
            if pd.notna(raw_date):
                date_str = str(int(float(raw_date))).zfill(8)
                birthdate = f"{date_str[:2]}.{date_str[2:4]}.{date_str[4:]}"
            
            times_sec = df_excel.iloc[3:, 9].apply(parse_time_to_seconds)
            vo2_col = pd.to_numeric(df_excel.iloc[3:, 13], errors='coerce')
            vco2_col = pd.to_numeric(df_excel.iloc[3:, 14], errors='coerce')
            spiro_df = pd.DataFrame({'Time': times_sec, 'VO2': vo2_col, 'VCO2': vco2_col}).dropna(subset=['Time'])
            
            if not spiro_df.empty:
                abs_vo2max = spiro_df['VO2'].rolling(window=3, center=True).mean().max()
                rel_vo2max = abs_vo2max / weight if weight > 0 else 0
                stage_sec = int(stufendauer * 60)
                pause_sec = int(pausendauer)
                for i in range(anzahl):
                    end_t = 60 + (i + 1) * stage_sec + i * pause_sec
                    mask = (spiro_df['Time'] >= end_t - 60) & (spiro_df['Time'] <= end_t)
                    vo2_steady_values.append(spiro_df.loc[mask, 'VO2'].mean() / weight if not spiro_df.loc[mask].empty else 0)
        except Exception as e:
            st.error(f"Fehler beim Einlesen: {e}")

    # Schwellenberechnung
    poly_coeffs = np.polyfit(speed, lactate, 3)
    poly_func = np.poly1d(poly_coeffs)
    v_start, v_end = speed[0], speed[-1]
    lac_start, lac_end = poly_func(v_start), poly_func(v_end)
    results = []

    def get_speed_at_lac(target):
        s_coeffs = poly_coeffs.copy(); s_coeffs[-1] -= target
        roots = [r.real for r in np.roots(s_coeffs) if np.isreal(r) and v_start <= r.real <= v_end]
        return roots[0] if roots else None

    v_40 = get_speed_at_lac(4.0)
    if v_40: results.append({'Modell': 'OBLA 4.0 (ANS)', 'v': v_40, 'Laktat': 4.0})
    
    deriv = np.polyder(poly_coeffs)
    m_dmax = (lac_end - lac_start) / (v_end - v_start)
    deriv_dmax = deriv.copy(); deriv_dmax[-1] -= m_dmax
    d_roots = [r.real for r in np.roots(deriv_dmax) if np.isreal(r) and (v_start-0.1) <= r.real <= (v_end+0.1)]
    if d_roots: results.append({'Modell': 'Dmax (Standard)', 'v': max(v_start, min(v_end, d_roots[0])), 'Laktat': poly_func(max(v_start, min(v_end, d_roots[0])))})

    th = np.min(lactate) + 0.4
    idx_a = np.where(lactate > th)[0]
    v_ms = speed[max(0, idx_a[0]-1)] if len(idx_a) > 0 else v_start
    m_m = (lac_end - poly_func(v_ms)) / (v_end - v_ms)
    deriv_mod = deriv.copy(); deriv_mod[-1] -= m_m
    m_roots = [r.real for r in np.roots(deriv_mod) if np.isreal(r) and (v_ms-0.1) <= r.real <= (v_end+0.1)]
    if m_roots: results.append({'Modell': 'Modified Dmax', 'v': max(v_ms, min(v_end, m_roots[0])), 'Laktat': poly_func(max(v_ms, min(v_end, m_roots[0])))})

    try:
        bps = pwlf.PiecewiseLinFit(speed, lactate).fit(3)
        for i, name in enumerate(['LTP1', 'LTP2']):
            if v_start <= bps[i+1] <= v_end: results.append({'Modell': name, 'v': bps[i+1], 'Laktat': poly_func(bps[i+1])})
    except: pass

    # Multi-VLamax
    for res in results:
        res['VLamax'] = calculate_vlamax_for_v(res['v'], rel_vo2max, speed, vo2_steady_values)

    # Header-Metriken
    st.subheader(f"Athlet: {athlete_name} | Geburtsdatum: {birthdate} | Protokoll: {test_type}")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Gewicht", f"{weight} kg")
    m2.metric("Körperfett", f"{max(0, body_fat_pct):.1f} %")
    m3.metric("VO2max (rel)", f"{rel_vo2max:.1f} ml/min/kg")
    m4.metric("Stufenanzahl", f"{anzahl}")
    st.markdown("---")

    # Ergebnisanzeige
    df_res = pd.DataFrame(results)
    if not df_res.empty:
        df_res['km/h'] = (df_res['v'] * 3.6).round(1)
        df_res['HF'] = np.interp(df_res['v'], speed, hr).round(0).astype(int)
        df_res['Laktat'] = df_res['Laktat'].round(2)
        df_res['VLamax [mmol/l/s]'] = df_res['VLamax'].apply(lambda x: f"{x:.3f}" if pd.notna(x) else "N/A")
        
        c1, c2 = st.columns([1, 1])
        with c1:
            st.subheader("Schwellen & VLamax Matrix")
            st.dataframe(df_res[['Modell', 'km/h', 'Laktat', 'HF', 'VLamax [mmol/l/s]']], hide_index=True, use_container_width=True)
        with c2:
            fig, ax = plt.subplots(figsize=(8, 4))
            v_smooth = np.linspace(v_start, v_end, 200)
            ax.plot(speed, lactate, 'ko', label='Messwerte')
            ax.plot(v_smooth, poly_func(v_smooth), color='#00a1e0', label='Laktatkurve')
            for _, r in df_res.iterrows():
                ax.plot(r['v'], r['Laktat'], 'X', markersize=8, label=r['Modell'])
            ax.set_xlabel("Geschwindigkeit (m/s)")
            ax.set_ylabel("Laktat (mmol/L)")
            ax.legend(loc='upper left', bbox_to_anchor=(1, 1))
            st.pyplot(fig)

    # Stufenauswertung
    if not spiro_df.empty:
        st.markdown("---")
        st.subheader("Stufenauswertung (Steady State letzte Minute)")
        step_data = []
        rest_mask = (spiro_df['Time'] >= 0) & (spiro_df['Time'] <= 60)
        step_data.append({"Stufe": "Ruhe", "km/h": 0.0, "VO2 (ml/min)": spiro_df.loc[rest_mask, 'VO2'].mean(), "VCO2 (ml/min)": spiro_df.loc[rest_mask, 'VCO2'].mean(), "RE (ml/kg/km)": None})
        
        stage_sec = int(stufendauer * 60)
        pause_sec = int(pausendauer)
        for i in range(anzahl):
            end_t = 60 + (i + 1) * stage_sec + i * pause_sec
            mask = (spiro_df['Time'] >= end_t - 60) & (spiro_df['Time'] <= end_t)
            vo2 = spiro_df.loc[mask, 'VO2'].mean() if not spiro_df.loc[mask].empty else 0
            v_kmh = speed[i] * 3.6
            re = (vo2 / weight) / (v_kmh / 60) if (weight > 0 and v_kmh > 0 and vo2 > 0) else None
            step_data.append({"Stufe": f"{i+1}", "km/h": v_kmh, "VO2 (ml/min)": vo2, "VCO2 (ml/min)": spiro_df.loc[mask, 'VCO2'].mean() if not spiro_df.loc[mask].empty else 0, "RE (ml/kg/km)": re})
        st.dataframe(pd.DataFrame(step_data).style.format({"km/h": "{:.1f}", "VO2 (ml/min)": "{:.0f}", "VCO2 (ml/min)": "{:.0f}", "RE (ml/kg/km)": "{:.1f}"}), use_container_width=True, hide_index=True)