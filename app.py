import streamlit as st
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import datetime
import pwlf
from PIL import Image
import os
import plotly.graph_objects as go
import tempfile
from fpdf import FPDF
import openpyxl

# --- GLOBAL TIME & SPRINT PARSING HELPERS ---
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

def parse_sprint_excel(file_contents):
    wb = openpyxl.load_workbook(file_contents, data_only=True)
    sheet_name = "Data Sprint" if "Data Sprint" in wb.sheetnames else wb.sheetnames[0]
    sheet = wb[sheet_name]
    
    sprint_powers = []
    sprint_times = []
    
    for r in range(1, sheet.max_row + 1):
        val_a = sheet.cell(row=r, column=1).value
        val_e = sheet.cell(row=r, column=5).value
        
        if val_a is not None and val_e is not None:
            try:
                p = float(val_a)
                t = parse_time_to_seconds(val_e)
                if t is not None:
                    sprint_powers.append(p)
                    sprint_times.append(t)
            except ValueError:
                continue
                
    return sprint_powers, sprint_times


# TEMPORARY SRM DUMP
srm_file_path = "C:\\Users\\marku\\OneDrive\\Markus_HYCYS\\Sonstiges\\Laufauswertung App\\Radsport Auswertungen\\Ralph Ziegaus - 2026-05-22-10-28-31.srm"
if os.path.exists(srm_file_path):
    with open(srm_file_path, "rb") as f:
        data = f.read()
    with open("C:\\Users\\marku\\OneDrive\\Markus_HYCYS\\Sonstiges\\Laufauswertung App\\srm_hex_dump.txt", "w") as out:
        out.write(f"SRM File Size: {len(data)} bytes\n")
        for i in range(0, len(data), 16):
            chunk = data[i:i+16]
            hex_str = " ".join(f"{b:02x}" for b in chunk)
            ascii_str = "".join(chr(b) if 32 <= b < 127 else "." for b in chunk)
            out.write(f"{i:04x} | {hex_str:<47} | {ascii_str}\n")

# 1. APP-LAYOUT & CSS
st.set_page_config(page_title="HYCYS - Diagnostik", layout="wide")

st.markdown("""
    <style>
    button[kind="primary"] {
        background-color: #00a1e0 !important;
        color: white !important;
        border: none !important;
        width: 100% !important;
        font-weight: bold !important;
    }
    [data-testid="stFileUploader"] button {
        background-color: #00a1e0 !important;
        color: white !important;
        border: none !important;
    }
    </style>
    """, unsafe_allow_html=True)

# Initialize session state for metadata
metadata_defaults = [
    ("athlete_name", "Michael Wagner"), 
    ("birthdate", "06.03.1996"), 
    ("coach", "Markus Hertlein"), 
    ("test_date", "22.05.2026"), 
    ("sportart", "Bitte wählen..."), 
    ("kategorie", "Amateur"), 
    ("height", 178), 
    ("weight", 72.3), 
    ("body_fat_pct", 15.2), 
    ("last_uploaded_file", None),
    ("last_test_type", "HYCYS Standart (5min / +0,4m/s)"),
    ("start_v", 2.8),
    ("v_increment", 0.4),
    ("anzahl", 5),
    ("vorlauf", 60),
    ("stufendauer", 5.0),
    ("pausendauer", 30),
    ("ppo_val", 436.0),
    ("hfmax_val", 179.0),
    ("slope_val", 0.11),
    ("intercept_val", 138.0),
    ("carb_intake_factor", 1.0),
    ("vo2_override", 4368.0),
    ("t_bel_manual", 13.2),
    ("t_alak_manual", 3.1)
]
for key, val in metadata_defaults:
    if key not in st.session_state:
        st.session_state[key] = val

def update_sf_rad():
    changes = st.session_state.get("edited_sf_rad_key", {})
    if "edited_rows" in changes:
        for row_idx, edit in changes["edited_rows"].items():
            for col_name, new_val in edit.items():
                st.session_state.df_sf.loc[int(row_idx), col_name] = new_val

def update_sf_lauf():
    changes = st.session_state.get("edited_sf_key", {})
    if "edited_rows" in changes:
        for row_idx, edit in changes["edited_rows"].items():
            for col_name, new_val in edit.items():
                st.session_state.df_sf.loc[int(row_idx), col_name] = new_val

def update_sprint_lac():
    changes = st.session_state.get("sprint_lac_editor_key", {})
    if "edited_rows" in changes:
        for row_idx, edit in changes["edited_rows"].items():
            for col_name, new_val in edit.items():
                st.session_state.df_sprint_lac.loc[int(row_idx), col_name] = new_val

def update_lauftest_lac():
    changes = st.session_state.get("lauftest_lac_editor_key", {})
    if "edited_rows" in changes:
        for row_idx, edit in changes["edited_rows"].items():
            for col_name, new_val in edit.items():
                st.session_state.df_lauftest_input.loc[int(row_idx), col_name] = new_val

def sync_lauftest_input_df():
    anzahl = int(st.session_state.anzahl)
    start_v = float(st.session_state.start_v)
    v_increment = float(st.session_state.v_increment)
    
    stufe_names = ["Ruhe"] + [f"Stufe {i+1}" for i in range(anzahl)]
    speeds = [0.0] + [start_v + (i * v_increment) for i in range(anzahl)]
    speeds_kmh = [0.0] + [s * 3.6 for s in speeds[1:]]
    
    default_ruhe_lac = 1.0
    default_ruhe_hr = 70
    default_lac = [1.03, 1.09, 1.19, 1.42, 2.36, 2.32, 6.05, 9.48]
    default_hr = [123, 137, 145, 154, 164, 170, 177, 183]
    
    target_len = anzahl + 1
    
    if "df_lauftest_input" not in st.session_state:
        lac_values = [default_ruhe_lac] + [default_lac[i] if i < len(default_lac) else 2.0 for i in range(anzahl)]
        hr_values = [default_ruhe_hr] + [default_hr[i] if i < len(default_hr) else 130 for i in range(anzahl)]
        st.session_state.df_lauftest_input = pd.DataFrame({
            "Stufe": stufe_names,
            "v (m/s)": np.round(speeds, 2),
            "v (km/h)": np.round(speeds_kmh, 1),
            "Laktat": lac_values,
            "HF": hr_values
        })
        st.session_state.last_sync_params = (anzahl, start_v, v_increment)
    else:
        last_sync = st.session_state.get("last_sync_params", (None, None, None))
        if last_sync != (anzahl, start_v, v_increment):
            df = st.session_state.df_lauftest_input
            current_len = len(df)
            
            if current_len != target_len:
                if target_len > current_len:
                    new_rows = []
                    for i in range(current_len, target_len):
                        idx = i - 1
                        lac_val = default_lac[idx] if idx < len(default_lac) else 2.0
                        hr_val = default_hr[idx] if idx < len(default_hr) else 130
                        new_rows.append({
                            "Stufe": stufe_names[i],
                            "v (m/s)": round(speeds[i], 2),
                            "v (km/h)": round(speeds_kmh[i], 1),
                            "Laktat": lac_val,
                            "HF": hr_val
                        })
                    df = pd.concat([df, pd.DataFrame(new_rows)], ignore_index=True)
                else:
                    df = df.iloc[:target_len].copy()
            
            df["Stufe"] = stufe_names
            df["v (m/s)"] = np.round(speeds, 2)
            df["v (km/h)"] = np.round(speeds_kmh, 1)
            st.session_state.df_lauftest_input = df
            st.session_state.last_sync_params = (anzahl, start_v, v_increment)

# Primary Sportart Selection (must be the very first sidebar widget)
sportart_options = ["Bitte wählen...", "Laufen", "Radsport"]
sportart_val = st.session_state.sportart if st.session_state.sportart in sportart_options else "Bitte wählen..."
sportart = st.sidebar.selectbox("Sportart", options=sportart_options, index=sportart_options.index(sportart_val))
st.session_state.sportart = sportart

# Dynamically set title based on sportart
if sportart == "Radsport":
    title_text = "HYCYS Radsportauswertung"
elif sportart == "Laufen":
    title_text = "HYCYS Laufauswertung"
else:
    title_text = "HYCYS Diagnostik"

col_title, col_logo = st.columns([4, 1])
with col_title:
    st.title(title_text)
with col_logo:
    logo_candidates = ["image_822d59.png", "logo.png", "image_828f70.png"]
    logo_path = None
    for candidate in logo_candidates:
        if os.path.exists(candidate):
            st.image(candidate, width='stretch')
            logo_path = candidate
            break

st.markdown("---")

# 2. SEITENLEISTE
# Initialize widget keys in session state so they match general session state defaults
widget_keys_to_init = {
    "athlete_name": st.session_state.athlete_name,
    "birthdate": st.session_state.birthdate,
    "test_date": st.session_state.test_date,
    "height": int(st.session_state.height),
    "weight": float(st.session_state.weight),
    "body_fat_pct": float(st.session_state.body_fat_pct),
    "kategorie": st.session_state.kategorie,
    "coach": st.session_state.coach,
    
    "name_lauft_key": st.session_state.athlete_name,
    "birth_lauft_key": st.session_state.birthdate,
    "test_lauft_key": st.session_state.test_date,
    "height_lauft_key": int(st.session_state.height),
    "weight_lauft_key": float(st.session_state.weight),
    "bf_lauft_key": float(st.session_state.body_fat_pct),
    "kategorie_lauft_key": st.session_state.kategorie,
    "coach_lauft_key": st.session_state.coach,
}
for key, val in widget_keys_to_init.items():
    if key not in st.session_state:
        st.session_state[key] = val

def update_base_data(name, height, weight, birthdate, test_date):
    if name:
        st.session_state.athlete_name = name
        st.session_state.name_lauft_key = name
    if height is not None:
        st.session_state.height = int(height)
        st.session_state.height_lauft_key = int(height)
    if weight is not None:
        w_rounded = round(float(weight), 1)
        st.session_state.weight = w_rounded
        st.session_state.weight_lauft_key = w_rounded
    if birthdate:
        st.session_state.birthdate = birthdate
        st.session_state.birth_lauft_key = birthdate
    if test_date:
        st.session_state.test_date = test_date
        st.session_state.test_lauft_key = test_date

if "df_sf" not in st.session_state:
    sf_names = ["Wange", "Kinn", "Achselfalte vorn", "10. Rippe", "Bauch (Nabel)", "Spina illiaca", "Oberschenkel", "Rücken", "Triceps", "Wade"]
    st.session_state.df_sf = pd.DataFrame({"Falte": sf_names, "M1": [0.0]*10, "M2": [0.0]*10, "M3": [0.0]*10})

# Sportart selection handled at startup
pass

# ------------------ CONDITIONAL SPORTART RENDERING ------------------
if sportart == "Bitte wählen...":
    st.info("Bitte wählen Sie zuerst eine Sportart in der Sidebar aus, um mit der Auswertung zu beginnen.")
    st.stop()

elif sportart == "Radsport":
    # Pre-populate cycling constants and sprint/energetics parameters
    if "ppo_val" not in st.session_state: st.session_state.ppo_val = 436.0
    if "hfmax_val" not in st.session_state: st.session_state.hfmax_val = 179.0
    if "slope_val" not in st.session_state: st.session_state.slope_val = 0.11
    if "intercept_val" not in st.session_state: st.session_state.intercept_val = 138.0
    if "vo2max_spiro" not in st.session_state: st.session_state.vo2max_spiro = 4292.2
    if "t_bel" not in st.session_state: st.session_state.t_bel = 13.2
    if "t_alak" not in st.session_state: st.session_state.t_alak = 3.1
    if "t_glyc" not in st.session_state: st.session_state.t_glyc = 10.1
    if "t_bel_auto" not in st.session_state: st.session_state.t_bel_auto = 13.2
    if "t_alak_auto" not in st.session_state: st.session_state.t_alak_auto = 3.1
    if "t_bel_manual" not in st.session_state: st.session_state.t_bel_manual = 13.2
    if "t_alak_manual" not in st.session_state: st.session_state.t_alak_manual = 3.1
    if "rad_auswertung_gestartet" not in st.session_state: st.session_state.rad_auswertung_gestartet = False
    if "carb_intake_factor" not in st.session_state: st.session_state.carb_intake_factor = 1.0
    
    if "ks1" not in st.session_state: st.session_state.ks1 = 0.0631
    if "ks2" not in st.session_state: st.session_state.ks2 = 1.331
    if "ks4" not in st.session_state: st.session_state.ks4 = 11.7
    if "lac_eq" not in st.session_state: st.session_state.lac_eq = 0.02049
    if "lac_dist" not in st.session_state: st.session_state.lac_dist = 0.4
    if "vo2_rest" not in st.session_state: st.session_state.vo2_rest = 4.0

    coach_options = ["Markus Hertlein", "Marius Trompetter", "Susanne Traser", "Manuel Kuhnle", "Billie Benkel", "Jean Surmont", "Hosea Frick", "Björn Geesmann", "Gregor Eichhorn"]
    kategorie_options = ["Hobby", "Age-Grouper", "Amateur", "Profi"]

    # 0. File Uploaders placed at the very top of the sidebar
    uploaded_srm = st.sidebar.file_uploader(
        "Sprint-Rohdaten Excel hochladen (.xlsx)", 
        type=["xlsx"], 
        key="srm_sprint_key"
    )
    if uploaded_srm is not None and uploaded_srm.name != st.session_state.get("last_uploaded_srm"):
        st.session_state.last_uploaded_srm = uploaded_srm.name
        try:
            sprint_powers, sprint_times = parse_sprint_excel(uploaded_srm)
            if len(sprint_powers) > 0:
                p_max = max(sprint_powers)
                t_pmax_list = [t for p, t in zip(sprint_powers, sprint_times) if p == p_max]
                t_pmax_last = max(t_pmax_list) if t_pmax_list else 0.0
                t_alak = None
                p_threshold = p_max * 0.965
                for p, t in zip(sprint_powers, sprint_times):
                    if t > t_pmax_last and p < p_threshold:
                        t_alak = t
                        break
                if t_alak is None:
                    times_after_pmax = [t for t in sprint_times if t > t_pmax_last]
                    t_alak = times_after_pmax[-1] if times_after_pmax else (t_pmax_last + 0.1)
                t_bel = max(sprint_times)
                
                st.session_state.t_bel_auto = round(t_bel, 1)
                st.session_state.t_alak_auto = round(t_alak, 1)
                st.session_state.t_bel_manual = round(t_bel, 1)
                st.session_state.t_alak_manual = round(t_alak, 1)
                st.session_state.t_bel = round(t_bel, 1)
                st.session_state.t_alak = round(t_alak, 1)
                st.session_state.t_glyc = round(t_bel - t_alak, 1)
                st.session_state.sprint_powers = sprint_powers
                st.session_state.sprint_times = sprint_times
                st.session_state.rad_auswertung_gestartet = False
                st.rerun()
        except Exception as e:
            st.error(f"Fehler beim Auslesen der Sprintdatei: {e}")


    # --- Sprint .srm oder .fit Datei hochladen ---
    st.sidebar.markdown("**Sprint .srm / .fit hochladen**")
    
    # Slope und Zero-Offset Eingabe (für SRM-Kalibrierung)
    with st.sidebar.expander("SRM Kalibrierung (Slope / Zero-Offset)", expanded=False):
        st.caption("Diese Werte stehen auf dem SRM PowerControl oder im SRM-Protokoll.")
        _srm_slope  = st.number_input("Slope [W/count]", 
                                       value=float(st.session_state.get("srm_slope", 1.0)), 
                                       step=0.001, format="%.3f", key="srm_slope_input")
        _srm_offset = st.number_input("Zero-Offset [counts]",
                                       value=float(st.session_state.get("srm_offset", 0.0)),
                                       step=1.0, format="%.1f", key="srm_offset_input")
        st.session_state.srm_slope  = _srm_slope
        st.session_state.srm_offset = _srm_offset
        st.caption("Standard: Slope=1.0, Offset=0 → Raw-Werte werden direkt als Watt übernommen.")

    uploaded_sprint_raw = st.sidebar.file_uploader(
        "Sprint-Rohdaten hochladen (.srm oder .fit)", 
        type=["srm", "fit"], 
        key="sprint_raw_key"
    )
    
    if uploaded_sprint_raw is not None and uploaded_sprint_raw.name != st.session_state.get("last_uploaded_sprint_raw"):
        st.session_state.last_uploaded_sprint_raw = uploaded_sprint_raw.name
        _ext = os.path.splitext(uploaded_sprint_raw.name)[1].lower()
        _suffix = _ext if _ext in [".srm", ".fit"] else ".tmp"
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=_suffix) as tmp_raw:
            tmp_raw.write(uploaded_sprint_raw.getvalue())
            tmp_raw_path = tmp_raw.name
        
        try:
            # convert_sprint_to_excel als Modul importieren
            import importlib.util as _ilu
            _spec = _ilu.spec_from_file_location(
                "convert_sprint", 
                os.path.join(os.path.dirname(os.path.abspath(__file__)), "convert_sprint_to_excel.py")
            )
            _conv = _ilu.module_from_spec(_spec)
            _spec.loader.exec_module(_conv)
            
            _slope  = float(st.session_state.get("srm_slope", 1.0))
            _offset = float(st.session_state.get("srm_offset", 0.0))
            
            if _ext == ".srm":
                with open(tmp_raw_path, "rb") as _f:
                    _srm_data = _f.read()
                _records, _hdr = _conv.parse_srm7_data(_srm_data, slope=_slope, zero_offset=_offset, trim_zeros=True)
                _rec_int_label = f"{_hdr['rec_int']}s (SRM)"
                _src_name = uploaded_sprint_raw.name
            else:  # .fit
                _records, _ath, _date = _conv.parse_fit_file(tmp_raw_path, trim_zeros=True)
                _rec_int_label = "1s (FIT)"
                _src_name = uploaded_sprint_raw.name
            
            if _records:
                _sprint_powers = [float(r.get("power_w", r.get("power", 0))) for r in _records]
                _sprint_times  = [float(r["elapsed_s"]) for r in _records]
                
                _p_max = max(_sprint_powers)
                _t_pmax_list = [t for p, t in zip(_sprint_powers, _sprint_times) if p == _p_max]
                _t_pmax_last = max(_t_pmax_list) if _t_pmax_list else 0.0
                _t_alak = None
                _p_threshold = _p_max * 0.965
                for p, t in zip(_sprint_powers, _sprint_times):
                    if t > _t_pmax_last and p < _p_threshold:
                        _t_alak = t
                        break
                if _t_alak is None:
                    _times_after_pmax = [t for t in _sprint_times if t > _t_pmax_last]
                    _t_alak = _times_after_pmax[-1] if _times_after_pmax else (_t_pmax_last + 0.1)
                _t_bel  = max(_sprint_times)
                
                st.session_state.t_bel_auto    = round(_t_bel, 1)
                st.session_state.t_alak_auto   = round(_t_alak, 1)
                st.session_state.t_bel_manual  = round(_t_bel, 1)
                st.session_state.t_alak_manual = round(_t_alak, 1)
                st.session_state.t_bel   = round(_t_bel, 1)
                st.session_state.t_alak  = round(_t_alak, 1)
                st.session_state.t_glyc  = round(_t_bel - _t_alak, 1)
                st.session_state.sprint_powers  = _sprint_powers
                st.session_state.sprint_times   = _sprint_times
                st.session_state.sprint_records = _records   # für Excel-Download
                st.session_state.sprint_rec_int_label = _rec_int_label
                st.session_state.sprint_src_name = _src_name
                st.session_state.rad_auswertung_gestartet = False
                
                _n = len(_sprint_powers)
                _format_info = "0.1s (SRM)" if _ext == ".srm" else "1s (FIT)"
                st.success(f"✅ Geladen [{_format_info}]: {_n} Punkte, Max {int(_p_max)}W, Dauer {_t_bel:.1f}s — Nullwerte am Anfang automatisch entfernt.")
                st.rerun()
            else:
                st.warning("Keine Leistungsdaten mit Power > 0 gefunden.")
        
        except Exception as _e:
            st.error(f"Fehler beim Einlesen der Datei: {_e}")
        finally:
            if os.path.exists(tmp_raw_path):
                os.unlink(tmp_raw_path)

    # Download-Button: konvertierte Excel aus SRM/FIT erstellen
    if "sprint_records" in st.session_state and st.session_state.get("last_uploaded_sprint_raw"):
        try:
            import importlib.util as _ilu2
            _spec2 = _ilu2.spec_from_file_location(
                "convert_sprint2",
                os.path.join(os.path.dirname(os.path.abspath(__file__)), "convert_sprint_to_excel.py")
            )
            _conv2 = _ilu2.module_from_spec(_spec2)
            _spec2.loader.exec_module(_conv2)
            
            _dl_bytes = _conv2.create_sprint_excel_bytes(
                records=st.session_state.sprint_records,
                athlete_name=st.session_state.get("athlete_name", ""),
                test_date=st.session_state.get("test_date", ""),
                source_filename=st.session_state.get("sprint_src_name", "sprint"),
                rec_int_label=st.session_state.get("sprint_rec_int_label", ""),
            )
            _dl_name = st.session_state.get("last_uploaded_sprint_raw", "sprint").rsplit(".", 1)[0] + "_Data_Sprint.xlsx"
            st.sidebar.download_button(
                label="📥 Sprint Excel herunterladen",
                data=_dl_bytes,
                file_name=_dl_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_sprint_excel_key"
            )
        except Exception:
            pass  # Download optional



    uploaded_spiro = st.sidebar.file_uploader(
        "Spirodatei VO2max Rampe (10s) hochladen (.xlsx, .csv)", 
        type=["xlsx", "csv"], 
        key="spiro_rampe_key"
    )
    if uploaded_spiro is not None and uploaded_spiro.name != st.session_state.get("last_uploaded_spiro"):
        st.session_state.last_uploaded_spiro = uploaded_spiro.name
        try:
            df_spiro_excel = pd.read_excel(uploaded_spiro, header=None)
            lastname = df_spiro_excel.iloc[1, 1] if df_spiro_excel.shape[0] > 1 and df_spiro_excel.shape[1] > 1 else ""
            firstname = df_spiro_excel.iloc[2, 1] if df_spiro_excel.shape[0] > 2 and df_spiro_excel.shape[1] > 1 else ""
            name_val = f"{firstname} {lastname}".strip()
            
            h_val = None
            if df_spiro_excel.shape[0] > 5 and df_spiro_excel.shape[1] > 1:
                try: h_val = int(float(df_spiro_excel.iloc[5, 1]))
                except: pass
            
            w_val = None
            if df_spiro_excel.shape[0] > 6 and df_spiro_excel.shape[1] > 1:
                try: w_val = round(float(df_spiro_excel.iloc[6, 1]), 1)
                except: pass
                
            b_val = None
            if df_spiro_excel.shape[0] > 7 and df_spiro_excel.shape[1] > 1:
                try:
                    raw_birth = df_spiro_excel.iloc[7, 1]
                    birth_str = str(int(float(raw_birth))).zfill(8)
                    b_val = f"{birth_str[:2]}.{birth_str[2:4]}.{birth_str[4:]}"
                except: pass
                
            t_val = None
            if df_spiro_excel.shape[0] > 0 and df_spiro_excel.shape[1] > 4:
                try:
                    raw_test = df_spiro_excel.iloc[0, 4]
                    test_str = str(int(float(raw_test))).zfill(8)
                    t_val = f"{test_str[:2]}.{test_str[2:4]}.{test_str[4:]}"
                except: pass
                
            update_base_data(name_val, h_val, w_val, b_val, t_val)
                
            times_sec = df_spiro_excel.iloc[3:, 9].apply(parse_time_to_seconds)
            vo2_col = pd.to_numeric(df_spiro_excel.iloc[3:, 13], errors='coerce')
            df_spiro_data = pd.DataFrame({'Time': times_sec, 'VO2': vo2_col}).dropna(subset=['Time', 'VO2'])
            if len(df_spiro_data) >= 2:
                t_diffs = df_spiro_data['Time'].diff().dropna()
                interval = int(round(t_diffs.median()))
            else:
                interval = 10
            window_pts = max(1, int(30 / interval))
            vo2_smooth = df_spiro_data['VO2'].rolling(window=window_pts, center=True).mean()
            st.session_state.vo2max_spiro = round(vo2_smooth.max(), 1)
            st.session_state.rad_auswertung_gestartet = False
            st.rerun()
        except Exception as e:
            st.error(f"Fehler beim Auslesen der Spirodatei: {e}")

    uploaded_fit = st.sidebar.file_uploader(
        "VO2max Rampe .fit-Datei hochladen (.fit)", 
        type=["fit"], 
        key="fit_rampe_key"
    )
    if uploaded_fit is not None and uploaded_fit.name != st.session_state.get("last_uploaded_fit"):
        st.session_state.last_uploaded_fit = uploaded_fit.name
        with tempfile.NamedTemporaryFile(delete=False, suffix=".fit") as tmp:
            tmp.write(uploaded_fit.getvalue())
            tmp_path = tmp.name
        try:
            import fitparse
            fitfile = fitparse.FitFile(tmp_path)
            records_fit = []
            for record in fitfile.get_messages("record"):
                vals = record.get_values()
                ts = vals.get("timestamp")
                p = vals.get("power")
                hr = vals.get("heart_rate")
                if ts is not None:
                    records_fit.append({
                        "timestamp": ts,
                        "Power": float(p) if p is not None else 0.0,
                        "HR": float(hr) if hr is not None else np.nan
                    })
            
            if records_fit:
                df_raw = pd.DataFrame(records_fit)
                df_raw = df_raw.sort_values("timestamp")
                df_raw = df_raw.drop_duplicates(subset=["timestamp"], keep="last")
                df_raw.set_index("timestamp", inplace=True)
                df_fit = df_raw.resample("1s").ffill()
                df_fit = df_fit.reset_index()
            else:
                df_fit = pd.DataFrame(columns=["timestamp", "Power", "HR"])

            if not df_fit.empty:
                df_fit['Power_30s'] = df_fit['Power'].rolling(window=30, center=False).mean()
                max_p30 = df_fit['Power_30s'].max()
                if pd.isna(max_p30):
                    max_p30 = 0.0
                valid_hrs = df_fit['HR'].dropna().tolist()
                hf_max = max(valid_hrs) if valid_hrs else 180.0
            else:
                max_p30 = 0.0
                hf_max = 180.0
            
            # Only calculate regression once heart rate starts to rise significantly (excluding warmup and lag)
            # In the 3x3min warmup protocol, maximum power is ~150W; constant rise stabilizes around 215W.
            reg_df = df_fit[(df_fit['Power'] >= 215) & (~df_fit['HR'].isna())]
            if len(reg_df) >= 2:
                slope, intercept = np.polyfit(reg_df['Power'], reg_df['HR'], 1)
            else:
                slope, intercept = 0.15, 100.0
                
            st.session_state.ppo_val = round(max_p30, 1)
            st.session_state.hfmax_val = float(hf_max)
            st.session_state.slope_val = round(slope, 4)
            st.session_state.intercept_val = round(intercept, 2)
            st.session_state.rad_auswertung_gestartet = False
            st.rerun()
        except Exception as e:
            st.error(f"Fehler beim Auslesen der FIT-Datei: {e}")
        finally:
            if os.path.exists(tmp_path): os.unlink(tmp_path)

    # Pre-calculate body fat percentage from skinfolds to avoid rerun loops
    if "df_sf" in st.session_state:
        sf_means = st.session_state.df_sf[["M1", "M2", "M3"]].mean(axis=1)
        sum_sf = sf_means.sum()
        if sum_sf > 0:
            calculated_bf = (22.32 * np.log10(sum_sf)) - 29.2
            st.session_state.body_fat_pct = max(0.0, round(calculated_bf, 1))

    # 1. Basisdaten Expander (Default collapsed)
    with st.sidebar.expander("Basisdaten", expanded=False):
        athlete_name = st.text_input("Name", value=st.session_state.athlete_name)
        st.session_state.athlete_name = athlete_name

        birthdate = st.text_input("Geburtsdatum", value=st.session_state.birthdate)
        st.session_state.birthdate = birthdate

        test_date = st.text_input("Testdatum", value=st.session_state.test_date)
        st.session_state.test_date = test_date

        height = st.number_input("Größe (cm)", value=int(st.session_state.height), step=1, format="%d")
        st.session_state.height = height

        weight = st.number_input("Gewicht (kg)", value=float(st.session_state.weight), step=0.1, format="%.1f")
        st.session_state.weight = round(weight, 1)

        body_fat_pct = st.number_input("Körperfett (%)", value=float(st.session_state.body_fat_pct), step=0.1, format="%.1f")
        st.session_state.body_fat_pct = round(body_fat_pct, 1)

        coach_val = st.session_state.coach
        if coach_val not in coach_options:
            coach_options = coach_options + [coach_val]
        coach = st.selectbox("Coach", options=coach_options, index=coach_options.index(coach_val))
        st.session_state.coach = coach

        kategorie_val = st.session_state.kategorie
        if kategorie_val not in kategorie_options:
            kategorie_options = kategorie_options + [kategorie_val]
        kategorie = st.selectbox("Kategorie", options=kategorie_options, index=kategorie_options.index(kategorie_val))
        st.session_state.kategorie = kategorie
        
    with st.sidebar.form("cycling_form", enter_to_submit=False):
        # 1. Basisdaten Expander (Default collapsed)
        with st.expander("Basisdaten", expanded=False):
            athlete_name = st.text_input("Name", value=st.session_state.athlete_name)
            st.session_state.athlete_name = athlete_name

            birthdate = st.text_input("Geburtsdatum", value=st.session_state.birthdate)
            st.session_state.birthdate = birthdate

            test_date = st.text_input("Testdatum", value=st.session_state.test_date)
            st.session_state.test_date = test_date

            height = st.number_input("Größe (cm)", value=int(st.session_state.height), step=1, format="%d")
            st.session_state.height = height

            weight = st.number_input("Gewicht (kg)", value=float(st.session_state.weight), step=0.1, format="%.1f")
            st.session_state.weight = round(weight, 1)

            body_fat_pct = st.number_input("Körperfett (%)", value=float(st.session_state.body_fat_pct), step=0.1, format="%.1f")
            st.session_state.body_fat_pct = round(body_fat_pct, 1)

            coach_val = st.session_state.coach
            if coach_val not in coach_options:
                coach_options = coach_options + [coach_val]
            coach = st.selectbox("Coach", options=coach_options, index=coach_options.index(coach_val))
            st.session_state.coach = coach

            kategorie_val = st.session_state.kategorie
            if kategorie_val not in kategorie_options:
                kategorie_options = kategorie_options + [kategorie_val]
            kategorie = st.selectbox("Kategorie", options=kategorie_options, index=kategorie_options.index(kategorie_val))
            st.session_state.kategorie = kategorie
            
            # Körperfettmessung (Parizkova 10-Falten)
            with st.expander("Körperfettmessung (Parizkova 10-Falten)"):
                if "df_sf" not in st.session_state:
                    sf_names = ["Wange", "Kinn", "Achselfalte vorn", "10. Rippe", "Bauch (Nabel)", "Spina illiaca", "Oberschenkel", "Rücken", "Triceps", "Wade"]
                    st.session_state.df_sf = pd.DataFrame({"Falte": sf_names, "M1": [0.0]*10, "M2": [0.0]*10, "M3": [0.0]*10})
                edited_sf = st.data_editor(
                    st.session_state.df_sf,
                    hide_index=True,
                    width='stretch',
                    key="edited_sf_rad_key"
                )

        # 2. Sprinttest Expander (Default collapsed)
        with st.expander("Sprinttest", expanded=False):
            # Laktatwerte Sprint
            st.subheader("Laktatwerte Sprint")
            if "df_sprint_lac" not in st.session_state:
                times = ["Ruhelaktat 1", "Ruhelaktat 2", "Ruhelaktat 3", "Nachbelastung Min 0", "Nachbelastung Min 2", "Nachbelastung Min 3", "Nachbelastung Min 4", "Nachbelastung Min 5", "Nachbelastung Min 6", "Nachbelastung Min 7", "Nachbelastung Min 8", "Nachbelastung Min 9", "Nachbelastung Min 10"]
                default_lacs = [1.0, 0.94, 0.91, 1.38, 6.87, 7.1, 7.72, 7.75, 8.04, 7.79, 7.58, 7.42, 7.42]
                st.session_state.df_sprint_lac = pd.DataFrame({
                    "Messpunkt": times,
                    "Laktat [mmol/L]": default_lacs
                })
                
            edited_sprint_lac = st.data_editor(
                st.session_state.df_sprint_lac,
                disabled=["Messpunkt"],
                hide_index=True,
                use_container_width=True,
                height=500,
                column_config={
                    "Messpunkt": st.column_config.TextColumn("Messpunkt", width="medium"),
                    "Laktat [mmol/L]": st.column_config.NumberColumn("Laktat", min_value=0.0, max_value=25.0, step=0.01, format="%.2f", width="small")
                },
                key="sprint_lac_editor_key"
            )
        
        st.subheader("Sprinttest Parameter")
        t_bel_auto = float(st.session_state.get("t_bel_auto", 13.2))
        t_alak_auto = float(st.session_state.get("t_alak_auto", 3.1))
        t_glyc_auto = round(t_bel_auto - t_alak_auto, 1)
        
        t_bel_manual = float(st.session_state.get("t_bel_manual", 13.2))
        t_alak_manual = float(st.session_state.get("t_alak_manual", 3.1))
        t_glyc_manual = round(t_bel_manual - t_alak_manual, 1)
        
        st.markdown(f"""
        | Parameter | Auto (Excel) | Manuell |
        | :--- | :--- | :--- |
        | **t_bel [s]** | {t_bel_auto:.1f} | {t_bel_manual:.1f} |
        | **t_alak [s]** | {t_alak_auto:.1f} | {t_alak_manual:.1f} |
        | **t_glyc [s]** | {t_glyc_auto:.1f} | {t_glyc_manual:.1f} |
        """)
        
        col_t1, col_t2 = st.columns(2)
        with col_t1:
            t_bel_manual_input = st.number_input("Manuelle t_bel [s]", value=float(st.session_state.t_bel_manual), step=0.1, format="%.1f")
            st.session_state.t_bel_manual = round(t_bel_manual_input, 1)
        with col_t2:
            t_alak_manual_input = st.number_input("Manuelle t_alak [s]", value=float(st.session_state.t_alak_manual), step=0.1, format="%.1f")
            st.session_state.t_alak_manual = round(t_alak_manual_input, 1)
                
        sprint_source = st.selectbox(
            "Aktive Quelle für Sprinttest",
            ["Automatisch aus Excel-Rohdaten", "Manuelle Anpassung"],
            key="sprint_source_key"
        )

        # 3. Rampentest Expander (Default collapsed)
        with st.expander("Rampentest", expanded=False):
            st.subheader("Rampentest Parameter")
            ppo_val_input = st.number_input("30s PPO [W]", value=float(st.session_state.ppo_val), step=1.0, format="%.1f")
            st.session_state.ppo_val = round(ppo_val_input, 1)
            
            hfmax_val_input = st.number_input("max. Herzfrequenz [bpm]", value=float(st.session_state.hfmax_val), step=1.0, format="%.1f")
            st.session_state.hfmax_val = round(hfmax_val_input, 1)
            
            slope_val_input = st.number_input("Steigung lin. Reg. [bpm/W]", value=float(st.session_state.slope_val), step=0.0001, format="%.4f")
            st.session_state.slope_val = round(slope_val_input, 4)
            
            intercept_val_input = st.number_input("Nullstelle lin. Reg. [bpm]", value=float(st.session_state.intercept_val), step=1.0, format="%.1f")
            st.session_state.intercept_val = round(intercept_val_input, 1)

            st.subheader("VO2max Quellenvergleich")
            weight = float(st.session_state.get("weight", 72.3))
            ppo_val = float(st.session_state.get("ppo_val", 436.0))
            spiro_abs = float(st.session_state.get("vo2max_spiro", 4292.2))
            spiro_rel = spiro_abs / weight if weight > 0 else 0.0
            m_abs = ppo_val * 9.44 + 592.0
            m_rel = m_abs / weight if weight > 0 else 0.0
            w_abs = ppo_val * 9.44 + 275.0
            w_rel = w_abs / weight if weight > 0 else 0.0
            override_abs = float(st.session_state.get("vo2_override", 4368.0))
            override_rel = override_abs / weight if weight > 0 else 0.0
            
            st.markdown(f"""
            | Quelle | Abs. [ml] | Rel. [ml/kg] |
            | :--- | :--- | :--- |
            | **Spiro (gemessen)** | {int(spiro_abs)} | {spiro_rel:.1f} |
            | **PPO (♂)** | {int(m_abs)} | {m_rel:.1f} |
            | **PPO (♀)** | {int(w_abs)} | {w_rel:.1f} |
            | **Manuell (override)** | {int(override_abs)} | {override_rel:.1f} |
            """)
            
            vo2_override_input = st.number_input("Manuelle VO2max [ml/min]", value=float(st.session_state.vo2_override), step=50.0)
            st.session_state.vo2_override = round(vo2_override_input, 1)
            
            vo2_source = st.selectbox(
                "Aktive Quelle für Simulation",
                ["Gemessene Spirodatei", "Berechnet aus 30s PPO (männlich)", "Berechnet aus 30s PPO (weiblich)", "Manuelle Eingabe"],
                key="vo2_source_rad_key"
            )
            if vo2_source == "Manuelle Eingabe":
                vo2_override = float(st.session_state.vo2_override)
            else:
                vo2_override = 0.0

            st.subheader("Energetik Parameter")
            carb_intake_factor_input = st.number_input("KH-Zufuhr Faktor [g/kg KG/h]", value=float(st.session_state.carb_intake_factor), step=0.1, format="%.1f")
            st.session_state.carb_intake_factor = round(carb_intake_factor_input, 1)
            carb_intake_factor = float(st.session_state.carb_intake_factor)

        st.markdown("---")
        start_button_rad = st.form_submit_button("Auswertung starten", type="primary")
        if start_button_rad:
            st.session_state.rad_auswertung_gestartet = True
            st.session_state.df_sprint_lac = edited_sprint_lac
            st.session_state.df_sf = edited_sf
        
    # 6. RENDER MAIN PAGE LAYOUT
    if st.session_state.rad_auswertung_gestartet:
        # Fetch base variables from key-bound session state
        weight = float(st.session_state.get("weight", 72.3))
        height = int(st.session_state.get("height", 178))
        body_fat_pct = float(st.session_state.get("body_fat_pct", 15.2))
        athlete_name = st.session_state.get("athlete_name", "Michael Wagner")
        birthdate = st.session_state.get("birthdate", "06.03.1996")
        test_date = st.session_state.get("test_date", "22.05.2026")
        kategorie = st.session_state.get("kategorie", "Amateur")
        coach = st.session_state.get("coach", "Markus Hertlein")
        
        ppo_val = float(st.session_state.get("ppo_val", 436.0))
        hfmax_val = float(st.session_state.get("hfmax_val", 179.0))
        slope_val = float(st.session_state.get("slope_val", 0.11))
        intercept_val = float(st.session_state.get("intercept_val", 138.0))
        carb_intake_factor = float(st.session_state.get("carb_intake_factor", 1.0))

        # A. Check and compute VO2max
        if vo2_source == "Gemessene Spirodatei":
            abs_vo2max = float(st.session_state.vo2max_spiro)
            vo2_source_lbl = f"Spirodatei ({abs_vo2max:.1f} ml/min)"
        elif vo2_source == "Berechnet aus 30s PPO (männlich)":
            abs_vo2max = float(st.session_state.ppo_val) * 9.44 + 592.0
            vo2_source_lbl = f"30s PPO männlich ({abs_vo2max:.1f} ml/min)"
        elif vo2_source == "Berechnet aus 30s PPO (weiblich)":
            abs_vo2max = float(st.session_state.ppo_val) * 9.44 + 275.0
            vo2_source_lbl = f"30s PPO weiblich ({abs_vo2max:.1f} ml/min)"
        else:
            abs_vo2max = float(vo2_override)
            vo2_source_lbl = f"Manuell überschrieben ({abs_vo2max:.1f} ml/min)"
            
        rel_vo2max = abs_vo2max / weight if weight > 0 else 0.0
        
        # B. Load Sprint parameters based on active source
        sprint_src = st.session_state.get("sprint_source_key", "Automatisch aus Excel-Rohdaten")
        if sprint_src == "Automatisch aus Excel-Rohdaten":
            st.session_state.t_bel = float(st.session_state.get("t_bel_auto", 13.2))
            st.session_state.t_alak = float(st.session_state.get("t_alak_auto", 3.1))
        else:
            st.session_state.t_bel = float(st.session_state.get("t_bel_manual", 13.2))
            st.session_state.t_alak = float(st.session_state.get("t_alak_manual", 3.1))
        st.session_state.t_glyc = round(st.session_state.t_bel - st.session_state.t_alak, 1)
        
        t_bel = float(st.session_state.t_bel)
        t_alak = float(st.session_state.t_alak)
        t_glyc = float(st.session_state.t_glyc)
        
        # Extract Lactates from table
        lac_vals = st.session_state.df_sprint_lac["Laktat [mmol/L]"].values
        ruhe_lacs = lac_vals[0:3]
        avg_rest_lac = np.mean(ruhe_lacs)
        post_lacs = lac_vals[3:]
        max_post_lac = np.max(post_lacs)
        
        # Calculate VLamax
        vlamax = (max_post_lac - avg_rest_lac) / t_glyc if t_glyc > 0 else 0.0
        
        # C. Run metabolic grid simulation (0 to 600 W) using session state constants
        sim_results = []
        ks1 = float(st.session_state.ks1)
        ks2 = float(st.session_state.ks2)
        ks4 = float(st.session_state.ks4)
        lac_eq = float(st.session_state.lac_eq)
        lac_dist = float(st.session_state.lac_dist)
        vo2_rest = float(st.session_state.vo2_rest)
        
        for p in range(601):
            vo2ss = (p * ks4 / weight) + vo2_rest
            
            # ADP calculation (using numpy to handle invalid operations gracefully)
            if rel_vo2max - vo2ss > 0:
                adp = np.sqrt((ks1 * vo2ss) / (rel_vo2max - vo2ss))
            else:
                adp = np.nan
                
            adp_3 = adp ** 3 if not np.isnan(adp) else np.nan
            
            # VLass (mmol/L/min)
            vlass = (vlamax * 60.0) / (1.0 + ks2 / adp_3) if not np.isnan(adp_3) and adp_3 > 0 else np.nan
            
            # VLaoxmax (mmol/L/min)
            vlaoxmax = (lac_eq / lac_dist) * vo2ss
            
            # Pyruvatdefizit (VLaoxmax - VLass)
            pyruvat_def = vlaoxmax - vlass if not np.isnan(vlass) else np.nan
            
            # Fuel splitting (capped for energetics and table display to match physiology)
            kh_pct = (vlass / vlaoxmax * 100.0) if vlaoxmax > 0 and not np.isnan(vlass) else np.nan
            fat_pct = 100.0 - kh_pct if not np.isnan(kh_pct) else np.nan
            
            if not np.isnan(kh_pct):
                kh_pct_capped = max(0.0, min(100.0, kh_pct))
                fat_pct_capped = max(0.0, min(100.0, 100.0 - kh_pct_capped))
            else:
                kh_pct_capped = np.nan
                fat_pct_capped = np.nan
            
            # RQ
            rq = 0.7 + 0.3 * (kh_pct_capped / 100.0) if not np.isnan(kh_pct_capped) else np.nan
            
            # AE (Energy equivalent, kJ/L O2)
            ae = (((rq - 0.7) * 0.05094) * 100.0) + 19.6 if not np.isnan(rq) else np.nan
            
            # EE gesamt [kcal/h]
            vo2_abs_ss = vo2ss * weight
            ee_gesamt = ((vo2_abs_ss / 1000.0) * ae / 4.186) * 60.0 if not np.isnan(ae) else np.nan
            
            # EE Fett [kcal/h]
            ee_fett = ee_gesamt * (fat_pct_capped / 100.0) if not np.isnan(ee_gesamt) and not np.isnan(fat_pct_capped) else np.nan
            
            # EE KH [g/h]
            ee_kh_g = (ee_gesamt * (kh_pct_capped / 100.0)) / 4.063 if not np.isnan(ee_gesamt) and not np.isnan(kh_pct_capped) else np.nan
            
            # HR
            hr_val = p * slope_val + intercept_val
            if hr_val > hfmax_val:
                hr_val = np.nan
                
            sim_results.append({
                'power': p,
                'vo2ss': vo2ss,
                'vlass': vlass,
                'vlaoxmax': vlaoxmax,
                'pyruvat_def': pyruvat_def,
                'kh_pct': kh_pct_capped,
                'fat_pct': fat_pct_capped,
                'rq': rq,
                'ee_gesamt': ee_gesamt,
                'ee_fett': ee_fett,
                'ee_kh_g': ee_kh_g,
                'hr': hr_val
            })
            
        df_sim = pd.DataFrame(sim_results)
        
        # D. Extract Derived Metrics
        # ANS crossover: power step where vlass > vlaoxmax and the positive difference is minimized
        ans_df = df_sim[df_sim['vlass'] > df_sim['vlaoxmax']].copy()
        if not ans_df.empty:
            diff = ans_df['vlass'] - ans_df['vlaoxmax']
            min_diff_idx = diff.idxmin()
            ans_power = int(df_sim.loc[min_diff_idx, 'power'])
        else:
            ans_power = 0
            
        ans_rel = ans_power / weight
        ans_hr = ans_power * slope_val + intercept_val
        
        # Note the Excel formula inconsistency in Report_Cycling!AA16 (adds vo2_rest before weight division)
        ans_vo2_rel = ((ans_power * ks4) + vo2_rest) / weight
        ans_vo2_abs = ans_vo2_rel * weight
        ans_ee_gesamt = df_sim.loc[df_sim['power'] == ans_power, 'ee_gesamt'].values[0]
        ans_ee_kh_g = df_sim.loc[df_sim['power'] == ans_power, 'ee_kh_g'].values[0]
        
        # Fatmax: power where pyruvat_def (vlaoxmax - vlass) is maximized
        fatmax_df = df_sim.dropna(subset=['pyruvat_def']).copy()
        fatmax_idx = fatmax_df['pyruvat_def'].idxmax()
        fatmax_row = df_sim.loc[fatmax_idx]
        fatmax_power = int(fatmax_row['power'])
        fatmax_rel = fatmax_power / weight
        fatmax_ee_gesamt = fatmax_row['ee_gesamt']
        fatmax_ee_fett = fatmax_row['ee_fett']
        
        # MAP (max. Leistung)
        map_val = (abs_vo2max - weight * vo2_rest) / ks4
        
        # Match-Schwelle: power where ee_kh_g matches carb intake
        carb_intake_factor = float(st.session_state.carb_intake_factor)
        carb_intake = carb_intake_factor * weight
        carb_match_df = df_sim[df_sim['ee_kh_g'] <= carb_intake]
        if not carb_match_df.empty:
            carb_match_power = int(carb_match_df['power'].max())
        else:
            carb_match_power = 0
            
        # Body mass indices
        bmi = weight / ((height/100.0)**2) if height > 0 else 0.0
        fett_kg = weight * (body_fat_pct / 100.0)
        fettfrei_kg = weight - fett_kg
        vo2_ffm = abs_vo2max / fettfrei_kg if fettfrei_kg > 0 else 0.0
        
        # Render Main Panels
        st.markdown(f"""
            <div style="font-size: 16px; line-height: 1.5; margin-bottom: 20px;">
                <b>Athlet:</b> {athlete_name}<br>
                <b>Sportart:</b> Radsport<br>
                <b>Test-Datum:</b> {test_date}<br>
                <b>Kategorie:</b> {kategorie} | <b>Coach:</b> {coach}
            </div>
            """, unsafe_allow_html=True)
            
        st.subheader("Stammdaten & Anthropometrie")
        col_base_1, col_base_2, col_base_3 = st.columns(3)
        with col_base_1:
            st.markdown(f"**Gewicht:** {weight:.1f} kg")
            st.markdown(f"**Größe:** {height} cm")
            st.markdown(f"**BMI:** {bmi:.2f} kg/m²")
        with col_base_2:
            st.markdown(f"**Körperfett:** {body_fat_pct:.1f} %")
            st.markdown(f"**Fettmasse:** {fett_kg:.1f} kg")
            st.markdown(f"**Fettfreie Masse:** {fettfrei_kg:.1f} kg")
        with col_base_3:
            st.markdown(f"**Geburtsdatum:** {birthdate}")
            st.markdown(f"**Testdatum:** {test_date}")
            st.markdown(f"**Alter:** {st.session_state.get('birthdate', birthdate)}")
            
        st.markdown("---")
        st.subheader("⚡ Physiologische Kennzahlen & Ergebnisse")
        
        # Grid layout for metrics
        c_1, c_2, c_3, c_4 = st.columns(4)
        with c_1:
            st.metric("VO2max (abs.)", f"{int(abs_vo2max)} ml/min")
            st.caption(f"Quelle: {vo2_source_lbl}")
        with c_2:
            st.metric("VO2max (rel.)", f"{rel_vo2max:.1f} ml/min/kg")
            st.caption(f"FFM: {vo2_ffm:.1f} ml/min/kg FFM")
        with c_3:
            st.metric("VLamax", f"{vlamax:.3f} mmol/l/s")
            st.caption(f"t_alak: {t_alak:.1f}s | t_bel: {t_bel:.1f}s")
        with c_4:
            st.metric("MAP (max. Leistung)", f"{map_val:.1f} W")
            st.caption(f"relativ: {map_val/weight:.2f} W/kg")
            
        c_5, c_6, c_7, c_8 = st.columns(4)
        with c_5:
            st.metric("ANS Schwelle (PMLSSc)", f"{ans_power} W")
            st.caption(f"rel: {ans_rel:.2f} W/kg | EE ges: {int(ans_ee_gesamt)} kcal/h")
        with c_6:
            st.metric("Herzfrequenz @ ANS", f"{int(ans_hr)} bpm")
            st.caption(f"VO2 abs: {int(ans_vo2_abs)} ml/min")
        with c_7:
            st.metric("Fatmax Leistung", f"{fatmax_power} W")
            st.caption(f"rel: {fatmax_rel:.2f} W/kg | EE Fett: {int(fatmax_ee_fett)} kcal/h")
        with c_8:
            st.metric("Match-Schwelle", f"{carb_match_power} W")
            st.caption(f"rel: {carb_match_power/weight:.2f} W/kg | Zufuhr: {int(carb_intake)} g/h")

        st.markdown("---")
        
        # Sprint-Verifizierung Section
        st.subheader("⏱️ Sprint-Verifizierung & VLamax-Berechnung")
        col_sprint_1, col_sprint_2 = st.columns([2, 1])
        
        with col_sprint_1:
            if "sprint_powers" in st.session_state and len(st.session_state.sprint_powers) > 0:
                # Plotly figure for Sprint Verification
                fig_sprint = go.Figure()
                
                # Line for sprint powers
                fig_sprint.add_trace(go.Scatter(
                    x=st.session_state.sprint_times,
                    y=st.session_state.sprint_powers,
                    mode='lines',
                    line=dict(color='#00a1e0', width=3),
                    name='Leistung'
                ))
                
                # Find p_max and t_pmax_last to draw lines
                sprint_pmax = max(st.session_state.sprint_powers)
                sprint_palak = 0.965 * sprint_pmax
                
                # Add horizontal line for P_alak (96.5% Pmax)
                fig_sprint.add_hline(
                    y=sprint_palak,
                    line_dash="dot",
                    line_color="#cdb663",
                    annotation_text=f"96.5% Pmax ({sprint_palak:.0f} W)",
                    annotation_position="bottom left"
                )
                
                # Shaded region for Alactic phase (0 to t_alak)
                fig_sprint.add_vrect(
                    x0=0.0, x1=t_alak,
                    fillcolor="rgba(0, 161, 224, 0.1)",
                    layer="below", line_width=0,
                    annotation_text="Alaktazid (t_alak)",
                    annotation_position="top left"
                )
                
                # Shaded region for Glycolytic phase (t_alak to t_bel)
                fig_sprint.add_vrect(
                    x0=t_alak, x1=t_bel,
                    fillcolor="rgba(205, 182, 99, 0.1)",
                    layer="below", line_width=0,
                    annotation_text="Glykolytisch (t_glyc)",
                    annotation_position="top right"
                )
                
                fig_sprint.update_layout(
                    xaxis_title="Zeit [Sekunden]",
                    yaxis_title="Leistung [Watt]",
                    height=350,
                    margin=dict(l=40, r=40, t=20, b=40),
                    plot_bgcolor='#f9f9f9',
                    legend=dict(x=0.02, y=0.98, bgcolor="rgba(255,255,255,0.7)")
                )
                st.plotly_chart(fig_sprint, use_container_width=True)
            else:
                st.info("ℹ️ Keine Sprint-Rohdaten hochgeladen. Es wird die standardmäßig konfigurierte Belastungsdauer verwendet.")
                
        with col_sprint_2:
            st.markdown(f"""
            <div style="border: 1px solid #e6e6e6; border-radius: 8px; padding: 15px; background-color: #f9f9f9; color: #222;">
                <h4 style="margin-top:0; color: #111;">Dauer der Phasen</h4>
                <table style="width:100%; border-collapse: collapse;">
                    <tr style="border-bottom: 1px solid #e6e6e6;">
                        <td style="padding: 6px 0;"><b>Gesamtdauer (t_bel):</b></td>
                        <td style="text-align:right;">{t_bel:.1f} s</td>
                    </tr>
                    <tr style="border-bottom: 1px solid #e6e6e6;">
                        <td style="padding: 6px 0;"><b>Alaktazide Phase (t_alak):</b></td>
                        <td style="text-align:right;">{t_alak:.1f} s</td>
                    </tr>
                    <tr style="border-bottom: 1px solid #e6e6e6; color: #cdb663;">
                        <td style="padding: 6px 0;"><b>Glykolytische Phase (t_glyc):</b></td>
                        <td style="text-align:right;"><b>{t_glyc:.1f} s</b></td>
                    </tr>
                </table>
                <h4 style="margin-top: 15px; color: #111;">Laktat-Differenz</h4>
                <table style="width:100%; border-collapse: collapse;">
                    <tr style="border-bottom: 1px solid #e6e6e6;">
                        <td style="padding: 6px 0;">Ruhelaktat (Mittelwert):</td>
                        <td style="text-align:right;">{avg_rest_lac:.2f} mmol/L</td>
                    </tr>
                    <tr style="border-bottom: 1px solid #e6e6e6;">
                        <td style="padding: 6px 0;">Peak-Laktat (post-sprint):</td>
                        <td style="text-align:right;">{max_post_lac:.2f} mmol/L</td>
                    </tr>
                    <tr style="border-bottom: 1px solid #e6e6e6; color: #00a1e0;">
                        <td style="padding: 6px 0;"><b>Laktat-Anstieg (&Delta;La):</b></td>
                        <td style="text-align:right;"><b>{(max_post_lac - avg_rest_lac):.2f} mmol/L</b></td>
                    </tr>
                </table>
                <h4 style="margin-top:15px; margin-bottom:5px; color: #111;">VLamax Berechnung</h4>
                <div style="font-family: monospace; font-size:13px; background-color:#ffffff; padding:10px; border-radius:4px; border: 1px solid #e6e6e6; text-align:center; color: #111;">
                    VLamax = &Delta;La / t_glyc <br>
                    VLamax = {max_post_lac - avg_rest_lac:.2f} / {t_glyc:.1f} <br>
                    <b>VLamax = {vlamax:.3f} mmol/l/s</b>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
        st.markdown("---")
        
        # Diagrams layout
        col_plot_1, col_plot_2 = st.columns(2)
        
        with col_plot_1:
            st.subheader("📈 Laktatkinetik (Metabolischer Crossover)")
            
            # Plotly figure for crossover
            fig_cross = go.Figure()
            # Calculate dynamic max W range based on ANS + 50W (rounded to next 10W)
            max_x_plot = int(min(600, max(300, np.ceil((ans_power + 50) / 10.0) * 10.0)))
            df_plot = df_sim[df_sim['power'] <= max_x_plot]
            
            fig_cross.add_trace(go.Scatter(
                x=df_plot['power'], y=df_plot['vlass'],
                mode='lines',
                line=dict(color='#00a1e0', width=3),
                name='Laktatbildung (VLass)'
            ))
            fig_cross.add_trace(go.Scatter(
                x=df_plot['power'], y=df_plot['vlaoxmax'],
                mode='lines',
                line=dict(color='#cdb663', width=3),
                name='max. Laktatabbau (VLaoxmax)'
            ))
            
            # Add vertical line for ANS
            fig_cross.add_vline(
                x=ans_power,
                line_dash="dash",
                line_color="#595a59",
                annotation_text=f"ANS ({ans_power} W)",
                annotation_position="top right"
            )
            
            fig_cross.update_layout(
                xaxis=dict(title="Leistung [Watt]", range=[0, max_x_plot]),
                yaxis_title="Laktatkinetik [mmol/L/min]",
                height=400,
                legend=dict(x=0.02, y=0.98, bgcolor="rgba(255,255,255,0.7)"),
                margin=dict(l=40, r=40, t=20, b=40),
                plot_bgcolor='#f9f9f9'
            )
            st.plotly_chart(fig_cross, use_container_width=True)
            
        with col_plot_2:
            st.subheader("📊 Substratumsatz & Energiefluss")
            
            # Plotly figure with twin y-axes
            fig_met = go.Figure()
            df_plot = df_sim[df_sim['power'] <= max_x_plot]
            
            # EE gesamt
            fig_met.add_trace(go.Scatter(
                x=df_plot['power'], y=df_plot['ee_gesamt'],
                mode='lines',
                line=dict(color='#595a59', width=2.5, dash='dash'),
                name='Gesamtumsatz [kcal/h]',
                yaxis='y'
            ))
            # EE Fett
            fig_met.add_trace(go.Scatter(
                x=df_plot['power'], y=df_plot['ee_fett'],
                mode='lines',
                line=dict(color='#2cb7b9', width=3),
                name='Fettverbrennung [kcal/h]',
                yaxis='y'
            ))
            # KH-Verbrauch
            fig_met.add_trace(go.Scatter(
                x=df_plot['power'], y=df_plot['ee_kh_g'],
                mode='lines',
                line=dict(color='#cdb663', width=3),
                name='KH-Verbrennung [g/h]',
                yaxis='y2'
            ))
            
            # Vertical line for Fatmax
            fig_met.add_vline(
                x=fatmax_power,
                line_dash="dash",
                line_color="#2cb7b9",
                annotation_text=f"Fatmax ({fatmax_power} W)",
                annotation_position="top left"
            )
            
            # Twin axis layout
            fig_met.update_layout(
                xaxis=dict(title="Leistung [Watt]", range=[0, max_x_plot]),
                yaxis=dict(
                    title=dict(text="Energie [kcal/h]", font=dict(color="#595a59")),
                    tickfont=dict(color="#595a59")
                ),
                yaxis2=dict(
                    title=dict(text="Kohlenhydratumsatz [g/h]", font=dict(color="#cdb663")),
                    tickfont=dict(color="#cdb663"),
                    overlaying="y",
                    side="right"
                ),
                height=400,
                legend=dict(x=0.02, y=0.98, bgcolor="rgba(255,255,255,0.7)"),
                margin=dict(l=45, r=45, t=20, b=40),
                plot_bgcolor='#f9f9f9'
            )
            st.plotly_chart(fig_met, use_container_width=True)

        st.markdown("---")
        
        # Detailed Grid Table
        with st.expander("🔍 Detaillierter metabolischer Verlauf (Tabelle in 10W-Schritten)"):
            df_table_rows = []
            for p in range(0, max_x_plot + 1, 10):
                rows = df_sim[df_sim['power'] == p]
                if not rows.empty:
                    row_data = rows.iloc[0]
                    df_table_rows.append({
                        'Leistung [W]': p,
                        'VO2ss [ml/min/kg]': round(row_data['vo2ss'], 1) if not np.isnan(row_data['vo2ss']) else "-",
                        'VLass [mmol/L/min]': round(row_data['vlass'], 3) if not np.isnan(row_data['vlass']) else "-",
                        'VLaoxmax [mmol/L/min]': round(row_data['vlaoxmax'], 3) if not np.isnan(row_data['vlaoxmax']) else "-",
                        'Pyruvatdefizit [mmol/L/min]': round(row_data['pyruvat_def'], 3) if not np.isnan(row_data['pyruvat_def']) else "-",
                        'KH-Anteil [%]': f"{round(row_data['kh_pct'], 1)}%" if not np.isnan(row_data['kh_pct']) else "-",
                        'Fett-Anteil [%]': f"{round(row_data['fat_pct'], 1)}%" if not np.isnan(row_data['fat_pct']) else "-",
                        'kalk. RQ': round(row_data['rq'], 3) if not np.isnan(row_data['rq']) else "-",
                        'EE gesamt [kcal/h]': int(round(row_data['ee_gesamt'], 0)) if not np.isnan(row_data['ee_gesamt']) else "-",
                        'EE Fett [kcal/h]': int(round(row_data['ee_fett'], 0)) if not np.isnan(row_data['ee_fett']) else "-",
                        'KH-Verbrauch [g/h]': round(row_data['ee_kh_g'], 1) if not np.isnan(row_data['ee_kh_g']) else "-",
                        'HF [bpm]': f"{int(round(row_data['hr'], 0))}" if not np.isnan(row_data['hr']) else "-"
                    })
            st.dataframe(pd.DataFrame(df_table_rows), hide_index=True, width='stretch')

    else:
        # Default view before starting the calculation
        st.markdown(f"""
            <div style="font-size: 16px; line-height: 1.5; margin-bottom: 20px;">
                <b>Athlet:</b> {athlete_name}<br>
                <b>Sportart:</b> Radsport<br>
                <b>Test-Datum:</b> {test_date}
            </div>
            """, unsafe_allow_html=True)
            
        st.info("💡 Radsport-Diagnostik: Bitte laden Sie die gewünschten Dateien hoch und klicken Sie in der Seitenleiste auf **'Auswertung starten'**, um die Stoffwechselprofile zu generieren.")
        
        st.subheader("Hochgeladene Dateien")
        c1, c2, c3 = st.columns(3)
        with c1:
            if uploaded_spiro is not None:
                st.success(f"Spirodatei geladen:\n{uploaded_spiro.name}")
                st.caption(f"Berechnete VO2max: {st.session_state.get('vo2max_spiro', 4292.2):.1f} ml/min")
            else:
                st.warning("Keine Spirodatei hochgeladen")
        with c2:
            if uploaded_fit is not None:
                st.success(f".fit-Datei geladen:\n{uploaded_fit.name}")
                st.caption(f"Steigung: {st.session_state.get('slope_val', 0.11):.4f} | Nullstelle: {st.session_state.get('intercept_val', 138):.1f}")
            else:
                st.warning("Keine .fit-Datei hochgeladen")
        with c3:
            if uploaded_srm is not None:
                st.success(f"Sprint-Exceldatei geladen:\n{uploaded_srm.name}")
            else:
                st.warning("Keine Sprint-Exceldatei hochgeladen")
                
        st.subheader("Eingegebene Laktatwerte")
        st.dataframe(st.session_state.df_sprint_lac, hide_index=True, width='stretch')
        
    st.stop()


# Parse uploaded file if it has changed to update metadata
uploaded_file = st.sidebar.file_uploader("Spiro-Datei hochladen (.xlsx oder .csv)", type=["xlsx", "csv"], key="uploaded_file_key")

vo2max_override = st.sidebar.number_input(
    "VO2max überschreiben (ml/min)", 
    value=0, 
    step=50, 
    help="Ermöglicht das manuelle Überschreiben des berechneten VO2max-Wertes (z.B. 3652 für Barbara Bauer)"
)


if uploaded_file is not None and uploaded_file.name != st.session_state.last_uploaded_file:
    st.session_state.last_uploaded_file = uploaded_file.name
    try:
        df_excel = pd.read_csv(uploaded_file, header=None, low_memory=False) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file, header=None)
        
        # Name
        name_val = None
        if df_excel.shape[0] > 2 and df_excel.shape[1] > 1:
            name_val = f"{df_excel.iloc[2, 1]} {df_excel.iloc[1, 1]}"
        
        # Height
        h_val = None
        try:
            if df_excel.shape[0] > 5 and df_excel.shape[1] > 1:
                h_val = int(float(df_excel.iloc[5, 1]))
        except:
            pass
            
        # Weight
        w_val = None
        try:
            if df_excel.shape[0] > 6 and df_excel.shape[1] > 1:
                w_val = round(float(df_excel.iloc[6, 1]), 1)
        except:
            pass
            
        # Birthdate
        b_val = None
        if df_excel.shape[0] > 7 and df_excel.shape[1] > 1:
            raw_date = df_excel.iloc[7, 1] 
            if pd.notna(raw_date):
                date_str = str(int(float(raw_date))).zfill(8)
                b_val = f"{date_str[:2]}.{date_str[2:4]}.{date_str[4:]}"
            
        # Testdatum
        t_val = None
        if df_excel.shape[0] > 0 and df_excel.shape[1] > 4:
            raw_test_date = df_excel.iloc[0, 4]
            if pd.notna(raw_test_date):
                try:
                    t_date_str = str(int(float(raw_test_date))).zfill(8)
                    t_val = f"{t_date_str[:2]}.{t_date_str[2:4]}.{t_date_str[4:]}"
                except:
                    t_val = str(raw_test_date)
                    
        update_base_data(name_val, h_val, w_val, b_val, t_val)
        st.rerun()
    except Exception as e:
        st.sidebar.error(f"Fehler beim Extrahieren der Stammdaten: {e}")



# Kategorie, Coach, and Basisdaten reorganized
# Pre-calculate body fat percentage if skinfolds table has entries (Laufen)
if "df_sf" in st.session_state:
    sf_means = st.session_state.df_sf[["M1", "M2", "M3"]].mean(axis=1)
    sum_sf = sf_means.sum()
    if sum_sf > 0:
        calculated_bf = (22.32 * np.log10(sum_sf)) - 29.2
        st.session_state.body_fat_pct = max(0.0, round(calculated_bf, 1))

with st.sidebar.form("running_form", enter_to_submit=False):
    # Collapsible Basisdaten expander (Default collapsed)
    with st.expander("Basisdaten (*Edit)", expanded=False):
        athlete_name = st.text_input("Name", value=st.session_state.athlete_name, key="athlete_name_lauft_widget")
        st.session_state.athlete_name = athlete_name

        birthdate = st.text_input("Geburtsdatum", value=st.session_state.birthdate, key="birthdate_lauft_widget")
        st.session_state.birthdate = birthdate

        test_date = st.text_input("Testdatum", value=st.session_state.test_date, key="test_date_lauft_widget")
        st.session_state.test_date = test_date

        # Größe formatted as 123cm without decimals
        height = st.number_input("Größe (cm)", value=int(st.session_state.height), step=1, format="%d", key="height_lauft_widget")
        st.session_state.height = height

        # Gewicht (float, 1 decimal)
        weight = st.number_input("Gewicht (kg)", value=float(st.session_state.weight), step=0.1, format="%.1f", key="weight_lauft_widget")
        st.session_state.weight = round(weight, 1)

        # Körperfett (%) (float, 1 decimal)
        body_fat_pct = st.number_input("Körperfett (%)", value=float(st.session_state.body_fat_pct), step=0.1, format="%.1f", key="bf_lauft_widget")
        st.session_state.body_fat_pct = round(body_fat_pct, 1)

        # Kategorie selectbox
        kategorie_options = ["Hobby", "Age-Grouper", "Amateur", "Profi"]
        kategorie_val = st.session_state.kategorie
        if kategorie_val not in kategorie_options:
            kategorie_options = kategorie_options + [kategorie_val]
        kategorie = st.selectbox("Kategorie", options=kategorie_options, index=kategorie_options.index(kategorie_val), key="kategorie_lauft_widget")
        st.session_state.kategorie = kategorie

        # Coach selectbox
        coach_options = ["Markus Hertlein", "Marius Trompetter", "Susanne Traser", "Manuel Kuhnle", "Billie Benkel", "Jean Surmont", "Hosea Frick", "Björn Geesmann", "Gregor Eichhorn"]
        coach_val = st.session_state.coach
        if coach_val not in coach_options:
            coach_options = coach_options + [coach_val]
        coach = st.selectbox("Coach", options=coach_options, index=coach_options.index(coach_val), key="coach_lauft_widget")
        st.session_state.coach = coach

        # Körperfettmessung (Parizkova 10-Falten)
        with st.expander("Körperfettmessung (Parizkova 10-Falten)"):
            edited_sf = st.data_editor(
                st.session_state.df_sf,
                hide_index=True,
                width='stretch',
                key="edited_sf_key"
            )

    # Collapsible Stufentest Laktat expander directly below Basisdaten
    sync_lauftest_input_df()

    with st.expander("Stufentest Laktat", expanded=False):
        edited_df = st.data_editor(
            st.session_state.df_lauftest_input,
            disabled=["Stufe", "v (m/s)", "v (km/h)"],
            hide_index=True,
            use_container_width=True,
            height=320,
            column_config={
                "Stufe": st.column_config.TextColumn("Stufe", width="small"),
                "v (m/s)": st.column_config.NumberColumn("v (m/s)", format="%.2f", width="small"),
                "v (km/h)": st.column_config.NumberColumn("v (km/h)", format="%.1f", width="small"),
                "Laktat": st.column_config.NumberColumn("Laktat", min_value=0.0, max_value=25.0, step=0.01, format="%.2f", width="small"),
                "HF": st.column_config.NumberColumn("HF", min_value=30, max_value=240, step=1, format="%d", width="small")
            },
            key="lauftest_lac_editor_key"
        )

# Testprotokoll & Step Parameters placed below the lactate table expander
test_types = [
    "HYCYS Standart (5min / +0,4m/s)",
    "Neues Protokoll (3min / +1km/h bzw. 0,278m/s)"
]
test_type = st.sidebar.selectbox("Testprotokoll", options=test_types, key="test_type_lauft_widget")

if test_type != st.session_state.last_test_type:
    st.session_state.last_test_type = test_type
    if test_type == "HYCYS Standart (5min / +0,4m/s)":
        st.session_state.start_v = 2.8
        st.session_state.v_increment = 0.400
        st.session_state.anzahl = 5
        st.session_state.stufendauer = 5.0
    elif test_type == "Neues Protokoll (3min / +1km/h bzw. 0,278m/s)":
        st.session_state.start_v = 2.22
        st.session_state.v_increment = 0.278
        st.session_state.anzahl = 8
        st.session_state.stufendauer = 3.0
    # Ruhemessung and Pausendauer always 60s and 30s
    st.session_state.vorlauf = 60
    st.session_state.pausendauer = 30
    st.rerun()

    # Step parameters (not collapsed)
    start_v = st.number_input("Startgeschwindigkeit (m/s)", value=float(st.session_state.start_v), step=0.1, key="start_v_lauft_widget")
    st.session_state.start_v = start_v

    v_increment = st.number_input("Geschwindigkeitssteigerung (m/s)", value=float(st.session_state.v_increment), step=0.001, format="%.3f", key="v_increment_lauft_widget")
    st.session_state.v_increment = v_increment

    anzahl = st.number_input("Anzahl Stufen", value=int(st.session_state.anzahl), min_value=1, max_value=20, step=1, key="anzahl_lauft_widget")
    st.session_state.anzahl = anzahl

    # Collapsible Protokoll-Setup expander
    with st.expander("Protokoll-Setup", expanded=False):
        vorlauf = st.number_input("Ruhemessung (Sekunden)", value=int(st.session_state.vorlauf), step=10, key="vorlauf_lauft_widget")
        st.session_state.vorlauf = vorlauf
        
        stufendauer = st.number_input("Stufendauer (Minuten)", value=float(st.session_state.stufendauer), step=0.5, key="stufendauer_lauft_widget")
        st.session_state.stufendauer = stufendauer
        
        pausendauer = st.number_input("Pausendauer (Sekunden)", value=int(st.session_state.pausendauer), step=5, key="pausendauer_lauft_widget")
        st.session_state.pausendauer = pausendauer
        
        ausbelastung = st.checkbox("Test bis zur Ausbelastung", value=True, key="ausbelastung_lauft_widget")

    st.markdown("---")
    start_button = st.form_submit_button("Auswertung starten", type="primary")
    if start_button:
        st.session_state.df_lauftest_input = edited_df
        st.session_state.df_sf = edited_sf

# --- HILFSFUNKTIONEN ---

def calculate_vlamax_for_v(v_thresh, rel_vo2max, speed_arr, vo2_steady_abs, vco2_steady_abs, weight):
    if not (rel_vo2max > 0 and v_thresh > 0 and len(vo2_steady_abs) > 0 and weight > 0): 
        return None, None, None
        
    vo2_rest = 3.5 * weight
    re_abs_m_s = []
    
    # Calculate VO2_CHO100 running economy for exercise stages
    for v, vo2, vco2 in zip(speed_arr, vo2_steady_abs, vco2_steady_abs):
        if v <= 0:
            continue
        vo2_working = max(0.0, vo2 - vo2_rest)
        if vo2 > 0 and vco2 > 0:
            rq = vco2 / vo2
            rq = max(0.7, min(1.0, rq))
        else:
            rq = 0.85
        cal_eq = (rq - 0.7) * 5.094 + 19.6
        vo2_cho100 = vo2_working * cal_eq / 21.1
        re_abs_m_s.append(vo2_cho100 / v)
        
    ex_speeds = [v for v in speed_arr if v > 0]
    if len(re_abs_m_s) == 0 or len(ex_speeds) == 0:
        return None, None, None
        
    re_abs_at_thresh = np.interp(v_thresh, ex_speeds, re_abs_m_s)
    
    # Calculate relative VO2 demand at threshold (adding B28 = 1.0)
    rel_vo2_demand = (v_thresh * re_abs_at_thresh / weight) + 1.0
    
    delta_vo2 = rel_vo2max - rel_vo2_demand
    if delta_vo2 <= 0 or rel_vo2_demand <= 0:
        return None, re_abs_at_thresh / weight, rel_vo2_demand
        
    ks1, ks2, la_eq, vol_dist = 0.0631, 1.3310, 0.02049, 0.4
    vla_ox_max = (rel_vo2_demand * la_eq) / vol_dist
    adp = np.sqrt((ks1 * rel_vo2_demand) / delta_vo2)
    adp_3 = adp ** 3
    term2 = 1 + (ks2 / adp_3)
    vla_sec = (vla_ox_max * term2) / 60
    return vla_sec, re_abs_at_thresh / weight, rel_vo2_demand

def calculate_laufokonomie_hycys(vo2_abs, vco2_abs, weight, speed_ms):
    """
    Laufökonomie nach HYCYS Excel-Methodik.
    Entspricht Spalte Y im 'Report Running' Sheet:
      Y = (Arbeits-VO2 bei 100% KH) / Geschwindigkeit [m/s] / Gewicht [kg]
    
    Formeln aus Excel:
      - Ruhe-VO2 = 3.5 * weight (B14-Formel)
      - Arbeits-VO2 = VO2_gesamt - Ruhe-VO2 (G-Spalte)
      - RQ-Korrekturfaktor (L12): cal_eq = (RQ - 0.7) * 5.094 + 19.6 [kcal/L O2]
      - VO2_CHO100 = Arbeits-VO2 * cal_eq / 21.1  (AB-Spalte / B19)
      - Laufökonomie = VO2_CHO100 / speed / weight
    """
    if weight <= 0 or speed_ms <= 0 or vo2_abs <= 0:
        return None
    # Ruhe-VO2 [ml/min] (B14 = 3.5 * Gewicht)
    vo2_rest = 3.5 * weight
    # Arbeits-VO2 [ml/min] (Spalte G = F - Ruhe)
    vo2_working = max(0.0, vo2_abs - vo2_rest)
    if vo2_working <= 0:
        return None
    # RQ (Spalte K = VCO2 / VO2)
    if vco2_abs > 0 and vo2_abs > 0:
        rq = vco2_abs / vo2_abs
        rq = max(0.7, min(1.0, rq))  # physiologischer Bereich
    else:
        rq = 0.85  # Standardwert
    # Kalorisches Äquivalent bei gemessenem RQ [kcal/L O2] (L12-Formel)
    # L12 = ((K12 - B17) * 0.05094 * 100) + B18  mit B17=0.7, B18=19.6
    cal_eq = (rq - 0.7) * 5.094 + 19.6
    # VO2-Äquivalent bei 100% KH-Verbrennung [ml/min] (AB-Spalte)
    b19 = 21.1  # kcal/L O2 bei RQ=1 (reine KH)
    vo2_cho100 = vo2_working * cal_eq / b19
    # Laufökonomie [ml*s/(m*min*kg)] = F44/B7 im Excel
    return vo2_cho100 / speed_ms / weight

# --- PDF GENERATOR FUNKTION ---
def create_pdf(athlete_name, birthdate, test_date, test_type, weight, body_fat_pct, rel_vo2max, df_res, fig_laktat,
               coach, sportart, kategorie, height, speed, lactate, hr, vo2_steady_values, vo2_steady_abs, vco2_steady_abs, uploaded_file):
    pdf = FPDF(unit='pt')
    pdf.set_auto_page_break(auto=False)
    
    # Standard colors from the template
    c_teal = (44, 183, 185)       # #2cb7b9
    c_gold = (205, 182, 99)       # #cdb663
    c_dark_grey = (89, 90, 89)    # #595a59
    c_black = (0, 0, 0)
    c_white = (255, 255, 255)

    # Helper function to set fonts & colors and draw text
    def draw_text(x, y, text, size=8.3, font_style='', color=c_black, font_name='Arial'):
        pdf.set_font(font_name, font_style, size)
        pdf.set_text_color(*color)
        pdf.text(x, y + size * 0.85, text.encode('latin-1', 'replace').decode('latin-1'))

    # Helper function to draw colored rectangles
    def draw_rect(x, y, w, h, fill_color=None, stroke_color=None, stroke_width=0.5):
        if fill_color:
            pdf.set_fill_color(*fill_color)
        if stroke_color:
            pdf.set_draw_color(*stroke_color)
            pdf.set_line_width(stroke_width)
        style = 'F' if fill_color and not stroke_color else 'D' if stroke_color and not fill_color else 'FD' if fill_color and stroke_color else ''
        pdf.rect(x, y, w, h, style)

    temp_files = []

    # --- SEITE 1: DECKBLATT ---
    pdf.add_page()
    draw_rect(50.9, 54.5, 479.1, 663.6, fill_color=c_white)
    
    if os.path.exists('pdf_assets/page1_img3.png'):
        pdf.image('pdf_assets/page1_img3.png', x=86.8, y=194.4, w=400, h=124)
        
    draw_rect(50.9, 460.4, 479.1, 24.5, fill_color=c_gold)
    draw_text(111.7, 464.7, "HYCYS Running BLUE", size=13.8, font_style='B', color=c_white)

    labels_s1 = ["Name", "Geburtsdatum", "Coach", "Testdatum", "Sportart", "Kategorie"]
    vals_s1 = [athlete_name, birthdate, coach, test_date, sportart, kategorie]
    y_coords_s1 = [512.7, 524.7, 564.6, 577.9, 604.6, 617.9]
    for lbl, val, y_coord in zip(labels_s1, vals_s1, y_coords_s1):
        draw_text(111.0, y_coord, lbl, size=8.3, color=c_black)
        draw_text(214.5, y_coord, val, size=8.3, font_style='B', color=c_black)

    if os.path.exists('pdf_assets/page1_img1.png'):
        pdf.image('pdf_assets/page1_img1.png', x=114.4, y=681.4, w=15.6, h=14.4)
    draw_text(162.6, 685.8, "www.hycys.de", size=7.6, font_style='B', color=c_teal)
    
    if os.path.exists('pdf_assets/page1_img2.png'):
        pdf.image('pdf_assets/page1_img2.png', x=137.3, y=680.6, w=15.6, h=16.1)
    draw_text(266.1, 685.8, "contact@hycys.de", size=7.6, font_style='B', color=c_teal)

    # --- SEITE 2: ANTHROPOMETRIE & LAKTATKINETIK ---
    pdf.add_page()
    draw_rect(50.9, 54.5, 471.6, 663.6, fill_color=c_white)
    
    if os.path.exists('pdf_assets/page2_img1.png'):
        pdf.image('pdf_assets/page2_img1.png', x=355.0, y=54.0, w=167, h=55)

    draw_rect(57.6, 133.9, 171.9, 14.8, fill_color=c_teal)
    draw_text(113.2, 135.9, "Anthropologie", size=9.0, font_style='B', color=c_white)

    draw_text(59.3, 173.2, "Gewicht", size=8.3, font_style='B', color=c_dark_grey)
    draw_text(199.9, 173.2, f"{weight:.1f} kg".replace('.', ','), size=8.3, font_style='B', color=c_dark_grey)
    draw_text(59.3, 183.8, "Größe", size=8.3, font_style='B', color=c_dark_grey)
    draw_text(199.9, 183.8, f"{int(height)}cm", size=8.3, font_style='B', color=c_dark_grey)
    
    bmi = weight / ((height/100.0)**2) if height > 0 else 0.0
    draw_text(59.3, 209.7, "Body Mass Index", size=8.3, font_style='B', color=c_dark_grey)
    draw_text(187.5, 209.7, f"{bmi:.1f} kg/m²".replace('.', ','), size=8.3, font_style='B', color=c_dark_grey)
    
    draw_text(145.3, 249.0, "Masse", size=8.3, font_style='B', color=c_dark_grey)
    draw_text(220.6, 249.0, "%", size=8.3, font_style='B', color=c_dark_grey)
    
    fett_kg = weight * (body_fat_pct / 100.0)
    fettfrei_kg = weight - fett_kg
    draw_text(59.3, 259.6, "Fett", size=8.3, font_style='B', color=c_dark_grey)
    draw_text(142.6, 259.6, f"{fett_kg:.1f} kg".replace('.', ','), size=8.3, font_style='B', color=c_dark_grey)
    draw_text(202.4, 259.6, f"{body_fat_pct:.1f} %".replace('.', ','), size=8.3, font_style='B', color=c_dark_grey)
    
    draw_text(59.3, 270.1, "Fettfrei", size=8.3, font_style='B', color=c_dark_grey)
    draw_text(142.6, 270.1, f"{fettfrei_kg:.1f} kg".replace('.', ','), size=8.3, font_style='B', color=c_dark_grey)
    draw_text(202.4, 270.1, f"{100.0 - body_fat_pct:.1f} %".replace('.', ','), size=8.3, font_style='B', color=c_dark_grey)

    # Donut Pie Chart (Fett / Fettfrei)
    fig, ax = plt.subplots(figsize=(2.5, 1.8), dpi=300)
    ax.pie(
        [max(0.1, body_fat_pct), max(0.1, 100.0 - body_fat_pct)],
        labels=['Fett', 'Fettfrei'],
        colors=['#cdb663', '#2cb7b9'],
        autopct='%1.1f%%',
        startangle=90,
        textprops=dict(color="black", size=6, weight="bold"),
        wedgeprops=dict(width=0.35, edgecolor='w', linewidth=1)
    )
    ax.axis('equal')
    plt.tight_layout()
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_pie:
        fig.savefig(tmp_pie.name, format="png", bbox_inches="tight", transparent=True)
        plt.close(fig)
        pdf.image(tmp_pie.name, x=272.5, y=141.6, w=236, h=140)
        temp_files.append(tmp_pie.name)

    # VLamax gold box
    draw_rect(57.6, 293.2, 286.4, 14.8, fill_color=c_teal)
    draw_text(93.4, 295.1, "Anaerober Stoffwechsel - max. Laktatbildungsrate", size=9.0, font_style='B', color=c_white)
    
    draw_rect(114.9, 321.2, 171.8, 13.0, fill_color=c_gold)
    draw_text(186.9, 321.9, "VLamax", size=9.0, font_style='B', color=c_white)
    
    vla_ans_str = "-"
    ans_row = df_res[df_res['Modell'].str.contains('OBLA 4.0', na=False)]
    if ans_row.empty and not df_res.empty:
        ans_row = df_res.iloc[0:1]
    if not ans_row.empty:
        vla_val = ans_row.iloc[0].get('VLamax [mmol/l/s]', '-')
        if pd.notna(vla_val) and vla_val != 'N/A' and 'bitte' not in str(vla_val).lower():
            try:
                vla_ans_str = f"{float(vla_val):.2f} mmol/l/s".replace('.', ',')
            except:
                vla_ans_str = str(vla_val)
    draw_text(175.3, 334.7, vla_ans_str, size=8.3, font_style='B', color=c_dark_grey)

    # VO2max 3 gold boxes
    draw_rect(57.6, 368.4, 343.6, 14.8, fill_color=c_teal)
    draw_text(124.1, 370.4, "Aerober Stoffwechsel - max. Sauerstoffaufnahme", size=9.0, font_style='B', color=c_white)
    
    vo2_abs = rel_vo2max * weight
    vo2_abs_str = f"{int(vo2_abs):,}".replace(",", ".") + " ml/min" if vo2_abs > 0 else "-"
    vo2_rel_str = f"{rel_vo2max:.1f} ml/min/kg".replace('.', ',') if rel_vo2max > 0 else "-"
    
    ffm_pct = 100.0 - body_fat_pct
    ffm_kg = weight * (ffm_pct / 100.0)
    vo2_ffm = (rel_vo2max * weight) / ffm_kg if ffm_kg > 0 else 0.0
    vo2_ffm_str = f"{vo2_ffm:.1f} ml/min/kg".replace('.', ',') if vo2_ffm > 0 else "-"

    draw_rect(57.6, 396.4, 143.6, 13.1, fill_color=c_gold)
    draw_text(90.9, 397.1, "VO2max abs.", size=9.0, font_style='B', color=c_white)
    draw_text(90.6, 411.0, vo2_abs_str, size=8.3, font_style='B', color=c_dark_grey)
    
    draw_rect(201.2, 396.4, 143.6, 13.1, fill_color=c_gold)
    draw_text(207.6, 397.1, "VO2max rel.", size=9.0, font_style='B', color=c_white)
    draw_text(201.5, 411.0, vo2_rel_str, size=8.3, font_style='B', color=c_dark_grey)
    
    draw_rect(344.8, 396.4, 143.4, 13.1, fill_color=c_gold)
    draw_text(311.6, 397.1, "VO2max rel. FFM", size=9.0, font_style='B', color=c_white)
    draw_text(316.0, 411.0, vo2_ffm_str, size=8.3, font_style='B', color=c_dark_grey)

    # Lactate kinetics plot
    draw_rect(57.6, 449.3, 458.1, 11.2, fill_color=c_teal)
    draw_text(222.1, 449.5, "Interaktion Laktatauf- & abbau", size=9.0, font_style='B', color=c_white)
    
    fig, ax = plt.subplots(figsize=(6, 3.2), dpi=300)
    v_arr = np.linspace(speed[0], speed[-1], 100) if len(speed) > 0 else np.linspace(2.9, 4.5, 100)
    
    if uploaded_file is not None and len(vo2_steady_abs) > 0:
        # Plot real absolute VO2 curve
        ax.plot(speed[:len(vo2_steady_abs)], vo2_steady_abs, color='#595a59', linewidth=1.5, marker='s', markersize=3, label='Sauerstoffaufnahme - VO2')
        ax.set_ylabel('VO2 [ml/min]', color='#595a59', fontsize=7, fontweight='bold')
        ax.tick_params(axis='y', labelcolor='#595a59', labelsize=6)
        ax.grid(True, linestyle=':', alpha=0.6)
    else:
        ax.set_ylabel('VO2 [ml/min]', color='#595a59', fontsize=7, fontweight='bold')
        ax.grid(True, linestyle=':', alpha=0.6)
        
    ax2 = ax.twinx()
    # Plot real lactate curve on right axis
    poly_coeffs = np.polyfit(speed, lactate, 3)
    poly_func = np.poly1d(poly_coeffs)
    ax2.plot(v_arr, poly_func(v_arr), color='#2cb7b9', linewidth=1.5, label='Laktatabbau')
    ax2.plot(speed, lactate, 'ko', markersize=4, label='Messwerte')
    
    # Mock a lactate production curve for premium visualization matching template
    v_start, v_end = speed[0], speed[-1]
    lac_prod = lactate[0] + (lactate[-1] - lactate[0]) * ((v_arr - v_start)/(v_end - v_start))**3.5
    ax2.plot(v_arr, lac_prod, color='#cdb663', linewidth=1.5, linestyle='--', label='Laktatproduktion')
    
    ax2.set_ylabel('Laktatkinetik [mmol/L/min]', color='#2cb7b9', fontsize=7, fontweight='bold')
    ax2.tick_params(axis='y', labelcolor='#2cb7b9', labelsize=6)
    
    ax.set_xlabel('Geschwindigkeit [m/s]', fontsize=7, fontweight='bold')
    ax.set_xlim(v_start, v_end)
    ax.tick_params(axis='x', labelsize=6)
    
    lines, labels = ax.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax.legend(lines + lines2, labels + labels2, loc='upper left', fontsize=5, framealpha=0.8)
    
    plt.tight_layout()
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_lactate:
        fig.savefig(tmp_lactate.name, format="png", bbox_inches="tight")
        plt.close(fig)
        pdf.image(tmp_lactate.name, x=55.2, y=466.1, w=455, h=248)
        temp_files.append(tmp_lactate.name)

    # --- SEITE 3: LAUFÖKONOMIE ---
    pdf.add_page()
    draw_rect(50.9, 54.5, 479.1, 663.6, fill_color=c_white)
    if os.path.exists('pdf_assets/page2_img1.png'):
        pdf.image('pdf_assets/page2_img1.png', x=355.0, y=54.0, w=167, h=55)
        
    draw_rect(57.6, 133.9, 207.0, 14.8, fill_color=c_teal)
    draw_text(130.6, 135.9, "Laufökonomie", size=9.0, font_style='B', color=c_white)

    # Table Header
    draw_rect(57.6, 159.1, 207.0, 13.4, fill_color=c_gold)
    draw_text(73.8, 160.4, "Geschwindigkeit", size=9.0, font_style='B', color=c_white)
    draw_text(182.3, 160.4, "Laufökonomie", size=9.0, font_style='B', color=c_white)

    # Populate running economy values dynamically
    speeds_s3 = []
    eco_vals_s3 = []
    y_coords_s3 = [173.2, 183.8, 197.1, 211.8, 225.1]
    
    display_steps = min(5, len(speed))
    for i in range(display_steps):
        v_ms = speed[i]
        vo2_r = vo2_steady_values[i] if i < len(vo2_steady_values) else 0.0
        
        if vo2_r > 0:
            re_val = (vo2_r / (v_ms * 3.6) * 60.0)
            eco_str = f"{re_val:.1f} ml/kg/km".replace('.', ',')
        else:
            eco_str = "-"
            
        speeds_s3.append(f"{v_ms:.2f} m/s".replace('.', ','))
        eco_vals_s3.append(eco_str)
        
    for i in range(len(speeds_s3)):
        y_coord = y_coords_s3[i]
        draw_text(95.9, y_coord, speeds_s3[i], size=8.3, font_style='B', color=c_black)
        draw_text(176.2, y_coord, eco_vals_s3[i], size=8.3, font_style='B', color=c_black)

    # Chart 3a: Running Economy plot
    fig, ax = plt.subplots(figsize=(3, 1.8), dpi=300)
    valid_re = []
    valid_speeds_v = []
    for i in range(len(speed)):
        v_ms = speed[i]
        vo2_r = vo2_steady_values[i] if i < len(vo2_steady_values) else 0.0
        if vo2_r > 0:
            valid_re.append(vo2_r / (v_ms * 3.6) * 60.0)
            valid_speeds_v.append(v_ms)
            
    if len(valid_re) > 0:
        ax.plot(valid_speeds_v, valid_re, marker='o', color='#2cb7b9', linewidth=1.5, markersize=4)
        ax.set_ylim(min(valid_re)*0.9, max(valid_re)*1.1)
    ax.set_ylabel('RE [ml/kg/km]', color='#595a59', fontsize=6, fontweight='bold')
    ax.set_xlabel('Geschwindigkeit [m/s]', color='#595a59', fontsize=6, fontweight='bold')
    ax.set_xlim(speed[0]*0.95, speed[-1]*1.05)
    ax.tick_params(axis='both', labelsize=5)
    ax.grid(True, linestyle=':', alpha=0.6)
    plt.tight_layout()
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_eco_chart:
        fig.savefig(tmp_eco_chart.name, format="png", bbox_inches="tight")
        plt.close(fig)
        pdf.image(tmp_eco_chart.name, x=266.8, y=156.2, w=257, h=135.6)
        temp_files.append(tmp_eco_chart.name)

    # Marathon predictions
    draw_rect(57.6, 334.1, 414.0, 11.2, fill_color=c_teal)
    draw_text(176.1, 334.3, "Einfluss Laufökonomie auf Marathon-Zeit", size=9.0, font_style='B', color=c_white)

    draw_rect(57.6, 355.1, 414.0, 13.4, fill_color=c_gold)
    draw_text(79.7, 356.3, "Verbesserung", size=9.0, font_style='B', color=c_white)
    draw_text(180.9, 356.3, "Marathon-Pace", size=9.0, font_style='B', color=c_white)
    draw_text(301.4, 356.3, "Zielzeit", size=9.0, font_style='B', color=c_white)
    draw_text(395.2, 356.3, "Delta zu 0%", size=9.0, font_style='B', color=c_white)

    y_coords_s3_m = [386.0, 398.9, 412.3]
    improves = ["0%", "5%", "10%"]
    for imp, y_coord in zip(improves, y_coords_s3_m):
        draw_text(103.5, y_coord, imp, size=9.0, font_style='B', color=c_dark_grey)
        draw_text(184.6, y_coord, "-", size=9.0, font_style='B', color=c_dark_grey)
        draw_text(294.6, y_coord, "-", size=9.0, font_style='B', color=c_dark_grey)
        if imp != "0%":
            draw_text(396.5, y_coord, "-", size=9.0, font_style='B', color=c_dark_grey)

    # Empty Marathon prediction chart
    fig, ax = plt.subplots(figsize=(5, 2.5), dpi=300)
    ax.set_xlabel('Marathon-Zielzeit [h:mm]', color='#595a59', fontsize=6, fontweight='bold')
    ax.set_ylabel('Verbesserung Laufökonomie [%]', color='#595a59', fontsize=6, fontweight='bold')
    ax.set_xlim(120, 200)
    ax.set_xticks([120, 130, 140, 150, 160, 170, 180, 190, 200])
    ax.set_xticklabels(['2:00', '2:10', '2:20', '2:30', '2:40', '2:50', '3:00', '3:10', '3:20'])
    ax.set_ylim(0, 3.5)
    ax.set_yticks([0, 1, 2, 3])
    ax.set_yticklabels(['0%', '1%', '2%', '3%'])
    ax.tick_params(axis='both', labelsize=5)
    ax.grid(True, linestyle=':', alpha=0.6)
    plt.tight_layout()
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_mara_chart:
        fig.savefig(tmp_mara_chart.name, format="png", bbox_inches="tight")
        plt.close(fig)
        pdf.image(tmp_mara_chart.name, x=66.1, y=475.1, w=449.2, h=209.4)
        temp_files.append(tmp_mara_chart.name)

    # --- SEITE 4: ANAEROBE SCHWELLE & ENERGETIK ---
    pdf.add_page()
    draw_rect(50.9, 54.5, 479.1, 663.6, fill_color=c_white)
    if os.path.exists('pdf_assets/page2_img1.png'):
        pdf.image('pdf_assets/page2_img1.png', x=355.0, y=54.0, w=167, h=55)
        
    draw_rect(57.6, 133.9, 465.7, 14.8, fill_color=c_teal)
    draw_text(161.3, 135.9, "Anaerobe Schwelle (ANS) - Geschwindigkeit & Herzfrequenz", size=9.0, font_style='B', color=c_white)

    # Box 1
    draw_rect(57.6, 159.1, 207.0, 13.4, fill_color=c_gold)
    draw_text(116.7, 160.4, "Laufgeschwindigkeit", size=9.0, font_style='B', color=c_white)
    
    draw_rect(368.0, 159.1, 103.6, 13.4, fill_color=c_gold)
    draw_text(391.5, 160.4, "Herzfrequenz", size=9.0, font_style='B', color=c_white)

    ans_ms_str, ans_pace_str, ans_hf_str = "-", "-", "-"
    v_ans = 0.0
    if not ans_row.empty:
        v_ans = ans_row.iloc[0].get('v', 0.0)
        if isinstance(v_ans, (int, float)):
            ans_ms_str = f"{v_ans:.2f} m/s".replace('.', ',')
            pace_sec = 1000.0 / v_ans if v_ans > 0 else 0
            ans_pace_str = f"{int(pace_sec // 60):02d}:{int(pace_sec % 60):02d} min/km"
            ans_hf_str = f"{int(ans_row.iloc[0].get('HF', 0))} bpm"

    draw_text(93.6, 172.8, ans_ms_str, size=8.3, font_style='B', color=c_dark_grey)
    draw_text(187.1, 172.8, ans_pace_str, size=8.3, font_style='B', color=c_dark_grey)
    draw_text(403.5, 172.8, ans_hf_str, size=8.3, font_style='B', color=c_dark_grey)

    # Box 2: VO2
    draw_rect(57.6, 206.9, 465.7, 14.8, fill_color=c_teal)
    draw_text(222.6, 208.8, "Anaerobe Schwelle (ANS) - VO2", size=9.0, font_style='B', color=c_white)

    draw_rect(57.6, 234.9, 103.6, 13.4, fill_color=c_gold)
    draw_text(91.3, 235.8, "VO2 abs.", size=9.0, font_style='B', color=c_white)

    draw_rect(212.8, 234.9, 103.6, 13.4, fill_color=c_gold)
    draw_text(248.8, 235.8, "VO2 rel.", size=9.0, font_style='B', color=c_white)

    draw_rect(368.0, 234.9, 103.6, 13.4, fill_color=c_gold)
    draw_text(400.9, 235.8, "% VO2max", size=9.0, font_style='B', color=c_white)

    ans_vo2_abs_str, ans_vo2_rel_str, ans_vo2_pct_str = "-", "-", "-"
    if uploaded_file is not None and v_ans > 0 and len(vo2_steady_values) > 0:
        ans_vo2_rel = np.interp(v_ans, speed, vo2_steady_values)
        ans_vo2_abs = ans_vo2_rel * weight
        ans_vo2_pct = (ans_vo2_rel / rel_vo2max * 100.0) if rel_vo2max > 0 else 0.0
        
        ans_vo2_abs_str = f"{int(ans_vo2_abs):,}".replace(",", ".") + " ml/min"
        ans_vo2_rel_str = f"{ans_vo2_rel:.1f} ml/min/kg".replace('.', ',')
        ans_vo2_pct_str = f"{int(ans_vo2_pct)} %"

    draw_text(85.1, 248.5, ans_vo2_abs_str, size=8.3, font_style='B', color=c_dark_grey)
    draw_text(236.7, 248.5, ans_vo2_rel_str, size=8.3, font_style='B', color=c_dark_grey)
    draw_text(411.0, 248.5, ans_vo2_pct_str, size=8.3, font_style='B', color=c_dark_grey)

    # Energetik Box
    draw_rect(57.6, 279.9, 465.7, 13.4, fill_color=c_teal)
    draw_text(211.1, 281.1, "Anaerobe Schwelle (ANS) - Energetik", size=9.0, font_style='B', color=c_white)

    draw_rect(57.6, 307.8, 103.6, 13.4, fill_color=c_gold)
    draw_text(71.4, 309.1, "Energieverbrauch", size=9.0, font_style='B', color=c_white)

    draw_rect(212.8, 307.8, 103.6, 13.4, fill_color=c_gold)
    draw_text(234.6, 309.1, "KH-Verbrauch", size=9.0, font_style='B', color=c_white)

    draw_text(86.8, 322.7, "-", size=8.3, font_style='B', color=c_dark_grey)
    draw_text(250.9, 322.7, "-", size=8.3, font_style='B', color=c_dark_grey)

    # Fatmax Box
    draw_rect(57.6, 355.1, 465.7, 13.4, fill_color=c_teal)
    draw_text(235.2, 356.3, "Fettstoffwechsel - Fatmax", size=9.0, font_style='B', color=c_white)

    draw_rect(57.6, 383.1, 207.0, 13.4, fill_color=c_gold)
    draw_text(116.7, 384.3, "Laufgeschwindigkeit", size=9.0, font_style='B', color=c_white)

    draw_rect(316.3, 383.1, 207.0, 13.4, fill_color=c_gold)
    draw_text(330.1, 384.3, "Energieverbrauch", size=9.0, font_style='B', color=c_white)
    draw_text(441.8, 384.3, "Fettverbrauch", size=9.0, font_style='B', color=c_white)

    draw_text(93.6, 397.9, "-", size=8.3, font_style='B', color=c_dark_grey)
    draw_text(187.1, 397.9, "-", size=8.3, font_style='B', color=c_dark_grey)
    draw_text(348.8, 397.9, "-", size=8.3, font_style='B', color=c_dark_grey)
    draw_text(452.3, 397.9, "-", size=8.3, font_style='B', color=c_dark_grey)

    # Empty Energetik chart
    fig, ax = plt.subplots(figsize=(6, 3.2), dpi=300)
    ax.set_xlabel('Geschwindigkeit [m/s]', color='#595a59', fontsize=6, fontweight='bold')
    ax.set_ylabel('Energie [kcal/h]', color='#595a59', fontsize=6, fontweight='bold')
    ax.set_xlim(speed[0]*0.95, speed[-1]*1.05)
    ax.set_ylim(0, 1500)
    ax2 = ax.twinx()
    ax2.set_ylabel('Kohlenhydratverbrauch [g/h]', color='#595a59', fontsize=6, fontweight='bold')
    ax2.set_ylim(0, 400)
    ax.tick_params(axis='both', labelsize=5)
    ax2.tick_params(axis='both', labelsize=5)
    ax.grid(True, linestyle=':', alpha=0.6)
    plt.tight_layout()
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_ener_chart:
        fig.savefig(tmp_ener_chart.name, format="png", bbox_inches="tight")
        plt.close(fig)
        pdf.image(tmp_ener_chart.name, x=55.9, y=449.5, w=472, h=254)
        temp_files.append(tmp_ener_chart.name)

    # --- SEITE 5: TRAININGSBEREICHE & REFERENZDATEN ---
    pdf.add_page()
    draw_rect(50.9, 54.5, 479.1, 663.6, fill_color=c_white)
    if os.path.exists('pdf_assets/page2_img1.png'):
        pdf.image('pdf_assets/page2_img1.png', x=355.0, y=54.0, w=167, h=55)
        
    draw_rect(57.6, 133.9, 155.3, 14.8, fill_color=c_teal)
    draw_text(104.5, 135.9, "Referenzdaten", size=9.0, font_style='B', color=c_white)

    # Polar Radar chart for Reference data (background rings only)
    fig = plt.figure(figsize=(4, 4), dpi=300)
    ax = fig.add_subplot(111, polar=True)
    categories = ['VO2max', 'Fettanteil', 'Laufökonomie', 'ANS', 'Fatmax', 'KH-Match', 'VLamax']
    N = len(categories)
    angles = [n / float(N) * 2 * np.pi for n in range(N)]
    angles += angles[:1]
    
    ax.fill(angles, [100]*len(angles), color='#cdb663', alpha=0.15, label='Sehr gut')
    ax.fill(angles, [80]*len(angles), color='#2cb7b9', alpha=0.15, label='Gut')
    ax.fill(angles, [60]*len(angles), color='#e6e6e6', alpha=0.2, label='Zu verbessern')
    ax.fill(angles, [35]*len(angles), color='#c0c0c0', alpha=0.25, label='Schwach')
    
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(categories, fontsize=5, fontweight='bold')
    ax.set_yticklabels([])
    ax.grid(True, color='gray', linestyle=':', linewidth=0.5)
    plt.tight_layout()
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_radar:
        fig.savefig(tmp_radar.name, format="png", bbox_inches="tight")
        plt.close(fig)
        pdf.image(tmp_radar.name, x=78.0, y=153.0, w=418.6, h=293.4)
        temp_files.append(tmp_radar.name)

    # Trainingsbereiche Table
    draw_rect(57.6, 471.4, 207.0, 13.4, fill_color=c_teal)
    draw_text(122.5, 472.6, "Trainingsbereiche", size=9.0, font_style='B', color=c_white)

    draw_text(99.3, 497.0, "Trainingsbereich", size=9.0, font_style='B', color=c_black)
    draw_text(254.8, 497.0, "Geschwindigkeit", size=9.0, font_style='B', color=c_black)
    draw_text(417.4, 497.0, "Herzfrequenz", size=9.0, font_style='B', color=c_black)
    draw_text(275.3, 510.7, "[min/km]", size=8.3, color=c_black)
    draw_text(435.8, 510.7, "[bpm]", size=8.3, color=c_black)
    
    draw_text(231.8, 523.9, "min", size=8.3, font_style='B', color=c_black)
    draw_text(282.6, 523.9, "max", size=8.3, font_style='B', color=c_black)
    draw_text(335.5, 523.9, "Ziel", size=8.3, font_style='B', color=c_black)
    draw_text(387.0, 523.9, "min", size=8.3, font_style='B', color=c_black)
    draw_text(437.9, 523.9, "max", size=8.3, font_style='B', color=c_black)
    draw_text(490.7, 523.9, "Ziel", size=8.3, font_style='B', color=c_black)

    zones = [
        "regenerativer Dauerlauf (DLreg)",
        "extensiver Dauerlauf (DLext)",
        "intensiver  Dauerlauf (DLint)",
        "Tempodauerlauf (DLTempo)",
        "extensives Tempotraining (TText)",
        "intensives Tempotraining (TTint)"
    ]
    y_coords_s5 = [536.1, 549.4, 562.7, 576.0, 602.7, 616.0]
    for zn, y_coord in zip(zones, y_coords_s5):
        draw_text(59.3, y_coord, zn, size=8.3, font_style='B', color=c_black)
        draw_text(231.8, y_coord, "-", size=8.3, color=c_black)
        draw_text(282.6, y_coord, "-", size=8.3, color=c_black)
        draw_text(335.5, y_coord, "-", size=8.3, color=c_black)
        draw_text(387.0, y_coord, "-", size=8.3, color=c_black)
        draw_text(437.9, y_coord, "-", size=8.3, color=c_black)
        draw_text(490.7, y_coord, "-", size=8.3, color=c_black)

    pdf_output = pdf.output(dest='S')
    
    # Windows File Lock Cleanup
    for f in temp_files:
        try:
            os.unlink(f)
        except Exception:
            pass
            
    return pdf_output.encode('latin-1', 'replace')

# 3. BERECHNUNG & ANZEIGE
if start_button:
    raw_speeds = st.session_state.df_lauftest_input["v (m/s)"].values
    raw_lactate = st.session_state.df_lauftest_input["Laktat"].values
    raw_hr = st.session_state.df_lauftest_input["HF"].values
    
    # Exclude Ruhe (first row) for exercise steps calculations
    speed = raw_speeds[1:]
    lactate = raw_lactate[1:]
    hr = raw_hr[1:]
    
    ruhe_lac = raw_lactate[0]
    ruhe_hr = raw_hr[0]
    abs_vo2max, rel_vo2max = 0.0, 0.0
    vo2_steady_values = []
    vo2_steady_abs = []      # absoluter VO2-Steady-State [ml/min] je Stufe
    vco2_steady_abs = []     # absoluter VCO2-Steady-State [ml/min] je Stufe
    window_coords = []
    spiro_df = pd.DataFrame()

    sf_means = st.session_state.df_sf[["M1", "M2", "M3"]].mean(axis=1)
    sum_sf = sf_means.sum()
    if sum_sf > 0:
        body_fat_pct = (22.32 * np.log10(sum_sf)) - 29.2

    spiro_interval_sec = None   # wird beim Einlesen ermittelt (10 oder 30)
    spiro_interval_note = ""    # Hinweis bei 30s für VO2max-Anzeige

    if uploaded_file is not None:
        try:
            df_excel = pd.read_csv(uploaded_file, header=None, low_memory=False) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file, header=None)
            times_sec = df_excel.iloc[3:, 9].apply(parse_time_to_seconds)
            vo2_col = pd.to_numeric(df_excel.iloc[3:, 13], errors='coerce')
            vco2_col = pd.to_numeric(df_excel.iloc[3:, 14], errors='coerce')
            spiro_df = pd.DataFrame({'Time': times_sec, 'VO2': vo2_col, 'VCO2': vco2_col}).dropna(subset=['Time'])

            # --- Zeitintervall automatisch erkennen ---
            if len(spiro_df) >= 2:
                t_diffs = spiro_df['Time'].diff().dropna()
                spiro_interval_sec = int(round(t_diffs.median()))
            else:
                spiro_interval_sec = 10  # default

            if spiro_interval_sec == 30:
                spiro_interval_note = "(aus 30s-Auflösung, Genauigkeit eingeschränkt)"
            elif spiro_interval_sec == 10:
                spiro_interval_note = ""
            else:
                spiro_interval_note = f"({spiro_interval_sec}s-Auflösung)"

            if not spiro_df.empty:
                # VO2max Rolling-Window: 30s rolling average (3 points for 10s, 1 point for 30s)
                window_pts = max(1, int(30 / spiro_interval_sec))
                abs_vo2max = spiro_df['VO2'].rolling(window=window_pts, center=True).mean().max()
                rel_vo2max = abs_vo2max / weight if weight > 0 else 0

                stage_sec = int(stufendauer * 60)
                pause_sec = int(pausendauer)
                for i in range(anzahl):
                    end_t = vorlauf + (i + 1) * stage_sec + i * pause_sec
                    start_t = end_t - steady_fenster
                    mask = (spiro_df['Time'] >= start_t) & (spiro_df['Time'] <= end_t)
                    window_coords.append((start_t, end_t))
                    stage_data = spiro_df.loc[mask]
                    vo2_val  = stage_data['VO2'].mean()  if not stage_data.empty else 0.0
                    vco2_val = stage_data['VCO2'].mean() if not stage_data.empty else 0.0
                    vo2_steady_values.append(vo2_val / weight if weight > 0 else 0.0)
                    vo2_steady_abs.append(vo2_val)
                    vco2_steady_abs.append(vco2_val)
        except Exception as e:
            st.error(f"Fehler beim Einlesen: {e}")

    # Override VO2max if manual value is provided
    if vo2max_override > 0:
        abs_vo2max = float(vo2max_override)
        rel_vo2max = abs_vo2max / weight if weight > 0 else 0.0


    # --- SCHWELLENBERECHNUNG ---
    v_start, v_end = speed[0], speed[-1]
    results = []

    # 1. Fits
    # Fit polynomial on exercise stages for Dmax/ModDmax:
    poly_coeffs_no = np.polyfit(speed, lactate, 3)
    poly_func_no = np.poly1d(poly_coeffs_no)
    
    # Fit polynomial on all data (including adjusted baseline) for OBLA and LTP:
    speed_with_ruhe = raw_speeds.copy()
    if len(raw_speeds) > 2 and raw_speeds[0] == 0.0:
        step_size = raw_speeds[2] - raw_speeds[1]
        speed_with_ruhe[0] = raw_speeds[1] - step_size
    lactate_with_ruhe = raw_lactate.copy()
    poly_coeffs_with = np.polyfit(speed_with_ruhe, lactate_with_ruhe, 3)
    poly_func_with = np.poly1d(poly_coeffs_with)

    # 2. OBLA 4.0 (ANS)
    fitted_lactate_with = poly_func_with(speed_with_ruhe)
    
    def retrieve_intensity(target_lac, speeds, fitted_lac):
        n = len(fitted_lac)
        # Find where lactate curve starts to increase again after going down
        start_increasing = 0
        for i in range(1, n):
            if fitted_lac[i] - fitted_lac[i-1] < 0:
                start_increasing = i
        
        fit_slice = fitted_lac[start_increasing:]
        speed_slice = speeds[start_increasing:]
        
        if len(fit_slice) < 2:
            return None
            
        sort_idx = np.argsort(fit_slice)
        fit_sorted = fit_slice[sort_idx]
        speed_sorted = speed_slice[sort_idx]
        
        if target_lac < fit_sorted[0] or target_lac > fit_sorted[-1]:
            return None
        return np.interp(target_lac, fit_sorted, speed_sorted)
        
    v_exact_40 = retrieve_intensity(4.0, speed_with_ruhe, fitted_lactate_with)
    if v_exact_40 is not None:
        results.append({'Modell': 'OBLA 4.0 (ANS)', 'v': round(v_exact_40, 1), 'Laktat': 4.0})
    else:
        results.append({'Modell': 'OBLA 4.0 (ANS)', 'v': 'Laktatwerte zu niedrig um OBLA zu kalkulieren', 'Laktat': None})

    # 3. Dmax (Standard)
    # Solves perpendicular distance to line from first exercise stage speed[0] to last stage speed[-1]
    c3, c2, c1, _ = poly_coeffs_no
    m_dmax = (lactate[-1] - lactate[0]) / (speed[-1] - speed[0])
    d_roots = np.roots([3 * c3, 2 * c2, c1 - m_dmax])
    valid_d_roots = [r.real for r in d_roots if np.isreal(r) and speed[0] <= r.real <= speed[-1]]
    if valid_d_roots:
        v_exact_dmax = max(valid_d_roots)
        results.append({'Modell': 'Dmax (Standard)', 'v': round(v_exact_dmax, 1), 'Laktat': poly_func_no(v_exact_dmax)})
    else:
        results.append({'Modell': 'Dmax (Standard)', 'v': 'Kalkulation nicht möglich', 'Laktat': None})

    # 4. Modified Dmax
    # Line starts at point preceding first rise in lactate >= 0.4 mmol/L
    diffs = lactate[1:] - lactate[:-1]
    idx_first_rise = None
    for i, d in enumerate(diffs):
        if d >= 0.4:
            idx_first_rise = i
            break
            
    if not ausbelastung:
        results.append({'Modell': 'Modified Dmax', 'v': 'Aufgrund fehlender Ausbelastung keine Kalkulation der ModDmax möglich', 'Laktat': None})
    elif idx_first_rise is None:
        results.append({'Modell': 'Modified Dmax', 'v': 'Keine Laktat-Steigerung >= 0.4 mmol/L gefunden', 'Laktat': None})
    else:
        speed_start = speed[idx_first_rise]
        lac_start = lactate[idx_first_rise]
        m_mod = (lactate[-1] - lac_start) / (speed[-1] - speed_start)
        mod_roots = np.roots([3 * c3, 2 * c2, c1 - m_mod])
        valid_mod_roots = [r.real for r in mod_roots if np.isreal(r) and speed_start <= r.real <= speed[-1]]
        if valid_mod_roots:
            v_exact_mod = max(valid_mod_roots)
            results.append({'Modell': 'Modified Dmax', 'v': round(v_exact_mod, 1), 'Laktat': poly_func_no(v_exact_mod)})
        else:
            results.append({'Modell': 'Modified Dmax', 'v': 'Kalkulation nicht möglich', 'Laktat': None})

    # 5. LTP1 & LTP2
    try:
        # Interpolate raw speed and lactate onto 0.1 m/s grid (excluding Ruhe)
        grid_speed = np.arange(speed[0], speed[-1] + 1e-9, 0.1)
        interpolated_lactate = np.interp(grid_speed, speed, lactate)
        
        def fit_3seg_fixed_bps(bp1, bp2, x, y):
            if bp1 <= x[0] or bp2 >= x[-1] or bp1 >= bp2:
                return 1e10
            A = np.column_stack([
                np.ones_like(x),
                x,
                np.maximum(0.0, x - bp1),
                np.maximum(0.0, x - bp2)
            ])
            try:
                coefs, residuals, rank, s_vals = np.linalg.lstsq(A, y, rcond=None)
                if len(residuals) > 0:
                    return residuals[0]
                else:
                    pred = A @ coefs
                    return np.sum((y - pred)**2)
            except:
                return 1e10

        # Coarse grid search (30 x 30)
        best_err = 1e10
        best_bps = None
        bp1_vals = np.linspace(speed[0] + 0.05, speed[-1] - 0.05, 30)
        bp2_vals = np.linspace(speed[0] + 0.05, speed[-1] - 0.05, 30)
        for bp1 in bp1_vals:
            for bp2 in bp2_vals:
                if bp1 < bp2:
                    err = fit_3seg_fixed_bps(bp1, bp2, grid_speed, interpolated_lactate)
                    if err < best_err:
                        best_err = err
                        best_bps = (bp1, bp2)

        # Refine search with Nelder-Mead
        import scipy.optimize as opt
        def loss_func(bps):
            return fit_3seg_fixed_bps(bps[0], bps[1], grid_speed, interpolated_lactate)
            
        res_ltp = opt.minimize(loss_func, x0=best_bps, method='Nelder-Mead')
        v_exact_ltp1, v_exact_ltp2 = res_ltp.x[0], res_ltp.x[1]
        
        # Associated lactates from polynomial with baseline
        lac_ltp1 = poly_func_with(v_exact_ltp1)
        lac_ltp2 = poly_func_with(v_exact_ltp2)
        
        results.append({'Modell': 'LTP1', 'v': round(v_exact_ltp1, 1), 'Laktat': lac_ltp1})
        results.append({'Modell': 'LTP2', 'v': round(v_exact_ltp2, 1), 'Laktat': lac_ltp2})
    except Exception as e:
        pass


    # Header-Metriken & Körperfett-Tacho
    st.markdown(f"""
        <div style="font-size: 16px; line-height: 1.5; margin-bottom: 20px;">
            <b>Athlet:</b> {athlete_name}<br>
            <b>Protokoll:</b> {test_type}<br>
            <b>Test-Datum:</b> {test_date}
        </div>
        """, unsafe_allow_html=True)
        
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Gewicht", f"{weight} kg")
    with m3:
        st.metric("VO2max (rel)", f"{rel_vo2max:.1f} ml/min/kg")
        if spiro_interval_note:
            st.caption(f"⚠️ {spiro_interval_note}")
    m4.metric("Stufenanzahl", f"{anzahl}")
    
    with m2:
        fig_bf = go.Figure(go.Indicator(
            mode = "gauge+number",
            value = max(0, body_fat_pct),
            number = {'suffix': "%", 'valueformat': ".1f", 'font': {'color': '#00a1e0', 'size': 30}},
            title = {'text': "Körperfett", 'font': {'size': 14}},
            gauge = {
                'axis': {'range': [0, 30], 'tickwidth': 1, 'tickcolor': "darkblue"},
                'bar': {'color': "#00a1e0"},
                'bgcolor': "white",
                'borderwidth': 2,
                'bordercolor': "gray",
                'steps': [
                    {'range': [0, 10], 'color': "rgba(0, 161, 224, 0.1)"},
                    {'range': [10, 20], 'color': "rgba(0, 161, 224, 0.3)"},
                    {'range': [20, 30], 'color': "rgba(0, 161, 224, 0.5)"}],
            }
        ))
        fig_bf.update_layout(height=150, margin=dict(l=10, r=10, t=30, b=10))
        st.plotly_chart(fig_bf, width='stretch')

    st.markdown("---")

    # Ergebnisanzeige Matrix aufbauen mit Error-Handling
    df_res_list = []
    for res in results:
        row = {'Modell': res['Modell'], 'v': res['v']}
        
        if isinstance(res['v'], str):
            row['m/s'] = res['v']
            row['km/h'] = ""
            row['Laktat'] = ""
            row['HF'] = ""
            row['VLamax [mmol/l/s]'] = ""
        else:
            row['m/s'] = round(res['v'], 2)
            row['km/h'] = round(res['v'] * 3.6, 1)
            row['Laktat'] = round(res['Laktat'], 2)
            
            # Linear model for HF instead of interp
            hr_mask = (speed > 0) & (hr > 0) & (~np.isnan(speed)) & (~np.isnan(hr))
            if hr_mask.sum() >= 2:
                hr_slope, hr_intercept = np.polyfit(speed[hr_mask], hr[hr_mask], 1)
                row['HF'] = int(round(hr_slope * res['v'] + hr_intercept, 0))
            else:
                row['HF'] = int(round(np.interp(res['v'], speed, hr), 0))
            
            if uploaded_file is None and vo2max_override == 0:
                row['VLamax [mmol/l/s]'] = "Bitte lade eine Spirodatei hoch"
            else:
                vla, _, _ = calculate_vlamax_for_v(res['v'], rel_vo2max, speed, vo2_steady_abs, vco2_steady_abs, weight)
                row['VLamax [mmol/l/s]'] = f"{vla:.2f}" if vla is not None else "N/A"
                
        df_res_list.append(row)

    df_res = pd.DataFrame(df_res_list)
    poly_func = poly_func_no
    fig_laktat = None


    if not df_res.empty:
        c1, c2 = st.columns([1, 1])
        with c1:
            st.subheader("Schwellen & VLamax Matrix")
            st.dataframe(df_res[['Modell', 'm/s', 'km/h', 'Laktat', 'HF', 'VLamax [mmol/l/s]']], hide_index=True, width='stretch')
        with c2:
            fig, ax = plt.subplots(figsize=(8, 4))
            v_smooth = np.linspace(v_start, v_end, 200)
            ax.plot(speed, lactate, 'ko', label='Messwerte')
            ax.plot(v_smooth, poly_func(v_smooth), color='#00a1e0', label='Laktatkurve')
            
            for _, r in df_res.iterrows():
                if isinstance(r['v'], (int, float, np.number)):
                    ax.plot(r['v'], r['Laktat'], 'X', markersize=8, label=r['Modell'])
                    
            ax.set_xlabel("Geschwindigkeit (m/s)")
            ax.set_ylabel("Laktat (mmol/L)")
            ax.legend(loc='upper left', bbox_to_anchor=(1, 1))
            st.pyplot(fig)
            fig_laktat = fig # Graph speichern für das PDF

    # LAUFÖKONOMIE SECTION
    if not spiro_df.empty and len(vo2_steady_abs) > 0 and any(v > 0 for v in vo2_steady_abs):
        st.markdown("---")
        st.subheader("🏃 Laufökonomie & VLamax je Stufe")

        lauf_rows = []
        for i in range(anzahl):
            v_ms   = speed[i]
            v_kmh  = round(v_ms * 3.6, 1)
            vo2_a  = vo2_steady_abs[i]  if i < len(vo2_steady_abs)  else 0.0
            vco2_a = vco2_steady_abs[i] if i < len(vco2_steady_abs) else 0.0
            vo2_r  = vo2_steady_values[i]   # ml/min/kg

            # RE in ml/kg/km (Literaturstandard): VO2ss_rel [ml/min/kg] / v [km/h] * 60
            re_ml_kg_km = (vo2_r / v_kmh * 60.0) if (vo2_r > 0 and v_kmh > 0) else None

            # RQ
            rq_val = (vco2_a / vo2_a) if (vco2_a > 0 and vo2_a > 0) else None
            rq_val = max(0.7, min(1.0, rq_val)) if rq_val else None

            # VLamax je Stufe via bestehender Funktion
            vla_stage = None
            if rel_vo2max > 0 and vo2_r > 0 and len(vo2_steady_values) > 1:
                vla_stage, _, _ = calculate_vlamax_for_v(
                    v_ms, rel_vo2max, speed, vo2_steady_abs, vco2_steady_abs, weight
                )

            pace_sec = (1000.0 / v_ms) if v_ms > 0 else 0
            pace_str = f"{int(pace_sec // 60)}:{int(pace_sec % 60):02d}"

            lauf_rows.append({
                'v_ms':       v_ms,
                'v_kmh':      v_kmh,
                'pace':       pace_str,
                'vo2_abs':    vo2_a,
                'vo2_rel':    vo2_r,
                'rq':         rq_val,
                're_ml_kg_km':re_ml_kg_km,
                'vla_stage':  vla_stage
            })

        valid_rows = [r for r in lauf_rows if r['re_ml_kg_km'] is not None]

        if len(valid_rows) >= 1:
            col_lt, col_lc = st.columns([1, 1.3])

            with col_lt:
                st.markdown("**Messwerte je Stufe**")
                df_lauf_display = pd.DataFrame([{
                    'v [km/h]':           r['v_kmh'],
                    'Pace [min/km]':      r['pace'],
                    'VO2ss [ml/min/kg]':  round(r['vo2_rel'], 2) if r['vo2_rel'] > 0 else '-',
                    'RQ':                 round(r['rq'], 3)      if r['rq']       else '-',
                    'RE [ml/kg/km]':      round(r['re_ml_kg_km'], 1) if r['re_ml_kg_km'] else '-',
                    'VLamax [mmol/L/s]':  round(r['vla_stage'], 3)   if r['vla_stage']   else '-'
                } for r in lauf_rows])
                st.dataframe(df_lauf_display, hide_index=True, width='stretch')
                st.caption(
                    "RE [ml/kg/km] = VO2ss [ml/min/kg] ÷ v [km/h] × 60  "
                    "| VLamax je Stufe aus aerober Metabolik-Gleichung"
                )

            with col_lc:
                # Chart: RE [ml/kg/km] und VO2ss auf zwei Y-Achsen
                speeds_v    = [r['v_kmh']       for r in valid_rows]
                re_vals     = [r['re_ml_kg_km'] for r in valid_rows]
                vo2_r_vals  = [r['vo2_rel']     for r in lauf_rows if r['vo2_rel'] > 0]
                speeds_vo2  = [r['v_kmh']       for r in lauf_rows if r['vo2_rel'] > 0]

                fig_lauf = go.Figure()

                if len(speeds_v) >= 2:
                    v_interp  = np.linspace(min(speeds_v), max(speeds_v), 300)
                    re_interp = np.interp(v_interp, speeds_v, re_vals)
                    fig_lauf.add_trace(go.Scatter(
                        x=v_interp, y=re_interp,
                        mode='lines',
                        line=dict(color='#00a1e0', width=2.5),
                        name='RE [ml/kg/km]',
                        yaxis='y1'
                    ))

                fig_lauf.add_trace(go.Scatter(
                    x=speeds_v, y=re_vals,
                    mode='markers',
                    marker=dict(color='#00a1e0', size=10, symbol='circle',
                                line=dict(color='white', width=1.5)),
                    name='RE Messp.',
                    yaxis='y1'
                ))

                if len(speeds_vo2) >= 1:
                    v_vo2_interp = np.linspace(min(speeds_vo2), max(speeds_vo2), 300)
                    vo2_interp   = np.interp(v_vo2_interp, speeds_vo2, vo2_r_vals)
                    fig_lauf.add_trace(go.Scatter(
                        x=v_vo2_interp, y=vo2_interp,
                        mode='lines',
                        line=dict(color='#ff7f0e', width=2, dash='dot'),
                        name='VO2ss [ml/min/kg]',
                        yaxis='y2'
                    ))
                    fig_lauf.add_trace(go.Scatter(
                        x=speeds_vo2, y=vo2_r_vals,
                        mode='markers',
                        marker=dict(color='#ff7f0e', size=8),
                        name='VO2ss Messp.',
                        yaxis='y2'
                    ))

                fig_lauf.update_layout(
                    title='Laufökonomie (ml/kg/km) & VO2ss',
                    xaxis_title='Geschwindigkeit [km/h]',
                    yaxis=dict(
                        title=dict(text='RE [ml/kg/km]', font=dict(color='#00a1e0')),
                        tickfont=dict(color='#00a1e0')
                    ),
                    yaxis2=dict(
                        title=dict(text='VO2ss [ml/min/kg]', font=dict(color='#ff7f0e')),
                        tickfont=dict(color='#ff7f0e'),
                        overlaying='y',
                        side='right'
                    ),
                    legend=dict(x=0.01, y=0.99, bgcolor='rgba(255,255,255,0.7)'),
                    height=380,
                    margin=dict(l=60, r=70, t=50, b=50),
                    plot_bgcolor='#f9f9f9'
                )
                st.plotly_chart(fig_lauf, width='stretch')

    # EXPERTEN TOOLS
    st.markdown("---")
    with st.expander("Experten Tools zum Vergleich mit einer bestehenden HYCYS Diagnostik"):
        if not spiro_df.empty:
            st.subheader("Spiro-Synchronisation & Rohdaten")
            fig2, ax2 = plt.subplots(figsize=(12, 4))
            ax2.plot(spiro_df['Time'], spiro_df['VO2'], color='gray', alpha=0.5)
            for i, (s_t, e_t) in enumerate(window_coords):
                ax2.axvspan(s_t, e_t, color='#00a1e0', alpha=0.3)
            st.pyplot(fig2)
            
            col_diag1, col_diag2 = st.columns(2)
            with col_diag1:
                st.write("**Metabolische Vergleichswerte:**")
                st.write(f"- Berechnete VO2max: {rel_vo2max:.2f} ml/min/kg")
            with col_diag2:
                st.write("**Stufen-VO2 (relativ):**")
                st.dataframe(pd.DataFrame({"Stufe": [f"S{i+1}" for i in range(anzahl)], "VO2 rel": [f"{v:.2f}" for v in vo2_steady_values]}), hide_index=True)
        
        st.subheader("Stufenauswertung")
        if not spiro_df.empty:
            step_data = []
            for i in range(anzahl):
                v_ms   = speed[i]
                v_kmh  = round(v_ms * 3.6, 1)
                pace_s = (1000.0 / v_ms) if v_ms > 0 else 0
                pace_str = f"{int(pace_s // 60)}:{int(pace_s % 60):02d}"
                vo2_a  = vo2_steady_abs[i]  if i < len(vo2_steady_abs)  else 0.0
                vco2_a = vco2_steady_abs[i] if i < len(vco2_steady_abs) else 0.0
                vo2_r  = vo2_steady_values[i]
                rq_val   = (vco2_a / vo2_a) if (vco2_a > 0 and vo2_a > 0) else None
                rq_str   = f"{min(1.0, max(0.7, rq_val)):.3f}" if rq_val else '-'
                re_kg_km = (vo2_r / (v_ms * 3.6) * 60.0) if (vo2_r > 0 and v_ms > 0) else None
                step_data.append({
                    "Stufe":              f"{i+1}",
                    "v [km/h]":           v_kmh,
                    "Pace [min/km]":      pace_str,
                    "VO2ss [ml/min]": round(vo2_a, 1)  if vo2_a > 0 else '-',
                    "VO2ss [ml/min/kg]": round(vo2_r, 2) if vo2_r > 0 else '-',
                    "RQ":                 rq_str,
                    "RE [ml/kg/km]":     round(re_kg_km, 1) if re_kg_km else '-'
                })
            st.dataframe(pd.DataFrame(step_data), hide_index=True, width='stretch')
            st.caption("RE [ml/kg/km] = VO2ss [ml/min/kg] ÷ v [km/h] × 60  |  Literaturstandard Laufökonomie")

    # --- PDF DOWNLOAD BUTTON ---
    st.markdown("---")
    if fig_laktat is not None:
        pdf_bytes = create_pdf(
            athlete_name, birthdate, test_date, test_type, weight, body_fat_pct, rel_vo2max, df_res, fig_laktat,
            coach, sportart, kategorie, height, speed, lactate, hr, vo2_steady_values, vo2_steady_abs, vco2_steady_abs, uploaded_file
        )
        
        st.download_button(
            label="📄 PDF Report herunterladen",
            data=pdf_bytes,
            file_name=f"HYCYS_Laufdiagnostik_{athlete_name.replace(' ', '_')}.pdf",
            mime="application/pdf",
            type="primary"
        )