"""
Radsport PDF Report Generator
Pixel-accurate recreation of the HYCYS Cycling BLUE Report
"""
import os
import tempfile
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from matplotlib.patches import Patch
from matplotlib.lines import Line2D
from fpdf import FPDF


def create_pdf_rad(athlete_name, birthdate, test_date, test_type, weight, body_fat_pct, rel_vo2max,
                   ans_power, ans_rel, ans_ee_gesamt, ans_hr, ans_vo2_abs, ans_vo2_pct, ans_ee_kh_g,
                   fatmax_power, fatmax_rel, fatmax_ee_gesamt, fatmax_ee_fett,
                   map_val, carb_match_power, carb_intake,
                   vlamax, avg_rest_lac, max_post_lac, t_glyc, t_bel, t_alak,
                   coach, sportart, kategorie, height, slope_val, intercept_val, hfmax_val,
                   df_sim, sprint_times, sprint_powers, sprint_cadences,
                   avg_sprint_power, max_sprint_power, max_force, force_at_pmax, force_pct_max,
                   pot_vo2max, pot_fat, pot_ans_abs, pot_ans_rel, pot_fatmax, pot_pmax, pot_match, pot_vlamax):

    pdf = FPDF(unit='pt')
    pdf.set_auto_page_break(auto=False)

    # Brand colors
    c_teal      = (44, 183, 185)    # #2cb7b9
    c_gold      = (205, 182, 99)    # #cdb663
    c_dark_grey = (89, 90, 89)      # #595a59
    c_black     = (0, 0, 0)
    c_white     = (255, 255, 255)

    def draw_text(x, y, text, size=8.2, font_style='', color=c_black, font_name='Arial'):
        """Draw text at (x, y) where y is top of the text cell."""
        pdf.set_font(font_name, font_style, size)
        pdf.set_text_color(*color)
        safe = str(text).encode('latin-1', 'replace').decode('latin-1')
        pdf.text(x, y + size * 0.85, safe)

    def draw_text_centered(cx, y, text, size=8.2, font_style='', color=c_black):
        """Draw text centered at x=cx."""
        pdf.set_font('Arial', font_style, size)
        w = pdf.get_string_width(str(text).encode('latin-1', 'replace').decode('latin-1'))
        draw_text(cx - w / 2.0, y, text, size=size, font_style=font_style, color=color)

    def draw_rect(x, y, w, h, fill_color=None, stroke_color=None, lw=0.5):
        if fill_color:
            pdf.set_fill_color(*fill_color)
        if stroke_color:
            pdf.set_draw_color(*stroke_color)
            pdf.set_line_width(lw)
        style = ''
        if fill_color and stroke_color:
            style = 'FD'
        elif fill_color:
            style = 'F'
        elif stroke_color:
            style = 'D'
        if style:
            pdf.rect(x, y, w, h, style)

    def gold_box(bx, by, bw, bh, label, value, val_size=7.6):
        """Gold header box with centered label, value text below."""
        draw_rect(bx, by, bw, bh, fill_color=c_gold)
        draw_text_centered(bx + bw / 2.0, by + 1.5, label, size=8.2, font_style='B', color=c_white)
        draw_text(bx + 5, by + bh + 5.0, value, size=val_size, color=c_dark_grey)

    def page_logo(pg_num):
        """Draw small logo top-right (pages 2-5)."""
        asset = f'pdf_assets/page{pg_num}_img1.png'
        if not os.path.exists(asset):
            asset = 'pdf_assets/page2_img1.png'
        if os.path.exists(asset):
            pdf.image(asset, x=390.0, y=56.0, w=140, h=46)

    temp_files = []

    # =========================================================
    # SEITE 1 – DECKBLATT
    # =========================================================
    pdf.add_page()
    draw_rect(50.9, 54.5, 479.1, 663.6, fill_color=c_white)

    # Big HYCYS logo centred
    if os.path.exists('pdf_assets/page1_img3.png'):
        pdf.image('pdf_assets/page1_img3.png', x=97.9, y=194.4, w=382.2, h=118.0)

    # Gold title bar
    draw_rect(50.9, 404.6, 479.1, 20.8, fill_color=c_gold)
    draw_text(109.9, 407.6, "HYCYS Cycling BLUE", size=12.6, font_style='B', color=c_white)

    # Athlete info block — label left, value at fixed column
    label_x = 109.2
    value_x = 210.3
    info_y = [
        (446.6, "Name",          athlete_name),
        (456.9, "Geburtsdatum",  birthdate),
        (486.4, "Coach",         coach),
        (496.3, "Testdatum",     test_date),
        (515.9, "Sportart",      sportart),
        (525.8, "Kategorie",     kategorie),
    ]
    for y, lbl, val in info_y:
        draw_text(label_x, y, lbl,      size=7.6, color=c_black)
        draw_text(value_x, y, str(val), size=7.6, font_style='B', color=c_black)

    # Social icons
    if os.path.exists('pdf_assets/page1_img1.png'):
        pdf.image('pdf_assets/page1_img1.png', x=112.0, y=585.0, w=14.4, h=13.1)
    if os.path.exists('pdf_assets/page1_img2.png'):
        pdf.image('pdf_assets/page1_img2.png', x=134.9, y=584.8, w=14.6, h=12.0)
    draw_text(159.6, 588.0, "www.hycys.de",     size=7.0, font_style='B', color=c_teal)
    draw_text(260.7, 588.0, "contact@hycys.de", size=7.0, font_style='B', color=c_teal)

    # =========================================================
    # SEITE 2 – ANTHROPOMETRIE & LAKTATKINETIK
    # =========================================================
    pdf.add_page()
    draw_rect(50.9, 54.5, 479.1, 663.6, fill_color=c_white)
    page_logo(2)

    # --- Anthropologie Section ---
    draw_rect(57.6, 133.9, 171.9, 14.8, fill_color=c_teal)
    draw_text_centered(143.5, 135.8, "Anthropologie", size=8.2, font_style='B', color=c_white)

    lbl_x  = 58.3
    val_x  = 220.0
    draw_text(lbl_x, 160.0, "Gewicht",       size=7.6, font_style='B', color=c_black)
    draw_text(val_x, 160.0, f"{weight:.1f} kg".replace('.', ','), size=7.6, color=c_black)

    draw_text(lbl_x, 170.0, "Größe",          size=7.6, font_style='B', color=c_black)
    draw_text(val_x, 170.0, f"{int(height)} cm", size=7.6, color=c_black)

    bmi = weight / ((height / 100.0) ** 2) if height > 0 else 0.0
    draw_text(lbl_x, 190.0, "Body Mass Index", size=7.6, font_style='B', color=c_black)
    draw_text(val_x, 190.0, f"{bmi:.1f} kg/m²".replace('.', ','), size=7.6, color=c_black)

    # Masse / % column headers (small gold boxes)
    draw_rect(150.0, 207.0, 60.0, 12.0, fill_color=c_gold)
    draw_rect(216.0, 207.0, 40.0, 12.0, fill_color=c_gold)
    draw_text_centered(180.0, 208.5, "Masse", size=8.2, font_style='B', color=c_white)
    draw_text_centered(236.0, 208.5, "%",     size=8.2, font_style='B', color=c_white)

    fett_kg    = weight * body_fat_pct / 100.0
    fettfrei_kg = weight - fett_kg
    draw_text(lbl_x, 224.0, "Fett",     size=7.6, font_style='B', color=c_black)
    draw_text(155.0, 224.0, f"{fett_kg:.1f} kg".replace('.', ','),      size=7.6, color=c_black)
    draw_text(219.0, 224.0, f"{body_fat_pct:.1f} %".replace('.', ','),  size=7.6, color=c_black)

    draw_text(lbl_x, 234.0, "Fettfrei", size=7.6, font_style='B', color=c_black)
    draw_text(155.0, 234.0, f"{fettfrei_kg:.1f} kg".replace('.', ','),       size=7.6, color=c_black)
    draw_text(219.0, 234.0, f"{100.0 - body_fat_pct:.1f} %".replace('.', ','), size=7.6, color=c_black)

    # Pie chart (right half of page 2)
    fig, ax = plt.subplots(figsize=(3.0, 3.0), dpi=200)
    pct_fett = max(0.1, body_fat_pct)
    pct_frei = max(0.1, 100.0 - body_fat_pct)
    wedges, texts, autotexts = ax.pie(
        [pct_fett, pct_frei],
        colors=['#cdb663', '#2cb7b9'],
        autopct='%1.1f %%',
        startangle=90,
        pctdistance=0.60,
        textprops=dict(size=8, weight='bold', color='white'),
        wedgeprops=dict(edgecolor='white', linewidth=2),
    )
    for at in autotexts:
        at.set_color('white')
    ax.axis('equal')
    # Manual legend matching reference
    handles = [Patch(color='#cdb663', label='Fett'), Patch(color='#2cb7b9', label='Fettfrei')]
    ax.legend(handles=handles, loc='center right', bbox_to_anchor=(1.35, 0.5), fontsize=8, frameon=False)
    fig.patch.set_facecolor('white')
    plt.tight_layout(pad=0.3)
    with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
        fig.savefig(tmp.name, format='png', bbox_inches='tight', facecolor='white')
        plt.close(fig)
        pdf.image(tmp.name, x=270.0, y=130.0, w=245.0, h=120.0)
        temp_files.append(tmp.name)

    # --- VLamax Section ---
    draw_rect(57.6, 261.4, 458.1, 14.8, fill_color=c_teal)
    draw_text(63.0, 263.4, "Anaerober Stoffwechsel - maximale Laktatbildungsrate",
              size=8.2, font_style='B', color=c_white)

    draw_rect(140.0, 282.0, 95.0, 13.0, fill_color=c_gold)
    draw_text_centered(187.5, 283.8, "VLamax", size=8.2, font_style='B', color=c_white)
    draw_text_centered(187.5, 299.5, f"{vlamax:.2f} mmol/L/s".replace('.', ','), size=7.6, color=c_dark_grey)

    # --- VO2max Section ---
    draw_rect(57.6, 315.1, 458.1, 14.8, fill_color=c_teal)
    draw_text(57.6 + 5, 317.1, "Aerober Stoffwechsel - maximale Sauerstoffaufnahme",
              size=8.2, font_style='B', color=c_white)

    vo2_abs = rel_vo2max * weight
    ffm_kg  = weight * (100.0 - body_fat_pct) / 100.0
    vo2_ffm = vo2_abs / ffm_kg if ffm_kg > 0 else 0.0

    vo2_abs_str = f"{int(round(vo2_abs)):,}".replace(",", ".") + " ml/min"
    vo2_rel_str = f"{rel_vo2max:.1f} ml/min/kg".replace('.', ',')
    vo2_ffm_str = f"{vo2_ffm:.1f} ml/min/kg".replace('.', ',')
    map_str     = f"{int(round(map_val))} W"

    # 4 gold boxes for VO2max row
    gold_box( 57.6, 335.1, 108.6, 13.1, "VO2max abs.",     vo2_abs_str)
    gold_box(172.6, 335.1, 108.6, 13.1, "VO2max rel.",     vo2_rel_str)
    gold_box(287.6, 335.1, 108.6, 13.1, "VO2max rel. FFM", vo2_ffm_str)
    gold_box(402.6, 335.1, 113.1, 13.1, "max. Leistung",   map_str)

    # --- Laktat interaction plot ---
    draw_rect(57.6, 380.0, 458.1, 11.2, fill_color=c_teal)
    draw_text_centered(286.7, 381.8, "Interaktion Laktatauf- & abbau",
                       size=8.2, font_style='B', color=c_white)

    max_x = float(np.ceil((ans_power + 10.0) / 25.0) * 25.0)
    fig, ax = plt.subplots(figsize=(6.4, 3.6), dpi=200)
    ax.plot(df_sim['power'], df_sim['vo2ss'] * weight,
            color='#595a59', linewidth=1.5, label='Sauerstoffaufnahme - VO₂')
    ax.set_ylabel('VO₂ [ml/min]', color='#595a59', fontsize=7, fontweight='bold')
    ax.tick_params(axis='y', labelcolor='#595a59', labelsize=6)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(
        lambda v, _: f"{int(v):,}".replace(",", ".")))
    ax.set_ylim(500, 5500)
    ax.grid(True, linestyle=':', alpha=0.45, color='#cccccc')

    ax2 = ax.twinx()
    ax2.plot(df_sim['power'], df_sim['vlass'],    color='#2cb7b9', linewidth=1.5, label='Laktatproduktion')
    ax2.plot(df_sim['power'], df_sim['vlaoxmax'], color='#cdb663', linewidth=1.5, label='Laktatabbau')
    ax2.set_ylabel('Laktatkinetik [mmol/L/min]', color='#2cb7b9', fontsize=7, fontweight='bold')
    ax2.tick_params(axis='y', labelcolor='#2cb7b9', labelsize=6)
    ax2.set_ylim(0, 5.0)

    # Shade max Pyruvatdefizit zone
    try:
        if fatmax_power > 0:
            ax2.axvspan(fatmax_power - 17.5, fatmax_power + 17.5, color='#2cb7b9', alpha=0.2)
            ax2.text(fatmax_power, 2.5, "max.\nPyruvat-\ndefizit", color='white', fontsize=5.5, fontweight='bold',
                     ha='center', va='center', bbox=dict(facecolor='#2cb7b9', alpha=0.85, edgecolor='none', boxstyle='square,pad=0.3'))
    except Exception:
        pass

    ax.set_xlabel('Leistung [Watt]', fontsize=7, fontweight='bold')
    ax.set_xlim(0, max_x)
    ax.set_xticks(np.arange(0, max_x + 1, 25))
    ax.tick_params(axis='x', labelsize=6)

    lh, ll = ax.get_legend_handles_labels()
    lh2, ll2 = ax2.get_legend_handles_labels()
    ax.legend(lh + lh2, ll + ll2, loc='upper left', fontsize=5, framealpha=0.85)
    fig.patch.set_facecolor('white')
    plt.tight_layout()
    with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
        fig.savefig(tmp.name, format='png', bbox_inches='tight', facecolor='white')
        plt.close(fig)
        pdf.image(tmp.name, x=57.6, y=396.0, w=458.1, h=222.0)
        temp_files.append(tmp.name)

    # =========================================================
    # SEITE 3 – PHYSIOLOGISCHE LEISTUNGSFÄHIGKEIT
    # =========================================================
    pdf.add_page()
    draw_rect(50.9, 54.5, 479.1, 663.6, fill_color=c_white)
    page_logo(3)

    draw_rect(57.6, 127.8, 458.1, 14.8, fill_color=c_teal)
    draw_text_centered(286.7, 129.6, "physiologische Leistungsfähigkeit",
                       size=8.2, font_style='B', color=c_white)

    # ANS Section
    draw_rect(57.6, 148.0, 458.1, 11.2, fill_color=c_teal)
    draw_text_centered(286.7, 149.8, "Anaerobe Schwelle (ANS)",
                       size=8.2, font_style='B', color=c_white)

    ans_vo2_rel     = ans_vo2_abs / weight if weight > 0 else 0.0
    ans_abs_str     = f"{int(round(ans_power))} W"
    ans_rel_str     = f"{ans_rel:.2f} W/kg".replace('.', ',')
    ans_hr_str      = f"{int(round(ans_hr))} bpm"
    ans_vo2abs_str  = f"{int(round(ans_vo2_abs)):,}".replace(",", ".") + " ml/min"
    ans_vo2rel_str  = f"{ans_vo2_rel:.1f} ml/min/kg".replace('.', ',')
    ans_vo2pct_str  = f"{int(round(ans_vo2_pct))} %"
    ans_ee_ges_str  = f"{int(round(ans_ee_gesamt)):,}".replace(",", ".") + " kcal/h"
    ans_ee_kh_str   = f"{int(round(ans_ee_kh_g))} g/h"

    # Row 1
    gold_box( 57.6, 168.0, 143.6, 13.1, "ANS abs.",     ans_abs_str)
    gold_box(214.8, 168.0, 143.6, 13.1, "ANS rel.",     ans_rel_str)
    gold_box(372.0, 168.0, 143.7, 13.1, "Herzfrequenz", ans_hr_str)

    # Row 2
    gold_box( 57.6, 199.0, 143.6, 13.1, "VO2 abs. @ ANS", ans_vo2abs_str)
    gold_box(214.8, 199.0, 143.6, 13.1, "VO2 rel. @ ANS", ans_vo2rel_str)
    gold_box(372.0, 199.0, 143.7, 13.1, "% VO2max",       ans_vo2pct_str)

    # Row 3
    gold_box( 57.6, 231.1, 222.0, 13.1, "Energieverbrauch",       ans_ee_ges_str)
    gold_box(293.6, 231.1, 222.1, 13.1, "Kohlenhydrat-Verbrauch", ans_ee_kh_str)

    # Fatmax Section
    draw_rect(57.6, 271.7, 458.1, 11.2, fill_color=c_teal)
    draw_text_centered(286.7, 273.5, "Fettstoffwechsel - Fatmax",
                       size=8.2, font_style='B', color=c_white)

    fatmax_abs_str    = f"{int(round(fatmax_power))} W"
    fatmax_rel_fstr   = f"{fatmax_rel:.1f} W/kg".replace('.', ',')
    fatmax_ee_ges_str = f"{int(round(fatmax_ee_gesamt))} kcal/h"
    fatmax_ee_fet_str = f"{int(round(fatmax_ee_fett))} kcal/h"

    gold_box( 57.6, 294.7, 108.6, 13.1, "Fatmax abs.",  fatmax_abs_str)
    gold_box(172.6, 294.7, 108.6, 13.1, "Fatmax rel.",  fatmax_rel_fstr)
    gold_box(287.6, 294.7, 108.6, 13.1, "Energie",      fatmax_ee_ges_str)
    gold_box(402.6, 294.7, 113.1, 13.1, "Energie Fett", fatmax_ee_fet_str)

    # Fett & KH plot
    draw_rect(57.6, 335.3, 458.1, 11.2, fill_color=c_teal)
    draw_text_centered(286.7, 337.1, "Fett- & Kohlenhydrat-Stoffwechsel",
                       size=8.2, font_style='B', color=c_white)

    max_x = float(np.ceil((ans_power + 10.0) / 25.0) * 25.0)
    fig, ax = plt.subplots(figsize=(6.4, 3.6), dpi=200)
    ax.plot(df_sim['power'], df_sim['ee_gesamt'], color='#595a59', linewidth=1.5,
            label='Gesamtenergieverbrauch')
    ax.plot(df_sim['power'], df_sim['ee_fett'],   color='#2cb7b9', linewidth=1.5,
            label='Verbrauch von Fetten')
    ax.set_ylabel('Energieverbrauch [kcal/h]', color='#595a59', fontsize=7, fontweight='bold')
    ax.tick_params(axis='y', labelcolor='#595a59', labelsize=6)
    ax.set_ylim(0, 1500)
    ax.grid(True, linestyle=':', alpha=0.45, color='#cccccc')

    ax2 = ax.twinx()
    ax2.plot(df_sim['power'], df_sim['ee_kh_g'], color='#cdb663', linewidth=1.5,
             label='Verbrauch von Kohlenhydraten')
    ax2.set_ylabel('Kohlenhydrat-Verbrauch [g/h]', color='#cdb663', fontsize=7, fontweight='bold')
    ax2.tick_params(axis='y', labelcolor='#cdb663', labelsize=6)
    ax2.set_ylim(0, 350)

    # Fatmax zone indicator
    try:
        if fatmax_power > 0:
            ax.axvspan(fatmax_power - 17.5, fatmax_power + 17.5, color='#2cb7b9', alpha=0.25)
            ax.text(fatmax_power, 750, "Fatmax", color='white', fontsize=6, fontweight='bold',
                    ha='center', va='center', bbox=dict(facecolor='#2cb7b9', alpha=0.85, edgecolor='none', boxstyle='square,pad=0.3'))
    except Exception:
        pass

    ax.set_xlabel('Leistung [Watt]', fontsize=7, fontweight='bold')
    ax.set_xlim(0, max_x)
    ax.set_xticks(np.arange(0, max_x + 1, 25))
    ax.tick_params(axis='x', labelsize=6)

    lh, ll = ax.get_legend_handles_labels()
    lh2, ll2 = ax2.get_legend_handles_labels()
    ax.legend(lh + lh2, ll + ll2, loc='upper left', fontsize=5, framealpha=0.85)
    fig.patch.set_facecolor('white')
    plt.tight_layout()
    with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
        fig.savefig(tmp.name, format='png', bbox_inches='tight', facecolor='white')
        plt.close(fig)
        pdf.image(tmp.name, x=57.6, y=352.0, w=458.1, h=218.0)
        temp_files.append(tmp.name)

    # =========================================================
    # SEITE 4 – REFERENZDATEN & TRAININGSBEREICHE
    # =========================================================
    pdf.add_page()
    draw_rect(50.9, 54.5, 479.1, 663.6, fill_color=c_white)
    page_logo(4)

    draw_rect(57.6, 127.8, 155.3, 13.4, fill_color=c_teal)
    draw_text_centered(135.3, 129.6, "Referenzdaten", size=8.2, font_style='B', color=c_white)

    def normalize_ref(val, cat):
        ref_params = {
            'ANS abs. [W]':             (325.0, 40.0,  False),
            'Fettanteil [%]':           (13.0,  2.5,   True),
            'VO2max [ml/min/kg]':       (67.0,  4.5,   False),
            'VLamax [mmol/L/s]':        (0.6,   0.1,   True),
            'KH-Match [W/kg]':          (3.3,   0.3,   False),
            'Maximale Leistung [W/kg]': (15.0,  2.0,   False),
            'Fatmax [W/kg]':            (3.1,   0.3,   False),
            'ANS rel. [W/kg]':          (4.2,   0.45,  False),
        }
        mean, sd, inverted = ref_params[cat]
        span = 6 * sd
        if span == 0:
            return 0.0
        rel = (val - (mean - 3 * sd)) / span * 100.0
        if inverted:
            rel = 100.0 - rel
        return max(0.0, min(100.0, rel))

    radar_cats = ['ANS abs. [W]', 'Fettanteil [%]', 'VO2max [ml/min/kg]', 'VLamax [mmol/L/s]',
                  'KH-Match [W/kg]', 'Maximale Leistung [W/kg]', 'Fatmax [W/kg]', 'ANS rel. [W/kg]']
    radar_labels = ['ANS abs. [W]', 'Fettanteil [%]', 'VO2max [ml/min/kg]', 'VLamax [mmol/L/s]',
                    'KH-Match [W/kg]', 'Maximale Leistung [W/kg]', 'Fatmax [W/kg]', 'ANS rel. [W/kg]']
    N      = len(radar_cats)
    angles = [n / float(N) * 2 * np.pi for n in range(N)]
    angles += angles[:1]

    fig = plt.figure(figsize=(5.5, 5.2), dpi=200)
    ax  = fig.add_subplot(111, polar=True)

    # Background rings  (Sehr gut > Gut > Zu verbessern > Schwach)
    ax.fill(angles, [100] * len(angles), color='#cdb663', alpha=0.40, zorder=1, label='Sehr gut')
    ax.fill(angles, [75]  * len(angles), color='#2cb7b9', alpha=0.40, zorder=2, label='Gut')
    ax.fill(angles, [50]  * len(angles), color='#e0e0e0', alpha=0.65, zorder=3, label='Zu verbessern')
    ax.fill(angles, [25]  * len(angles), color='#c0c0c0', alpha=0.80, zorder=4, label='Schwach')

    ath_pmax_rel  = max_sprint_power / weight if weight > 0 else 0.0
    ath_match_rel = carb_match_power  / weight if weight > 0 else 0.0

    athlete_scores = [
        normalize_ref(ans_power,     'ANS abs. [W]'),
        normalize_ref(body_fat_pct,  'Fettanteil [%]'),
        normalize_ref(rel_vo2max,    'VO2max [ml/min/kg]'),
        normalize_ref(vlamax,        'VLamax [mmol/L/s]'),
        normalize_ref(ath_match_rel, 'KH-Match [W/kg]'),
        normalize_ref(ath_pmax_rel,  'Maximale Leistung [W/kg]'),
        normalize_ref(fatmax_rel,    'Fatmax [W/kg]'),
        normalize_ref(ans_rel,       'ANS rel. [W/kg]'),
    ]
    athlete_scores += athlete_scores[:1]

    potential_scores = [
        normalize_ref(pot_ans_abs, 'ANS abs. [W]'),
        normalize_ref(pot_fat,     'Fettanteil [%]'),
        normalize_ref(pot_vo2max,  'VO2max [ml/min/kg]'),
        normalize_ref(pot_vlamax,  'VLamax [mmol/L/s]'),
        normalize_ref(pot_match,   'KH-Match [W/kg]'),
        normalize_ref(pot_pmax,    'Maximale Leistung [W/kg]'),
        normalize_ref(pot_fatmax,  'Fatmax [W/kg]'),
        normalize_ref(pot_ans_rel, 'ANS rel. [W/kg]'),
    ]
    potential_scores += potential_scores[:1]

    ax.plot(angles, athlete_scores,   color='#1a1a1a', linewidth=2.0, marker='o', markersize=5, zorder=6)
    ax.fill(angles, athlete_scores,   color='#1a1a1a', alpha=0.08, zorder=5)
    ax.plot(angles, potential_scores, color='#595a59', linewidth=1.5, linestyle='--',
            marker='s', markersize=4, zorder=6)

    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(radar_labels, fontsize=6.0, fontweight='bold', color='#333333')
    ax.set_yticklabels([])
    ax.set_ylim(0, 100)
    ax.grid(True, color='gray', linestyle=':', linewidth=0.5, alpha=0.55)
    ax.spines['polar'].set_visible(False)

    legend_handles = [
        Patch(facecolor='#cdb663', alpha=0.6,  label='Sehr gut'),
        Patch(facecolor='#2cb7b9', alpha=0.6,  label='Gut'),
        Patch(facecolor='#e0e0e0', alpha=0.8,  label='Zu verbessern'),
        Patch(facecolor='#c0c0c0', alpha=0.9,  label='Schwach'),
        Line2D([0], [0], color='#1a1a1a', lw=2, marker='o', ms=5, label='Daten Sportler'),
        Line2D([0], [0], color='#595a59', lw=1.5, ls='--', marker='s', ms=4, label='Coaching Potential'),
    ]
    ax.legend(handles=legend_handles, loc='lower center', bbox_to_anchor=(0.5, -0.24),
              ncol=3, fontsize=6, frameon=True, framealpha=0.9)
    fig.patch.set_facecolor('white')
    plt.tight_layout()
    with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
        fig.savefig(tmp.name, format='png', bbox_inches='tight', facecolor='white')
        plt.close(fig)
        pdf.image(tmp.name, x=57.6, y=143.0, w=458.1, h=265.0)
        temp_files.append(tmp.name)

    # Training zones table
    draw_rect(57.6, 414.9, 155.3, 13.4, fill_color=c_teal)
    draw_text_centered(135.3, 416.8, "Trainingsbereiche", size=8.2, font_style='B', color=c_white)

    # Table header
    draw_text(99.1,  435.7, "Trainingsbereich", size=8.2, font_style='B', color=c_black)
    draw_text(256.9, 435.7, "Leistung",         size=8.2, font_style='B', color=c_black)
    draw_text(382.8, 435.7, "Herzfrequenz",     size=8.2, font_style='B', color=c_black)
    draw_text(481.9, 435.7, "Tretfreq.",        size=8.2, font_style='B', color=c_black)
    draw_text(263.8, 446.1, "[Watt]",           size=7.6, color=c_black)
    draw_text(399.1, 446.1, "[bpm]",            size=7.6, color=c_black)
    draw_text(486.3, 446.1, "[U/min]",          size=7.6, color=c_black)
    draw_text(222.4, 456.6, "min",  size=7.6, color=c_black)
    draw_text(266.5, 456.6, "max",  size=7.6, color=c_black)
    draw_text(312.3, 456.6, "Ziel", size=7.6, color=c_black)
    draw_text(357.1, 456.6, "min",  size=7.6, color=c_black)
    draw_text(401.1, 456.6, "max",  size=7.6, color=c_black)
    draw_text(447.0, 456.6, "Ziel", size=7.6, color=c_black)

    def hr_at_p(p):
        return int(round(min(hfmax_val, max(40, p * slope_val + intercept_val))))

    # Zone boundaries
    kb_max  = ans_power * 0.5
    g1_min  = kb_max
    g1_max  = ans_power * 0.75
    try:
        g1_i  = (df_sim['ee_kh_g'] - 0.6 * weight).abs().idxmin()
        g1_zl = int(round(float(df_sim.loc[g1_i, 'power']) / 10) * 10)
    except Exception:
        g1_zl = int(round(g1_max * 0.93 / 10) * 10)
    g2_min  = g1_max
    g2_max  = ans_power * 0.90
    eb_min  = g2_max
    eb_max  = ans_power * 1.05
    sb_min  = eb_max
    sb_max  = map_val * 1.10
    k3_min  = ans_power * 0.84
    k3_max  = ans_power * 0.98

    def rnd(v): return int(round(v / 10) * 10)

    zones = [
        ("Kompensationsbereich (KB)", "-",         rnd(kb_max), rnd(kb_max*0.9),
         "-",              hr_at_p(kb_max),  hr_at_p(kb_max*0.9),  "80-100"),
        ("Grundlagenausdauer 1 (G1)", rnd(g1_min), rnd(g1_max), g1_zl,
         hr_at_p(g1_min), hr_at_p(g1_max),  hr_at_p(g1_zl),       "85-110"),
        ("Grundlagenausdauer 2 (G2)", rnd(g2_min), rnd(g2_max), rnd(ans_power*0.86),
         hr_at_p(g2_min), hr_at_p(g2_max),  hr_at_p(ans_power*0.86), "85-110"),
        ("Entwicklungsbereich (EB)",  rnd(eb_min), rnd(eb_max), rnd(ans_power),
         hr_at_p(eb_min), hr_at_p(eb_max),  hr_at_p(ans_power),   "85-110"),
        ("Spitzenbereich (SB)",       rnd(sb_min), rnd(sb_max), rnd(map_val*1.075),
         hr_at_p(sb_min), "max",             hr_at_p(map_val*1.075), "95-120"),
        ("Kraftausdauer (K3)",        rnd(k3_min), rnd(k3_max), rnd(ans_power*0.92),
         hr_at_p(k3_min), hr_at_p(k3_max),  hr_at_p(ans_power*0.92), "40-60"),
    ]

    y_row = 466.2
    for name, p_min, p_max, p_zl, hr_mn, hr_mx, hr_zl, tf in zones:
        draw_text(58.7,  y_row, name,       size=7.6, font_style='B', color=c_black)
        draw_text(222.4, y_row, str(p_min), size=7.6, color=c_black)
        draw_text(266.5, y_row, str(p_max), size=7.6, color=c_black)
        draw_text(312.3, y_row, str(p_zl),  size=7.6, font_style='B', color=c_black)
        draw_text(357.1, y_row, str(hr_mn), size=7.6, color=c_black)
        draw_text(401.1, y_row, str(hr_mx), size=7.6, color=c_black)
        draw_text(447.0, y_row, str(hr_zl), size=7.6, font_style='B', color=c_black)
        draw_text(486.7, y_row, tf,          size=7.6, color=c_black)
        y_row += 9.9

    # IE / LC sub-table
    draw_text(279.3, 535.1, "Leistung",  size=8.2, font_style='B', color=c_black)
    draw_text(481.9, 535.1, "Tretfreq.", size=8.2, font_style='B', color=c_black)
    draw_text(286.2, 545.9, "[Watt]",   size=7.6, color=c_black)
    draw_text(486.3, 545.9, "[U/min]",  size=7.6, color=c_black)
    draw_text(259.4, 555.8, "Intervall", size=7.6, color=c_black)
    draw_text(307.8, 555.8, "Pause",    size=7.6, color=c_black)

    ie_int   = rnd(map_val * 1.05)
    ie_pause = rnd(ans_power * 0.5)
    draw_text(58.8,  565.6, "Intermitted exercise (IE)", size=8.2, font_style='B', color=c_black)
    draw_text(267.6, 565.6, str(ie_int),   size=7.6, color=c_black)
    draw_text(312.5, 565.6, str(ie_pause), size=7.6, color=c_black)
    draw_text(486.7, 565.6, "85-110",      size=7.6, color=c_black)

    try:
        lc_sim = df_sim.copy()
        lc_above = lc_sim[lc_sim['vlass'] > lc_sim['vlaoxmax']]
        if not lc_above.empty:
            lc_above = lc_above.copy()
            lc_above['d1'] = (lc_above['vlass'] - lc_above['vlaoxmax'] - 1.0).abs()
            lc_int = rnd(float(lc_above.loc[lc_above['d1'].idxmin(), 'power']))
        else:
            lc_int = 0
    except Exception:
        lc_int = 0
    lc_pause = rnd(fatmax_power)
    draw_text(58.8,  575.9, "Laktat Auf-/Abbau (LC)", size=8.2, font_style='B', color=c_black)
    draw_text(267.6, 575.9, str(lc_int),   size=7.6, color=c_black)
    draw_text(312.5, 575.9, str(lc_pause), size=7.6, color=c_black)
    draw_text(486.7, 575.9, "85-110",      size=7.6, color=c_black)

    # =========================================================
    # SEITE 5 – SPRINT & ANTRITT
    # =========================================================
    pdf.add_page()
    draw_rect(50.9, 54.5, 479.1, 663.6, fill_color=c_white)
    page_logo(5)

    # Large logo centred (as in reference page 5)
    if os.path.exists('pdf_assets/page1_img3.png'):
        pdf.image('pdf_assets/page1_img3.png', x=112.3, y=155.0, w=362.5, h=112.0)

    # Sprint & Antritt header bar
    draw_rect(57.6, 414.9, 458.1, 14.8, fill_color=c_teal)
    draw_text_centered(286.7, 416.8, "Sprint & Antritt", size=8.2, font_style='B', color=c_white)

    # Leistungswerte section
    draw_rect(57.6, 452.0, 143.6, 13.1, fill_color=c_gold)
    draw_text_centered(129.4, 454.4, "Leistungswerte", size=8.2, font_style='B', color=c_white)

    avg_p_str     = f"{int(round(avg_sprint_power))} W"
    max_p_str     = f"{int(round(max_sprint_power))} W"
    max_p_rel_str = (f"{max_sprint_power / weight:.1f} W/kg".replace('.', ',')
                     if weight > 0 else "0,0 W/kg")

    draw_text(58.7, 466.2, "durchschn. Leistung", size=7.6, font_style='B', color=c_black)
    draw_text(172.8, 466.7, avg_p_str,             size=7.6, color=c_black)
    draw_text(58.7, 476.1, "max. Leistung abs.",  size=7.6, font_style='B', color=c_black)
    draw_text(170.7, 476.6, max_p_str,             size=7.6, color=c_black)
    draw_text(58.7, 485.9, "max. Leistung rel.",  size=7.6, font_style='B', color=c_black)
    draw_text(166.6, 486.4, max_p_rel_str,         size=7.6, color=c_black)

    # Beschleunigung & Kraft section
    draw_rect(57.6, 531.2, 143.6, 13.1, fill_color=c_gold)
    draw_text_centered(129.4, 533.6, "Beschleunigung & Kraft", size=8.2, font_style='B', color=c_white)

    max_f_str     = f"{int(round(max_force))} N"
    f_at_pmax_str = f"{int(round(force_at_pmax))} N"
    f_pct_str     = f"{int(round(force_pct_max))} %"

    draw_text(58.7, 545.5, "max. Kraft",              size=7.6, font_style='B', color=c_black)
    draw_text(173.8, 545.9, max_f_str,                 size=7.6, color=c_black)
    draw_text(58.7, 555.3, "Kraft bei max. Leistung", size=7.6, font_style='B', color=c_black)
    draw_text(173.8, 555.8, f_at_pmax_str,             size=7.6, color=c_black)
    draw_text(58.7, 565.4, "% der max. Kraft",        size=7.6, font_style='B', color=c_black)
    draw_text(176.3, 565.9, f_pct_str,                 size=7.6, color=c_black)

    # Sprint Power vs Time chart
    fig, ax = plt.subplots(figsize=(3.9, 2.5), dpi=200)
    ax.plot(sprint_times, sprint_powers, color='#595a59', linewidth=1.5)
    try:
        if t_alak > 0 and t_bel > t_alak:
            ax.axvspan(0.0,    t_alak, color=(0.0, 0.631, 0.878), alpha=0.18)
            ax.axvspan(t_alak, t_bel,  color=(0.804, 0.714, 0.388), alpha=0.18)
    except Exception:
        pass

    ax.set_ylabel('Leistung [W]', color='#595a59', fontsize=6, fontweight='bold')
    ax.tick_params(axis='y', labelcolor='#595a59', labelsize=5)
    ax.set_xlabel('Zeit [s]', color='#595a59', fontsize=6, fontweight='bold')
    ax.tick_params(axis='x', labelsize=5)
    ax.grid(True, linestyle=':', alpha=0.45, color='#cccccc')
    fig.patch.set_facecolor('white')
    plt.tight_layout()
    with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
        fig.savefig(tmp.name, format='png', bbox_inches='tight', facecolor='white')
        plt.close(fig)
        # Chart placed fully to the right of the value boxes, below the header bar
        pdf.image(tmp.name, x=220.0, y=435.0, w=288.0, h=162.0)
        temp_files.append(tmp.name)

    # Output bytes
    pdf_output = pdf.output(dest='S')
    for f in temp_files:
        try:
            os.unlink(f)
        except Exception:
            pass
    return pdf_output.encode('latin-1', 'replace')
