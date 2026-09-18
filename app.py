from flask import Flask, render_template, request, redirect, url_for, session, send_file, flash
from functools import wraps
from datetime import datetime
import sqlite3
import os
import io

# Word generovanie
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

# PDF generovanie
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Font s podporou slovenskej diakritiky (Liberation Serif - metricky kompatibilný s Times New Roman)
FONT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'fonts')
PDF_FONT = 'Times-Roman'
PDF_FONT_BOLD = 'Times-Bold'
try:
    pdfmetrics.registerFont(TTFont('PDFSerif', os.path.join(FONT_DIR, 'LiberationSerif-Regular.ttf')))
    pdfmetrics.registerFont(TTFont('PDFSerif-Bold', os.path.join(FONT_DIR, 'LiberationSerif-Bold.ttf')))
    pdfmetrics.registerFontFamily('PDFSerif', normal='PDFSerif', bold='PDFSerif-Bold')
    PDF_FONT = 'PDFSerif'
    PDF_FONT_BOLD = 'PDFSerif-Bold'
except Exception:
    pass  # ak font chýba, použije sa vstavaný Times-Roman (bez slovenskej diakritiky)

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'zmen-toto-heslo-pred-nasadenim')

# ============================================================
# PRIHLASENIE - zmen USERNAME a PASSWORD
# ============================================================
USERNAME = os.environ.get('APP_USERNAME', 'admin')
PASSWORD = os.environ.get('APP_PASSWORD', 'heslo123')

# ============================================================
# DATABAZA
# ============================================================
DB_PATH = 'audity.db'

def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS audity (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        typ TEXT NOT NULL,
        datum TEXT NOT NULL,
        nazov_budovy TEXT,
        adresa TEXT,
        vlastnik TEXT,
        udaje TEXT,
        vysledky TEXT
    )''')
    conn.commit()
    conn.close()

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

# ============================================================
# VZORCE - ENERGETICKY AUDIT
# ============================================================
def vypocitaj_audit(data):
    try:
        plocha = float(data.get('plocha', 0))
        spotreba_teplo = float(data.get('spotreba_teplo', 0))
        spotreba_elektrina = float(data.get('spotreba_elektrina', 0))
        spotreba_celkova = spotreba_teplo + spotreba_elektrina

        # Merná spotreba energie
        if plocha > 0:
            merna_spotreba = spotreba_celkova / plocha
        else:
            merna_spotreba = 0

        # Energetická trieda
        if merna_spotreba < 30:
            trieda = "A0"
        elif merna_spotreba < 60:
            trieda = "A1"
        elif merna_spotreba < 100:
            trieda = "B"
        elif merna_spotreba < 150:
            trieda = "C"
        elif merna_spotreba < 200:
            trieda = "D"
        elif merna_spotreba < 250:
            trieda = "E"
        else:
            trieda = "F"

        return {
            'merna_spotreba': round(merna_spotreba, 2),
            'spotreba_celkova': round(spotreba_celkova, 2),
            'trieda': trieda
        }
    except:
        return {'merna_spotreba': 0, 'spotreba_celkova': 0, 'trieda': 'N/A'}

# ============================================================
# VZORCE - ENERGETICKY CERTIFIKAT
# ============================================================
def vypocitaj_certifikat(data):
    try:
        plocha = float(data.get('plocha', 0))
        potreba_tepla = float(data.get('potreba_tepla', 0))
        potreba_chlad = float(data.get('potreba_chlad', 0))
        potreba_svietenie = float(data.get('potreba_svietenie', 0))

        celkova_potreba = potreba_tepla + potreba_chlad + potreba_svietenie

        if plocha > 0:
            merna_potreba = celkova_potreba / plocha
        else:
            merna_potreba = 0

        # Energetická trieda certifikátu (podľa vyhlášky 364/2012)
        if merna_potreba < 32:
            trieda = "A0"
        elif merna_potreba < 54:
            trieda = "A1"
        elif merna_potreba < 95:
            trieda = "B"
        elif merna_potreba < 150:
            trieda = "C"
        elif merna_potreba < 210:
            trieda = "D"
        elif merna_potreba < 270:
            trieda = "E"
        else:
            trieda = "F"

        return {
            'celkova_potreba': round(celkova_potreba, 2),
            'merna_potreba': round(merna_potreba, 2),
            'trieda': trieda
        }
    except:
        return {'celkova_potreba': 0, 'merna_potreba': 0, 'trieda': 'N/A'}

# ============================================================
# VZORCE - KONTROLA VYKUROVANIA
# podľa Zákona č. 314/2012 Z.z. a Vyhl. č. 422/2012 Z.z.
# (pravidelná kontrola kotla a vykurovacieho systému)
# ============================================================

# Siegertove koeficienty komínovej straty pre jednotlivé druhy paliva
KOEFICIENTY_PALIVA = {
    'Zemný plyn':                {'A2': 0.65, 'B': 0.008, 'vyhrevnost': 9.86},
    'Vykurovací olej':           {'A2': 0.68, 'B': 0.007, 'vyhrevnost': 11.86},
    'Tuhé palivo - čierne uhlie':{'A2': 0.60, 'B': 0.010, 'vyhrevnost': 7.00},
    'Tuhé palivo - drevo/biomasa':{'A2': 0.65, 'B': 0.025, 'vyhrevnost': 4.00},
}

STRATA_SALANIM = 0.8  # % - typická strata sálaním kotla do okolia (odhad)

def komin_strata(t_spalin, t_vzduch, o2, palivo):
    """Komínová strata (Siegertova metóda) v %."""
    k = KOEFICIENTY_PALIVA.get(palivo, KOEFICIENTY_PALIVA['Zemný plyn'])
    if o2 >= 20.9 or t_spalin <= t_vzduch:
        return 0.0
    return (t_spalin - t_vzduch) * (k['A2'] / (20.9 - o2) + k['B'])

def _f(data, key, default=0):
    try:
        v = data.get(key)
        return float(v) if v not in (None, '') else default
    except (TypeError, ValueError):
        return default

def vypocitaj_vykurovanie(data):
    try:
        min_ucinnost = _f(data, 'min_ucinnost', 96)

        # --- 1) Meranie účinnosti kotlov pri 50 % a 100 % zaťažení (komínová strata) ---
        kotly = []
        for pref in ['k1', 'k2']:
            if not data.get(f'{pref}_menovity_vykon'):
                continue
            palivo = data.get(f'{pref}_typ_paliva') or 'Zemný plyn'
            garantovana = _f(data, f'{pref}_garantovana_ucinnost')

            merania = []
            for bod, label in [('50', '50 %'), ('100', '100 %')]:
                t_spalin = _f(data, f'{pref}_t_spalin_{bod}')
                t_vzduch = _f(data, f'{pref}_t_vzduch_{bod}')
                o2 = _f(data, f'{pref}_o2_{bod}')
                if t_spalin == 0 and t_vzduch == 0 and o2 == 0:
                    continue
                qA = round(komin_strata(t_spalin, t_vzduch, o2, palivo), 2)
                ucinnost_bod = round(100 - qA - STRATA_SALANIM, 2)
                merania.append({
                    'zatazenie': label, 't_spalin': t_spalin, 't_vzduch': t_vzduch,
                    'o2': o2, 'komin_strata': qA, 'ucinnost': ucinnost_bod,
                })

            if merania:
                priemer = round(sum(m['ucinnost'] for m in merania) / len(merania), 2)
            else:
                priemer = 0
            stav = 'Vyhovuje' if priemer >= min_ucinnost else 'Nevyhovuje'

            kotly.append({
                'oznacenie': pref.upper(),
                'vyrobca': data.get(f'{pref}_vyrobca', ''),
                'typ': data.get(f'{pref}_typ', ''),
                'vyrobne_cislo': data.get(f'{pref}_vyrobne_cislo', ''),
                'rok_vyroby': data.get(f'{pref}_rok_vyroby', ''),
                'prevadzkovy_stav': data.get(f'{pref}_prevadzkovy_stav') or 'v prevádzke',
                'menovity_vykon': data.get(f'{pref}_menovity_vykon', ''),
                'max_vykon': data.get(f'{pref}_max_vykon', ''),
                'max_prikon': data.get(f'{pref}_max_prikon', ''),
                'min_vykon': data.get(f'{pref}_min_vykon', ''),
                'min_prikon': data.get(f'{pref}_min_prikon', ''),
                'kondenzacny': data.get(f'{pref}_kondenzacny') or 'kondenzačný',
                'oznacenie_ce': data.get(f'{pref}_oznacenie_ce', ''),
                'typ_regulacie': data.get(f'{pref}_typ_regulacie', ''),
                'sposob_odvodu_spalin': data.get(f'{pref}_sposob_odvodu_spalin', ''),
                'sposob_privodu_vzduchu': data.get(f'{pref}_sposob_privodu_vzduchu', ''),
                'teplonosne_medium': data.get(f'{pref}_teplonosne_medium') or 'Teplá voda',
                'sposob_vyuzitia': data.get(f'{pref}_sposob_vyuzitia') or 'Ústredné vykurovanie a teplá voda',
                'typ_paliva': palivo,
                'sposob_davkovania': data.get(f'{pref}_sposob_davkovania') or 'automatické',
                'merania': merania,
                'priemerna_ucinnost': priemer,
                'garantovana_ucinnost': garantovana,
                'stav': stav,
            })

        # --- 2) Priama metóda — až 3 roky (rovnako ako v protokole) ---
        roky = []
        for i in range(1, 4):
            rok = data.get(f'rok{i}')
            odber = _f(data, f'rok{i}_odber_paliva')
            if not rok or odber == 0:
                continue
            vyhrevnost = _f(data, f'rok{i}_vyhrevnost')
            vyrobene_teplo = _f(data, f'rok{i}_vyrobene_teplo')
            spotreba_vykurovanie = _f(data, f'rok{i}_spotreba_vykurovanie')
            spotreba_tuv = _f(data, f'rok{i}_spotreba_tuv')
            spotreba_vody_tuv = _f(data, f'rok{i}_spotreba_vody_tuv')

            teplo_v_palive = round(odber * vyhrevnost, 2)
            ucinnost_priama = round((vyrobene_teplo / teplo_v_palive) * 100, 2) if teplo_v_palive > 0 else 0
            merna_spotreba_tuv = round(spotreba_tuv / spotreba_vody_tuv, 2) if spotreba_vody_tuv > 0 else 0
            celkova_spotreba = spotreba_vykurovanie + spotreba_tuv
            podiel_tuv = round((spotreba_tuv / celkova_spotreba) * 100, 2) if celkova_spotreba > 0 else 0

            # Posúdenie potrebného výkonu (STN 06 0210 — stupeň-deňová metóda)
            ti = _f(data, f'rok{i}_ti', 20)
            te = _f(data, f'rok{i}_te', -15)
            tepr = _f(data, f'rok{i}_tepr')
            d = _f(data, f'rok{i}_pocet_dni', 200)
            if d > 0 and (ti - tepr) != 0:
                potrebny_vykon = round(spotreba_vykurovanie * (ti - te) / (24 * d * (ti - tepr)), 2)
            else:
                potrebny_vykon = 0

            roky.append({
                'rok': rok, 'odber_paliva': odber, 'vyhrevnost': vyhrevnost,
                'teplo_v_palive': teplo_v_palive, 'vyrobene_teplo': vyrobene_teplo,
                'ucinnost_priama': ucinnost_priama,
                'spotreba_vykurovanie': spotreba_vykurovanie, 'spotreba_tuv': spotreba_tuv,
                'merna_spotreba_tuv': merna_spotreba_tuv, 'podiel_tuv': podiel_tuv,
                'ti': ti, 'te': te, 'tepr': tepr, 'pocet_dni': d,
                'potrebny_vykon': potrebny_vykon,
            })

        instalovany_vykon = round(sum(_f(data, f'{p}_menovity_vykon') for p in ['k1', 'k2']), 2)
        posledny_potrebny_vykon = roky[-1]['potrebny_vykon'] if roky else 0
        vykon_rezerva = round(instalovany_vykon - posledny_potrebny_vykon, 2)

        # --- Celkové vyhodnotenie ---
        if kotly:
            priemerna_ucinnost = round(sum(k['priemerna_ucinnost'] for k in kotly) / len(kotly), 2)
            celkovy_stav = 'Vyhovuje' if all(k['stav'] == 'Vyhovuje' for k in kotly) else 'Nevyhovuje'
        else:
            priemerna_ucinnost = 0
            celkovy_stav = 'N/A'

        # --- Termín nasledujúcej kontroly ---
        interval = int(data.get('interval_kontroly') or 4)
        nasledujuca_kontrola = ''
        datum_kontroly = data.get('datum_kontroly', '')
        try:
            mesiac, rok_k = datum_kontroly.split('/')
            nasledujuca_kontrola = f"{mesiac}/{int(rok_k) + interval}"
        except Exception:
            nasledujuca_kontrola = ''

        return {
            'kotly': kotly,
            'min_ucinnost': min_ucinnost,
            'priemerna_ucinnost': priemerna_ucinnost,
            'celkovy_stav': celkovy_stav,
            'roky': roky,
            'instalovany_vykon': instalovany_vykon,
            'potrebny_vykon': posledny_potrebny_vykon,
            'vykon_rezerva': vykon_rezerva,
            'interval_kontroly': interval,
            'nasledujuca_kontrola': nasledujuca_kontrola,
        }
    except Exception:
        return {
            'kotly': [], 'min_ucinnost': 96, 'priemerna_ucinnost': 0,
            'celkovy_stav': 'N/A', 'roky': [], 'instalovany_vykon': 0,
            'potrebny_vykon': 0, 'vykon_rezerva': 0, 'interval_kontroly': 4,
            'nasledujuca_kontrola': '',
        }

# ============================================================
# GENEROVANIE WORD - pomocné funkcie pre tabuľky
# ============================================================
def _cell(cell, text, bold=False, size=9, align=None):
    cell.text = ''
    p = cell.paragraphs[0]
    if align is not None:
        p.alignment = align
    run = p.add_run('' if text is None else str(text))
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = 'Times New Roman'

def _docx_table(doc, header_row, data_rows):
    table = doc.add_table(rows=1, cols=len(header_row))
    table.style = 'Table Grid'
    for i, h in enumerate(header_row):
        _cell(table.rows[0].cells[i], h, bold=True)
    for row in data_rows:
        cells = table.add_row().cells
        for i, val in enumerate(row):
            _cell(cells[i], val)
    doc.add_paragraph("")
    return table

def _docx_section_heading(doc, text):
    h = doc.add_heading(text, level=1)
    for run in h.runs:
        run.font.name = 'Times New Roman'
        run.font.color.rgb = RGBColor(0x1a, 0x1a, 0x1a)

def _docx_kotly_identifikacia(doc, kotly):
    header = ["Parameter"] + [f"Kotol {k['oznacenie']}" for k in kotly]
    riadky = [
        ("Výrobca kotla", 'vyrobca'), ("Typ kotla", 'typ'),
        ("Výrobné číslo kotla", 'vyrobne_cislo'), ("Rok výroby kotla", 'rok_vyroby'),
        ("Prevádzkový stav", 'prevadzkovy_stav'),
        ("Menovitý výkon (kW)", 'menovity_vykon'), ("Maximálny výkon 80/60 °C (kW)", 'max_vykon'),
        ("Maximálny príkon (kW)", 'max_prikon'), ("Minimálny výkon 80/60 °C (kW)", 'min_vykon'),
        ("Minimálny príkon (kW)", 'min_prikon'), ("Kondenzačný/nekondenzačný", 'kondenzacny'),
        ("Označenie CE", 'oznacenie_ce'), ("Druh paliva", 'typ_paliva'),
        ("Spôsob dávkovania paliva", 'sposob_davkovania'), ("Typ výkonovej regulácie", 'typ_regulacie'),
        ("Spôsob odvodu spalín", 'sposob_odvodu_spalin'), ("Spôsob prívodu vzduchu", 'sposob_privodu_vzduchu'),
        ("Teplonosné médium", 'teplonosne_medium'), ("Spôsob využitia kotla", 'sposob_vyuzitia'),
    ]
    data_rows = []
    for label, key in riadky:
        data_rows.append([label] + [k.get(key, '') for k in kotly])
    _docx_table(doc, header, data_rows)

# ============================================================
# GENEROVANIE WORD
# ============================================================
def generuj_word(typ, data, vysledky):
    doc = Document()

    # Štýl nadpisu
    nazov = {
        'audit': 'ENERGETICKÝ AUDIT',
        'certifikat': 'ENERGETICKÝ CERTIFIKÁT',
        'vykurovanie': 'Správa z vykonanej pravidelnej kontroly vykurovacieho systému'
    }.get(typ, 'SPRÁVA')

    h = doc.add_heading(nazov, 0)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in h.runs:
        run.font.name = 'Times New Roman'
        run.font.color.rgb = RGBColor(0x8B, 0x1A, 0x1A)

    if typ == 'vykurovanie':
        kotly = vysledky.get('kotly', [])
        roky = vysledky.get('roky', [])

        sub = doc.add_paragraph("podľa Zákona č. 314/2012 Z.z.")
        sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in sub.runs:
            run.font.name = 'Times New Roman'
        doc.add_paragraph("")

        for label, key, extra in [
            ("Objekt", 'nazov_budovy', None), ("Lokalita", 'adresa', None),
            ("Majiteľ", 'vlastnik', None), ("Adresa vlastníka", 'adresa_vlastnika', None),
            ("Správca", 'spravca', None), ("Prevádzkovateľ", 'prevadzkovatel', None),
        ]:
            p = doc.add_paragraph()
            r1 = p.add_run(f"{label}: ")
            r1.bold = True
            r1.font.name = 'Times New Roman'
            r2 = p.add_run(str(data.get(key, '')))
            r2.font.name = 'Times New Roman'
        doc.add_paragraph("")
        for label, key in [
            ("Vypracoval", 'vypracoval_firma'), ("Kontakt", 'vypracoval_kontakt'),
            ("Dátum kontroly", 'datum_kontroly'), ("Poradové číslo", 'poradove_cislo'),
        ]:
            p = doc.add_paragraph()
            r1 = p.add_run(f"{label}: ")
            r1.bold = True
            r1.font.name = 'Times New Roman'
            r2 = p.add_run(str(data.get(key, '')))
            r2.font.name = 'Times New Roman'

        doc.add_page_break()

        # --- 1. KONTROLA KOTLA ---
        _docx_section_heading(doc, "1. Kontrola kotla")
        doc.add_paragraph("1.1 Identifikačné údaje kotla").runs[0].bold = True
        if kotly:
            _docx_kotly_identifikacia(doc, kotly)
        else:
            doc.add_paragraph("Nebol zadaný žiadny kotol.")

        # --- 2. VIZUÁLNA KONTROLA ---
        _docx_section_heading(doc, "2. Vizuálna kontrola a zhodnotenie funkčnosti a údržby kotla")
        vizualna = [
            ("Stav kotla", data.get('stav_kotla', '')), ("Únik paliva", data.get('unik_paliva', '')),
            ("Únik teplonosnej látky", data.get('unik_teplonosnej_latky', '')),
            ("Znečistenie spaľovacej komory / horákov", data.get('znecistenie', '')),
            ("Funkčnosť armatúr", data.get('armatury', '')),
            ("Kvalita teplonosnej látky", data.get('kvalita_vody', '')),
            ("Meracie prístroje", data.get('meracie_pristroje', '')),
            ("Čistota kotolne", data.get('cistota_kotolne', '')),
            ("Doklady o údržbe a opravách", data.get('doklady_udrzba', '')),
        ]
        _docx_table(doc, ["Kontrolovaná položka", "Zistený stav"], vizualna)

        # --- 3. ROZŠÍRENÁ KONTROLA ---
        _docx_section_heading(doc, "3. Rozšírená kontrola vykurovacieho systému")
        rozsirena = [
            ("Vykurovacie telesá", data.get('vykurovacie_telesa', '')),
            ("Tepelná izolácia rozvodov tepla", data.get('tepelna_izolacia_rozvodov', '')),
            ("Čistota obehovej vody", data.get('cistota_obehovej_vody', '')),
            ("Stav rozvodov tepla a TÚV", data.get('stav_rozvodov', '')),
            ("Zmena využívania od poslednej kontroly", data.get('zmena_vyuzivania', '')),
        ]
        _docx_table(doc, ["Kontrolovaná položka", "Zistený stav"], rozsirena)

        # --- 4. MERANIE ÚČINNOSTI KOTLA ---
        _docx_section_heading(doc, "4. Meranie účinnosti kotla — komínová strata")
        for k in kotly:
            doc.add_paragraph(f"Kotol {k['oznacenie']} — {k['vyrobca']} {k['typ']}").runs[0].bold = True
            if k['merania']:
                header = ["Parameter"] + [m['zatazenie'] for m in k['merania']]
                riadky = [
                    ["Teplota spaľovacieho vzduchu (°C)"] + [m['t_vzduch'] for m in k['merania']],
                    ["Teplota spalín (°C)"] + [m['t_spalin'] for m in k['merania']],
                    ["Obsah O₂ v spalinách (%)"] + [m['o2'] for m in k['merania']],
                    ["Komínová strata (%)"] + [m['komin_strata'] for m in k['merania']],
                    ["Účinnosť kotla (%)"] + [m['ucinnost'] for m in k['merania']],
                ]
                _docx_table(doc, header, riadky)
            doc.add_paragraph(f"Priemerná účinnosť kotla: {k['priemerna_ucinnost']} %")
            if k['garantovana_ucinnost']:
                doc.add_paragraph(f"Garantovaná účinnosť podľa výrobcu: {k['garantovana_ucinnost']} %")
            doc.add_paragraph(f"Minimálna požadovaná účinnosť podľa Vyhl. č. 328/2005 Z.z.: {vysledky.get('min_ucinnost')} %")
            p = doc.add_paragraph()
            p.add_run(f"Vyhodnotenie: {k['stav']}").bold = True
            doc.add_paragraph("")

        # --- 5. PRIAMA METÓDA ---
        _docx_section_heading(doc, "5. Vyhodnotenie účinnosti výroby tepla priamou metódou")
        if roky:
            header = ["Ukazovateľ"] + [str(r['rok']) for r in roky]
            riadky = [
                ["Odber paliva"] + [r['odber_paliva'] for r in roky],
                ["Výhrevnosť paliva (kWh/jedn.)"] + [r['vyhrevnost'] for r in roky],
                ["Teplo v palive (kWh)"] + [r['teplo_v_palive'] for r in roky],
                ["Vyrobené teplo (kWh)"] + [r['vyrobene_teplo'] for r in roky],
                ["Účinnosť výroby tepla priamou metódou (%)"] + [r['ucinnost_priama'] for r in roky],
                ["Spotreba tepla na vykurovanie (kWh)"] + [r['spotreba_vykurovanie'] for r in roky],
                ["Spotreba tepla na TÚV (kWh)"] + [r['spotreba_tuv'] for r in roky],
                ["Merná spotreba tepla na TÚV (kWh/m³)"] + [r['merna_spotreba_tuv'] for r in roky],
                ["Podiel TÚV z celkovej spotreby (%)"] + [r['podiel_tuv'] for r in roky],
            ]
            _docx_table(doc, header, riadky)
        else:
            doc.add_paragraph("Neboli zadané ročné údaje o spotrebe paliva.")

        # --- 6. POSÚDENIE VÝKONU ---
        _docx_section_heading(doc, "6. Posúdenie výkonu kotla vzhľadom na potrebu tepla v budove")
        if roky:
            header = ["Ukazovateľ"] + [str(r['rok']) for r in roky]
            riadky = [
                ["ti - te (°C)"] + [round(r['ti'] - r['te'], 1) for r in roky],
                ["ti - tepr (°C)"] + [round(r['ti'] - r['tepr'], 1) for r in roky],
                ["Počet vykurovacích dní"] + [r['pocet_dni'] for r in roky],
                ["Potrebný tepelný výkon (kW)"] + [r['potrebny_vykon'] for r in roky],
            ]
            _docx_table(doc, header, riadky)
        doc.add_paragraph(f"Inštalovaný výkon kotlov v objekte spolu: {vysledky.get('instalovany_vykon')} kW")
        doc.add_paragraph(f"Potreba tepelného výkonu v objekte (posledný rok): {vysledky.get('potrebny_vykon')} kW")
        rezerva = vysledky.get('vykon_rezerva', 0)
        if rezerva >= 0:
            doc.add_paragraph(f"Inštalovaný výkon kotolne pokrýva potrebu tepelného výkonu s rezervou {rezerva} kW.")
        else:
            doc.add_paragraph(f"Inštalovaný výkon kotolne nepokrýva potrebu tepelného výkonu (chýba {abs(rezerva)} kW).")

        # --- 7. VYHODNOTENIE KONTROLY KOTLA ---
        _docx_section_heading(doc, "7. Vyhodnotenie kontroly kotla a návrhy na opatrenia")
        doc.add_paragraph(f"Minimálna požadovaná priemerná účinnosť podľa Vyhlášky č. 328/2005 Z.z.: {vysledky.get('min_ucinnost')} %")
        for k in kotly:
            doc.add_paragraph(f"Nameraná priemerná účinnosť kotla {k['oznacenie']}: {k['priemerna_ucinnost']} %")
        p = doc.add_paragraph()
        p.add_run(f"Kotly {'spĺňajú' if vysledky.get('celkovy_stav') == 'Vyhovuje' else 'nespĺňajú'} požiadavky Vyhlášky č. 328/2005 Z.z.").bold = True

        # --- 8. VYHODNOTENIE ROZŠÍRENEJ KONTROLY ---
        _docx_section_heading(doc, "8. Vyhodnotenie rozšírenej kontroly vykurovacieho systému a návrh opatrení")
        if data.get('navrh_opatreni'):
            doc.add_paragraph(f"Pre zlepšenie prevádzkového stavu navrhujeme tieto opatrenia: {data.get('navrh_opatreni')}")
        doc.add_paragraph(f"Nasledujúcu kontrolu v zmysle Zákona č. 314/2012 Z.z. je potrebné vykonať do: {vysledky.get('nasledujuca_kontrola')}")
        doc.add_paragraph("")
        doc.add_paragraph("Vlastník, prevádzkovateľ: .....................................")
        doc.add_paragraph("Oprávnená osoba: .....................................")
        doc.add_paragraph(f"Dňa: {data.get('datum_kontroly', '.....................................')}")
        doc.add_paragraph("")
        doc.add_paragraph(f"Vypracoval: {data.get('vypracoval_firma', '')}")

        section = doc.sections[0]
        footer = section.footer
        fp = footer.paragraphs[0]
        fp.text = f"{data.get('poradove_cislo', '')}    {data.get('vypracoval_firma', '')}"
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in fp.runs:
            run.font.size = Pt(8)
            run.font.name = 'Times New Roman'

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        return buf

    doc.add_paragraph(f"Dátum: {datetime.now().strftime('%d.%m.%Y')}")
    doc.add_paragraph(f"Budova: {data.get('nazov_budovy', '')}")
    doc.add_paragraph(f"Adresa: {data.get('adresa', '')}")
    doc.add_paragraph(f"Vlastník: {data.get('vlastnik', '')}")
    doc.add_paragraph("")

    doc.add_heading("Vstupné údaje", level=1)
    for k, v in data.items():
        if k not in ['nazov_budovy', 'adresa', 'vlastnik']:
            doc.add_paragraph(f"{k.replace('_', ' ').capitalize()}: {v}")

    doc.add_heading("Výsledky", level=1)
    for k, v in vysledky.items():
        doc.add_paragraph(f"{k.replace('_', ' ').capitalize()}: {v}")

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# ============================================================
# GENEROVANIE PDF - pomocné funkcie
# ============================================================
def _pdf_table(rows, col_widths=None, header=True):
    t = Table(rows, colWidths=col_widths)
    style = [
        ('FONTNAME', (0, 0), (-1, -1), PDF_FONT),
        ('FONTSIZE', (0, 0), (-1, -1), 8.5),
        ('GRID', (0, 0), (-1, -1), 0.6, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
    ]
    if header:
        style.append(('FONTNAME', (0, 0), (-1, 0), PDF_FONT_BOLD))
        style.append(('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#e5e5e5')))
    t.setStyle(TableStyle(style))
    return t

def _pdf_footer(poradove_cislo, firma):
    def _draw(canvas, doc):
        canvas.saveState()
        canvas.setFont(PDF_FONT, 8)
        canvas.drawString(2*cm, 1.3*cm, str(poradove_cislo or ''))
        canvas.drawRightString(A4[0] - 2*cm, 1.3*cm, f"{firma or ''}    strana {doc.page}")
        canvas.restoreState()
    return _draw

# ============================================================
# GENEROVANIE PDF
# ============================================================
def generuj_pdf(typ, data, vysledky):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            rightMargin=2*cm, leftMargin=2*cm,
                            topMargin=2*cm, bottomMargin=2.2*cm)
    styles = getSampleStyleSheet()
    story = []

    nazov = {
        'audit': 'ENERGETICKÝ AUDIT',
        'certifikat': 'ENERGETICKÝ CERTIFIKÁT',
        'vykurovanie': 'Správa z vykonanej pravidelnej kontroly vykurovacieho systému'
    }.get(typ, 'SPRÁVA')

    if typ == 'vykurovanie':
        title_style = ParagraphStyle('title', parent=styles['Title'], fontName=PDF_FONT_BOLD,
                                      fontSize=16, leading=20, alignment=TA_CENTER,
                                      textColor=colors.HexColor('#8B1A1A'), spaceAfter=6)
        h1 = ParagraphStyle('h1', parent=styles['Heading1'], fontName=PDF_FONT_BOLD, fontSize=12.5,
                             textColor=colors.black, spaceBefore=14, spaceAfter=8)
        normal = ParagraphStyle('n', parent=styles['Normal'], fontName=PDF_FONT, fontSize=9.5, leading=13)
        bold = ParagraphStyle('b', parent=normal, fontName=PDF_FONT_BOLD)

        story.append(Paragraph(nazov, title_style))
        story.append(Paragraph("podľa Zákona č. 314/2012 Z.z.", ParagraphStyle('sub', parent=normal, alignment=TA_CENTER)))
        story.append(Spacer(1, 1*cm))

        for label, key in [
            ("Objekt", 'nazov_budovy'), ("Lokalita", 'adresa'), ("Majiteľ", 'vlastnik'),
            ("Adresa vlastníka", 'adresa_vlastnika'), ("Správca", 'spravca'), ("Prevádzkovateľ", 'prevadzkovatel'),
        ]:
            story.append(Paragraph(f"<b>{label}:</b> {data.get(key, '')}", normal))
        story.append(Spacer(1, 0.6*cm))
        for label, key in [
            ("Vypracoval", 'vypracoval_firma'), ("Kontakt", 'vypracoval_kontakt'),
            ("Dátum kontroly", 'datum_kontroly'), ("Poradové číslo", 'poradove_cislo'),
        ]:
            story.append(Paragraph(f"<b>{label}:</b> {data.get(key, '')}", normal))
        story.append(PageBreak())

        kotly = vysledky.get('kotly', [])
        roky = vysledky.get('roky', [])

        # --- 1. KONTROLA KOTLA ---
        story.append(Paragraph("1. Kontrola kotla", h1))
        story.append(Paragraph("1.1 Identifikačné údaje kotla", bold))
        story.append(Spacer(1, 0.2*cm))
        if kotly:
            header_row = ["Parameter"] + [f"Kotol {k['oznacenie']}" for k in kotly]
            riadky = [
                ("Výrobca kotla", 'vyrobca'), ("Typ kotla", 'typ'),
                ("Výrobné číslo kotla", 'vyrobne_cislo'), ("Rok výroby kotla", 'rok_vyroby'),
                ("Prevádzkový stav", 'prevadzkovy_stav'), ("Menovitý výkon (kW)", 'menovity_vykon'),
                ("Maximálny výkon 80/60°C (kW)", 'max_vykon'), ("Maximálny príkon (kW)", 'max_prikon'),
                ("Minimálny výkon 80/60°C (kW)", 'min_vykon'), ("Minimálny príkon (kW)", 'min_prikon'),
                ("Kondenzačný/nekondenzačný", 'kondenzacny'), ("Označenie CE", 'oznacenie_ce'),
                ("Druh paliva", 'typ_paliva'), ("Spôsob dávkovania paliva", 'sposob_davkovania'),
                ("Typ výkonovej regulácie", 'typ_regulacie'), ("Spôsob odvodu spalín", 'sposob_odvodu_spalin'),
                ("Spôsob prívodu vzduchu", 'sposob_privodu_vzduchu'), ("Teplonosné médium", 'teplonosne_medium'),
                ("Spôsob využitia kotla", 'sposob_vyuzitia'),
            ]
            data_rows = [header_row] + [[label] + [str(k.get(key, '')) for k in kotly] for label, key in riadky]
            col0 = 5.5*cm
            colN = (17*cm - col0) / len(kotly)
            story.append(_pdf_table(data_rows, col_widths=[col0] + [colN]*len(kotly)))
        else:
            story.append(Paragraph("Nebol zadaný žiadny kotol.", normal))
        story.append(Spacer(1, 0.5*cm))

        # --- 2. VIZUÁLNA KONTROLA ---
        story.append(Paragraph("2. Vizuálna kontrola a zhodnotenie funkčnosti a údržby kotla", h1))
        vizualna = [["Kontrolovaná položka", "Zistený stav"]] + [
            [label, data.get(key, '') or '—'] for label, key in [
                ("Stav kotla", 'stav_kotla'), ("Únik paliva", 'unik_paliva'),
                ("Únik teplonosnej látky", 'unik_teplonosnej_latky'),
                ("Znečistenie spaľovacej komory / horákov", 'znecistenie'),
                ("Funkčnosť armatúr", 'armatury'), ("Kvalita teplonosnej látky", 'kvalita_vody'),
                ("Meracie prístroje", 'meracie_pristroje'), ("Čistota kotolne", 'cistota_kotolne'),
                ("Doklady o údržbe a opravách", 'doklady_udrzba'),
            ]
        ]
        story.append(_pdf_table(vizualna, col_widths=[9*cm, 8*cm]))
        story.append(Spacer(1, 0.5*cm))

        # --- 3. ROZŠÍRENÁ KONTROLA ---
        story.append(Paragraph("3. Rozšírená kontrola vykurovacieho systému", h1))
        rozsirena = [["Kontrolovaná položka", "Zistený stav"]] + [
            [label, data.get(key, '') or '—'] for label, key in [
                ("Vykurovacie telesá", 'vykurovacie_telesa'),
                ("Tepelná izolácia rozvodov tepla", 'tepelna_izolacia_rozvodov'),
                ("Čistota obehovej vody", 'cistota_obehovej_vody'),
                ("Stav rozvodov tepla a TÚV", 'stav_rozvodov'),
                ("Zmena využívania od poslednej kontroly", 'zmena_vyuzivania'),
            ]
        ]
        story.append(_pdf_table(rozsirena, col_widths=[9*cm, 8*cm]))
        story.append(PageBreak())

        # --- 4. MERANIE ÚČINNOSTI KOTLA ---
        story.append(Paragraph("4. Meranie účinnosti kotla — komínová strata", h1))
        for k in kotly:
            story.append(Paragraph(f"Kotol {k['oznacenie']} — {k['vyrobca']} {k['typ']}", bold))
            story.append(Spacer(1, 0.15*cm))
            if k['merania']:
                header_row = ["Parameter"] + [m['zatazenie'] for m in k['merania']]
                riadky = [
                    ["Teplota spaľovacieho vzduchu (°C)"] + [str(m['t_vzduch']) for m in k['merania']],
                    ["Teplota spalín (°C)"] + [str(m['t_spalin']) for m in k['merania']],
                    ["Obsah O₂ v spalinách (%)"] + [str(m['o2']) for m in k['merania']],
                    ["Komínová strata (%)"] + [str(m['komin_strata']) for m in k['merania']],
                    ["Účinnosť kotla (%)"] + [str(m['ucinnost']) for m in k['merania']],
                ]
                colN = (17*cm - 7*cm) / max(len(k['merania']), 1)
                story.append(_pdf_table([header_row] + riadky, col_widths=[7*cm] + [colN]*len(k['merania'])))
            story.append(Paragraph(f"Priemerná účinnosť kotla: <b>{k['priemerna_ucinnost']} %</b>", normal))
            if k['garantovana_ucinnost']:
                story.append(Paragraph(f"Garantovaná účinnosť podľa výrobcu: {k['garantovana_ucinnost']} %", normal))
            story.append(Paragraph(f"Minimálna požadovaná účinnosť podľa Vyhl. č. 328/2005 Z.z.: {vysledky.get('min_ucinnost')} %", normal))
            story.append(Paragraph(f"<b>Vyhodnotenie: {k['stav']}</b>", normal))
            story.append(Spacer(1, 0.4*cm))

        # --- 5. PRIAMA METÓDA ---
        story.append(Paragraph("5. Vyhodnotenie účinnosti výroby tepla priamou metódou", h1))
        if roky:
            header_row = ["Ukazovateľ"] + [str(r['rok']) for r in roky]
            riadky = [
                ["Odber paliva"] + [str(r['odber_paliva']) for r in roky],
                ["Výhrevnosť paliva (kWh/jedn.)"] + [str(r['vyhrevnost']) for r in roky],
                ["Teplo v palive (kWh)"] + [str(r['teplo_v_palive']) for r in roky],
                ["Vyrobené teplo (kWh)"] + [str(r['vyrobene_teplo']) for r in roky],
                ["Účinnosť priamou metódou (%)"] + [str(r['ucinnost_priama']) for r in roky],
                ["Spotreba tepla na vykurovanie (kWh)"] + [str(r['spotreba_vykurovanie']) for r in roky],
                ["Spotreba tepla na TÚV (kWh)"] + [str(r['spotreba_tuv']) for r in roky],
                ["Merná spotreba tepla na TÚV (kWh/m³)"] + [str(r['merna_spotreba_tuv']) for r in roky],
                ["Podiel TÚV z celkovej spotreby (%)"] + [str(r['podiel_tuv']) for r in roky],
            ]
            col0 = 7*cm
            colN = (17*cm - col0) / len(roky)
            story.append(_pdf_table([header_row] + riadky, col_widths=[col0] + [colN]*len(roky)))
        else:
            story.append(Paragraph("Neboli zadané ročné údaje o spotrebe paliva.", normal))
        story.append(Spacer(1, 0.5*cm))

        # --- 6. POSÚDENIE VÝKONU ---
        story.append(Paragraph("6. Posúdenie výkonu kotla vzhľadom na potrebu tepla v budove", h1))
        if roky:
            header_row = ["Ukazovateľ"] + [str(r['rok']) for r in roky]
            riadky = [
                ["ti - te (°C)"] + [str(round(r['ti'] - r['te'], 1)) for r in roky],
                ["ti - tepr (°C)"] + [str(round(r['ti'] - r['tepr'], 1)) for r in roky],
                ["Počet vykurovacích dní"] + [str(r['pocet_dni']) for r in roky],
                ["Potrebný tepelný výkon (kW)"] + [str(r['potrebny_vykon']) for r in roky],
            ]
            col0 = 7*cm
            colN = (17*cm - col0) / len(roky)
            story.append(_pdf_table([header_row] + riadky, col_widths=[col0] + [colN]*len(roky)))
            story.append(Spacer(1, 0.3*cm))
        story.append(Paragraph(f"Inštalovaný výkon kotlov v objekte spolu: <b>{vysledky.get('instalovany_vykon')} kW</b>", normal))
        story.append(Paragraph(f"Potreba tepelného výkonu v objekte (posledný rok): <b>{vysledky.get('potrebny_vykon')} kW</b>", normal))
        rezerva = vysledky.get('vykon_rezerva', 0)
        if rezerva >= 0:
            story.append(Paragraph(f"Inštalovaný výkon kotolne pokrýva potrebu tepelného výkonu s rezervou {rezerva} kW.", normal))
        else:
            story.append(Paragraph(f"Inštalovaný výkon kotolne nepokrýva potrebu tepelného výkonu (chýba {abs(rezerva)} kW).", normal))
        story.append(PageBreak())

        # --- 7. VYHODNOTENIE KONTROLY KOTLA ---
        story.append(Paragraph("7. Vyhodnotenie kontroly kotla a návrhy na opatrenia", h1))
        story.append(Paragraph(f"Minimálna požadovaná priemerná účinnosť podľa Vyhlášky č. 328/2005 Z.z.: {vysledky.get('min_ucinnost')} %", normal))
        for k in kotly:
            story.append(Paragraph(f"Nameraná priemerná účinnosť kotla {k['oznacenie']}: {k['priemerna_ucinnost']} %", normal))
        splna = 'spĺňajú' if vysledky.get('celkovy_stav') == 'Vyhovuje' else 'nespĺňajú'
        story.append(Paragraph(f"<b>Kotly {splna} požiadavky Vyhlášky č. 328/2005 Z.z.</b>", normal))
        story.append(Spacer(1, 0.5*cm))

        # --- 8. VYHODNOTENIE ROZŠÍRENEJ KONTROLY ---
        story.append(Paragraph("8. Vyhodnotenie rozšírenej kontroly vykurovacieho systému a návrh opatrení", h1))
        if data.get('navrh_opatreni'):
            story.append(Paragraph(f"Pre zlepšenie prevádzkového stavu navrhujeme tieto opatrenia: {data.get('navrh_opatreni')}", normal))
        story.append(Paragraph(f"Nasledujúcu kontrolu v zmysle Zákona č. 314/2012 Z.z. je potrebné vykonať do: <b>{vysledky.get('nasledujuca_kontrola')}</b>", normal))
        story.append(Spacer(1, 1*cm))
        story.append(Paragraph("Vlastník, prevádzkovateľ: .....................................", normal))
        story.append(Paragraph("Oprávnená osoba: .....................................", normal))
        story.append(Paragraph(f"Dňa: {data.get('datum_kontroly', '.....................................')}", normal))
        story.append(Spacer(1, 0.5*cm))
        story.append(Paragraph(f"Vypracoval: {data.get('vypracoval_firma', '')}", normal))

        footer_fn = _pdf_footer(data.get('poradove_cislo'), data.get('vypracoval_firma'))
        doc.build(story, onFirstPage=footer_fn, onLaterPages=footer_fn)
        buf.seek(0)
        return buf

    title_style = ParagraphStyle('title', parent=styles['Title'], fontSize=16, spaceAfter=20)
    story.append(Paragraph(nazov, title_style))
    story.append(Spacer(1, 0.5*cm))

    # Základné info
    info = [
        ['Dátum:', datetime.now().strftime('%d.%m.%Y')],
        ['Budova:', data.get('nazov_budovy', '')],
        ['Adresa:', data.get('adresa', '')],
        ['Vlastník:', data.get('vlastnik', '')],
    ]
    t = Table(info, colWidths=[5*cm, 12*cm])
    t.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
        ('FONTSIZE', (0,0), (-1,-1), 10),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold'),
    ]))
    story.append(t)
    story.append(Spacer(1, 0.5*cm))

    story.append(Paragraph("Vstupné údaje", styles['Heading2']))
    for k, v in data.items():
        if k not in ['nazov_budovy', 'adresa', 'vlastnik']:
            story.append(Paragraph(f"<b>{k.replace('_',' ').capitalize()}:</b> {v}", styles['Normal']))

    story.append(Spacer(1, 0.5*cm))
    story.append(Paragraph("Výsledky", styles['Heading2']))

    vysl_data = [['Ukazovateľ', 'Hodnota']]
    for k, v in vysledky.items():
        vysl_data.append([k.replace('_',' ').capitalize(), str(v)])

    tv = Table(vysl_data, colWidths=[9*cm, 8*cm])
    tv.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2c7a4b')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 10),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#f0f7f3')]),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('TOPPADDING', (0,0), (-1,-1), 8),
    ]))
    story.append(tv)

    doc.build(story)
    buf.seek(0)
    return buf

# ============================================================
# DEKORATOR - prihlasenie
# ============================================================
def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'logged_in' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated

# ============================================================
# ROUTES
# ============================================================
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        if request.form['username'] == USERNAME and request.form['password'] == PASSWORD:
            session['logged_in'] = True
            return redirect(url_for('dashboard'))
        flash('Nesprávne meno alebo heslo.')
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('index'))

@app.route('/dashboard')
@login_required
def dashboard():
    conn = get_db()
    zaznamy = conn.execute('SELECT * FROM audity ORDER BY datum DESC LIMIT 10').fetchall()
    conn.close()
    return render_template('dashboard.html', zaznamy=zaznamy)

# --- AUDIT ---
@app.route('/audit', methods=['GET', 'POST'])
@login_required
def audit():
    if request.method == 'POST':
        data = request.form.to_dict()
        vysledky = vypocitaj_audit(data)
        conn = get_db()
        conn.execute('INSERT INTO audity (typ, datum, nazov_budovy, adresa, vlastnik, udaje, vysledky) VALUES (?,?,?,?,?,?,?)',
            ('audit', datetime.now().strftime('%d.%m.%Y %H:%M'),
             data.get('nazov_budovy'), data.get('adresa'), data.get('vlastnik'),
             str(data), str(vysledky)))
        conn.commit()
        conn.close()
        return render_template('vysledok.html', typ='audit', data=data, vysledky=vysledky)
    return render_template('audit.html')

# --- CERTIFIKAT ---
@app.route('/certifikat', methods=['GET', 'POST'])
@login_required
def certifikat():
    if request.method == 'POST':
        data = request.form.to_dict()
        vysledky = vypocitaj_certifikat(data)
        conn = get_db()
        conn.execute('INSERT INTO audity (typ, datum, nazov_budovy, adresa, vlastnik, udaje, vysledky) VALUES (?,?,?,?,?,?,?)',
            ('certifikat', datetime.now().strftime('%d.%m.%Y %H:%M'),
             data.get('nazov_budovy'), data.get('adresa'), data.get('vlastnik'),
             str(data), str(vysledky)))
        conn.commit()
        conn.close()
        return render_template('vysledok.html', typ='certifikat', data=data, vysledky=vysledky)
    return render_template('certifikat.html')

# --- VYKUROVANIE ---
@app.route('/vykurovanie', methods=['GET', 'POST'])
@login_required
def vykurovanie():
    if request.method == 'POST':
        data = request.form.to_dict()
        vysledky = vypocitaj_vykurovanie(data)
        conn = get_db()
        conn.execute('INSERT INTO audity (typ, datum, nazov_budovy, adresa, vlastnik, udaje, vysledky) VALUES (?,?,?,?,?,?,?)',
            ('vykurovanie', datetime.now().strftime('%d.%m.%Y %H:%M'),
             data.get('nazov_budovy'), data.get('adresa'), data.get('vlastnik'),
             str(data), str(vysledky)))
        conn.commit()
        conn.close()
        return render_template('vysledok.html', typ='vykurovanie', data=data, vysledky=vysledky)
    return render_template('vykurovanie.html')

# --- STIAHNUTIE ---
@app.route('/stiahnut/<format>/<typ>', methods=['POST'])
@login_required
def stiahnut(format, typ):
    data = request.form.to_dict()
    vysledky_fn = {'audit': vypocitaj_audit, 'certifikat': vypocitaj_certifikat, 'vykurovanie': vypocitaj_vykurovanie}
    vysledky = vysledky_fn[typ](data)

    nazov = data.get('nazov_budovy', 'sprava').replace(' ', '_')
    datum = datetime.now().strftime('%Y%m%d')

    if format == 'word':
        buf = generuj_word(typ, data, vysledky)
        return send_file(buf, as_attachment=True,
                        download_name=f"{typ}_{nazov}_{datum}.docx",
                        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    else:
        buf = generuj_pdf(typ, data, vysledky)
        return send_file(buf, as_attachment=True,
                        download_name=f"{typ}_{nazov}_{datum}.pdf",
                        mimetype='application/pdf')

@app.route('/historia')
@login_required
def historia():
    conn = get_db()
    zaznamy = conn.execute('SELECT * FROM audity ORDER BY datum DESC').fetchall()
    conn.close()
    return render_template('historia.html', zaznamy=zaznamy)

with app.app_context():
    init_db()

if __name__ == '__main__':
    app.run(debug=True)
