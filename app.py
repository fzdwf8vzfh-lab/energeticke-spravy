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
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib import colors

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

def vypocitaj_vykurovanie(data):
    try:
        min_ucinnost = float(data.get('min_ucinnost') or 96)

        kotly = []
        for pref in ['k1', 'k2']:
            if not data.get(f'{pref}_menovity_vykon'):
                continue
            t_spalin = float(data.get(f'{pref}_teplota_spalin') or 0)
            t_vzduch = float(data.get(f'{pref}_teplota_vzduchu') or 0)
            o2 = float(data.get(f'{pref}_o2') or 0)
            palivo = data.get(f'{pref}_typ_paliva') or 'Zemný plyn'

            qA = round(komin_strata(t_spalin, t_vzduch, o2, palivo), 2)
            ucinnost = round(100 - qA - STRATA_SALANIM, 2)
            stav = 'Vyhovuje' if ucinnost >= min_ucinnost else 'Nevyhovuje'

            kotly.append({
                'oznacenie': pref.upper(),
                'vyrobca': data.get(f'{pref}_vyrobca', ''),
                'typ': data.get(f'{pref}_typ', ''),
                'vyrobne_cislo': data.get(f'{pref}_vyrobne_cislo', ''),
                'rok_vyroby': data.get(f'{pref}_rok_vyroby', ''),
                'menovity_vykon': data.get(f'{pref}_menovity_vykon', ''),
                'typ_paliva': palivo,
                'teplota_spalin': t_spalin,
                'teplota_vzduchu': t_vzduch,
                'o2': o2,
                'komin_strata': qA,
                'strata_salanim': STRATA_SALANIM,
                'ucinnost': ucinnost,
                'stav': stav,
            })

        # --- Priama metóda (z ročnej spotreby paliva) ---
        rocna_spotreba_paliva = float(data.get('rocna_spotreba_paliva') or 0)
        vyhrevnost = float(data.get('vyhrevnost') or 0)
        vyrobene_teplo = float(data.get('vyrobene_teplo') or 0)
        teplo_v_palive = round(rocna_spotreba_paliva * vyhrevnost, 2)
        ucinnost_priama = round((vyrobene_teplo / teplo_v_palive) * 100, 2) if teplo_v_palive > 0 else 0

        # --- Celkové vyhodnotenie ---
        if kotly:
            priemerna_ucinnost = round(sum(k['ucinnost'] for k in kotly) / len(kotly), 2)
            celkovy_stav = 'Vyhovuje' if all(k['stav'] == 'Vyhovuje' for k in kotly) else 'Nevyhovuje'
        else:
            priemerna_ucinnost = 0
            celkovy_stav = 'N/A'

        # --- Termín nasledujúcej kontroly ---
        interval = int(data.get('interval_kontroly') or 4)
        nasledujuca_kontrola = ''
        datum_kontroly = data.get('datum_kontroly', '')
        try:
            mesiac, rok = datum_kontroly.split('/')
            nasledujuca_kontrola = f"{mesiac}/{int(rok) + interval}"
        except Exception:
            nasledujuca_kontrola = ''

        return {
            'kotly': kotly,
            'min_ucinnost': min_ucinnost,
            'priemerna_ucinnost': priemerna_ucinnost,
            'celkovy_stav': celkovy_stav,
            'teplo_v_palive': teplo_v_palive,
            'ucinnost_priama_metoda': ucinnost_priama,
            'interval_kontroly': interval,
            'nasledujuca_kontrola': nasledujuca_kontrola,
        }
    except Exception:
        return {
            'kotly': [], 'min_ucinnost': 96, 'priemerna_ucinnost': 0,
            'celkovy_stav': 'N/A', 'teplo_v_palive': 0,
            'ucinnost_priama_metoda': 0, 'interval_kontroly': 4,
            'nasledujuca_kontrola': '',
        }

# ============================================================
# GENEROVANIE WORD
# ============================================================
def generuj_word(typ, data, vysledky):
    doc = Document()

    # Štýl nadpisu
    nazov = {
        'audit': 'ENERGETICKÝ AUDIT',
        'certifikat': 'ENERGETICKÝ CERTIFIKÁT',
        'vykurovanie': 'SPRÁVA Z PRAVIDELNEJ KONTROLY VYKUROVACIEHO SYSTÉMU'
    }.get(typ, 'SPRÁVA')

    h = doc.add_heading(nazov, 0)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER

    if typ == 'vykurovanie':
        doc.add_paragraph("podľa Zákona č. 314/2012 Z.z. a Vyhl. č. 422/2012 Z.z.").alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph("")
        doc.add_paragraph(f"Objekt: {data.get('nazov_budovy', '')}")
        doc.add_paragraph(f"Lokalita: {data.get('adresa', '')}")
        doc.add_paragraph(f"Majiteľ: {data.get('vlastnik', '')}")
        doc.add_paragraph(f"Správca: {data.get('spravca', '')}")
        doc.add_paragraph(f"Vypracoval: {data.get('vypracoval_firma', '')}")
        doc.add_paragraph(f"Dátum kontroly: {data.get('datum_kontroly', '')}")
        doc.add_paragraph(f"Poradové číslo: {data.get('poradove_cislo', '')}")
        doc.add_paragraph("")

        doc.add_heading("1. Identifikačné údaje kotlov", level=1)
        for k in vysledky.get('kotly', []):
            doc.add_heading(f"Kotol {k['oznacenie']}", level=2)
            for label, val in [
                ("Výrobca", k['vyrobca']), ("Typ", k['typ']),
                ("Výrobné číslo", k['vyrobne_cislo']), ("Rok výroby", k['rok_vyroby']),
                ("Menovitý výkon", f"{k['menovity_vykon']} kW"), ("Druh paliva", k['typ_paliva']),
            ]:
                doc.add_paragraph(f"{label}: {val}")

        doc.add_heading("2. Meranie účinnosti kotla — komínová strata", level=1)
        for k in vysledky.get('kotly', []):
            doc.add_heading(f"Kotol {k['oznacenie']}", level=2)
            doc.add_paragraph(f"Teplota spalín: {k['teplota_spalin']} °C")
            doc.add_paragraph(f"Teplota spaľovacieho vzduchu: {k['teplota_vzduchu']} °C")
            doc.add_paragraph(f"Obsah O₂ v spalinách: {k['o2']} %")
            doc.add_paragraph(f"Komínová strata: {k['komin_strata']} %")
            doc.add_paragraph(f"Strata sálaním (odhad): {k['strata_salanim']} %")
            doc.add_paragraph(f"Vypočítaná účinnosť kotla: {k['ucinnost']} %")
            doc.add_paragraph(f"Minimálna požadovaná účinnosť (Vyhl. č. 328/2005 Z.z.): {vysledky.get('min_ucinnost')} %")
            p = doc.add_paragraph()
            p.add_run(f"Vyhodnotenie: {k['stav']}").bold = True

        doc.add_heading("3. Účinnosť výroby tepla priamou metódou", level=1)
        doc.add_paragraph(f"Ročná spotreba paliva: {data.get('rocna_spotreba_paliva', '')}")
        doc.add_paragraph(f"Výhrevnosť paliva: {data.get('vyhrevnost', '')} kWh/jedn.")
        doc.add_paragraph(f"Teplo v palive: {vysledky.get('teplo_v_palive')} kWh")
        doc.add_paragraph(f"Vyrobené teplo: {data.get('vyrobene_teplo', '')} kWh")
        doc.add_paragraph(f"Účinnosť výroby tepla priamou metódou: {vysledky.get('ucinnost_priama_metoda')} %")

        doc.add_heading("4. Vizuálna kontrola a stav vykurovacej sústavy", level=1)
        for label, key in [
            ("Stav kotla", 'stav_kotla'), ("Únik paliva", 'unik_paliva'),
            ("Únik teplonosnej látky", 'unik_teplonosnej_latky'),
            ("Znečistenie spaľovacej komory / horákov", 'znecistenie'),
            ("Funkčnosť armatúr", 'armatury'), ("Kvalita teplonosnej látky", 'kvalita_vody'),
            ("Meracie prístroje", 'meracie_pristroje'), ("Čistota kotolne", 'cistota_kotolne'),
            ("Vykurovacie telesá", 'vykurovacie_telesa'), ("Tepelná izolácia rozvodov", 'tepelna_izolacia_rozvodov'),
            ("Stav rozvodov tepla a TÚV", 'stav_rozvodov'), ("Zmena využívania od poslednej kontroly", 'zmena_vyuzivania'),
        ]:
            if data.get(key):
                doc.add_paragraph(f"{label}: {data.get(key)}")

        doc.add_heading("5. Vyhodnotenie kontroly a návrh opatrení", level=1)
        doc.add_paragraph(f"Priemerná účinnosť kotlov: {vysledky.get('priemerna_ucinnost')} %")
        p = doc.add_paragraph()
        p.add_run(f"Celkové vyhodnotenie: {vysledky.get('celkovy_stav')}").bold = True
        if data.get('navrh_opatreni'):
            doc.add_paragraph(f"Návrh opatrení: {data.get('navrh_opatreni')}")
        doc.add_paragraph(f"Nasledujúcu kontrolu je potrebné vykonať do: {vysledky.get('nasledujuca_kontrola')}")

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
# GENEROVANIE PDF
# ============================================================
def generuj_pdf(typ, data, vysledky):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            rightMargin=2*cm, leftMargin=2*cm,
                            topMargin=2*cm, bottomMargin=2*cm)
    styles = getSampleStyleSheet()
    story = []

    nazov = {
        'audit': 'ENERGETICKÝ AUDIT',
        'certifikat': 'ENERGETICKÝ CERTIFIKÁT',
        'vykurovanie': 'SPRÁVA Z PRAVIDELNEJ KONTROLY VYKUROVACIEHO SYSTÉMU'
    }.get(typ, 'SPRÁVA')

    title_style = ParagraphStyle('title', parent=styles['Title'], fontSize=16, spaceAfter=20)
    story.append(Paragraph(nazov, title_style))
    story.append(Spacer(1, 0.5*cm))

    if typ == 'vykurovanie':
        story.append(Paragraph("podľa Zákona č. 314/2012 Z.z. a Vyhl. č. 422/2012 Z.z.", styles['Normal']))
        story.append(Spacer(1, 0.5*cm))

        info = [
            ['Objekt:', data.get('nazov_budovy', '')],
            ['Lokalita:', data.get('adresa', '')],
            ['Majiteľ:', data.get('vlastnik', '')],
            ['Správca:', data.get('spravca', '')],
            ['Vypracoval:', data.get('vypracoval_firma', '')],
            ['Dátum kontroly:', data.get('datum_kontroly', '')],
            ['Poradové číslo:', data.get('poradove_cislo', '')],
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

        story.append(Paragraph("Meranie účinnosti kotlov — komínová strata", styles['Heading2']))
        kotol_data = [['Kotol', 'Výrobca / typ', 'T spalín (°C)', 'T vzduchu (°C)', 'O₂ (%)', 'Účinnosť (%)', 'Vyhodnotenie']]
        for k in vysledky.get('kotly', []):
            kotol_data.append([
                k['oznacenie'], f"{k['vyrobca']} {k['typ']}", str(k['teplota_spalin']),
                str(k['teplota_vzduchu']), str(k['o2']), str(k['ucinnost']), k['stav']
            ])
        tk = Table(kotol_data, colWidths=[1.7*cm, 4.3*cm, 2.3*cm, 2.3*cm, 1.8*cm, 2.3*cm, 2.3*cm])
        tk.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2c7a4b')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,-1), 8),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#f0f7f3')]),
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ('BOTTOMPADDING', (0,0), (-1,-1), 6),
            ('TOPPADDING', (0,0), (-1,-1), 6),
        ]))
        story.append(tk)
        story.append(Paragraph(f"Minimálna požadovaná účinnosť podľa Vyhl. č. 328/2005 Z.z.: {vysledky.get('min_ucinnost')} %", styles['Normal']))
        story.append(Spacer(1, 0.5*cm))

        story.append(Paragraph("Účinnosť výroby tepla priamou metódou", styles['Heading2']))
        story.append(Paragraph(f"<b>Teplo v palive:</b> {vysledky.get('teplo_v_palive')} kWh", styles['Normal']))
        story.append(Paragraph(f"<b>Vyrobené teplo:</b> {data.get('vyrobene_teplo', '')} kWh", styles['Normal']))
        story.append(Paragraph(f"<b>Účinnosť výroby tepla priamou metódou:</b> {vysledky.get('ucinnost_priama_metoda')} %", styles['Normal']))
        story.append(Spacer(1, 0.5*cm))

        story.append(Paragraph("Vizuálna kontrola a stav vykurovacej sústavy", styles['Heading2']))
        for label, key in [
            ("Stav kotla", 'stav_kotla'), ("Únik paliva", 'unik_paliva'),
            ("Únik teplonosnej látky", 'unik_teplonosnej_latky'),
            ("Znečistenie spaľovacej komory / horákov", 'znecistenie'),
            ("Funkčnosť armatúr", 'armatury'), ("Kvalita teplonosnej látky", 'kvalita_vody'),
            ("Meracie prístroje", 'meracie_pristroje'), ("Čistota kotolne", 'cistota_kotolne'),
            ("Vykurovacie telesá", 'vykurovacie_telesa'), ("Tepelná izolácia rozvodov", 'tepelna_izolacia_rozvodov'),
            ("Stav rozvodov tepla a TÚV", 'stav_rozvodov'), ("Zmena využívania od poslednej kontroly", 'zmena_vyuzivania'),
        ]:
            if data.get(key):
                story.append(Paragraph(f"<b>{label}:</b> {data.get(key)}", styles['Normal']))

        story.append(Spacer(1, 0.5*cm))
        story.append(Paragraph("Vyhodnotenie kontroly a návrh opatrení", styles['Heading2']))
        story.append(Paragraph(f"<b>Priemerná účinnosť kotlov:</b> {vysledky.get('priemerna_ucinnost')} %", styles['Normal']))
        story.append(Paragraph(f"<b>Celkové vyhodnotenie:</b> {vysledky.get('celkovy_stav')}", styles['Normal']))
        if data.get('navrh_opatreni'):
            story.append(Paragraph(f"<b>Návrh opatrení:</b> {data.get('navrh_opatreni')}", styles['Normal']))
        story.append(Paragraph(f"<b>Nasledujúcu kontrolu je potrebné vykonať do:</b> {vysledky.get('nasledujuca_kontrola')}", styles['Normal']))

        doc.build(story)
        buf.seek(0)
        return buf

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
