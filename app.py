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
# ============================================================
def vypocitaj_vykurovanie(data):
    try:
        vykon_kotla = float(data.get('vykon_kotla', 0))
        vek_kotla = int(data.get('vek_kotla', 0))
        ucinnost = float(data.get('ucinnost', 0))

        # Odporúčaná účinnosť podľa veku
        if vek_kotla < 5:
            odp_ucinnost = 92
        elif vek_kotla < 10:
            odp_ucinnost = 88
        elif vek_kotla < 20:
            odp_ucinnost = 82
        else:
            odp_ucinnost = 75

        rozdiel = ucinnost - odp_ucinnost
        stav = "Vyhovujúci" if rozdiel >= 0 else "Nevyhovujúci"
        odporucanie = "Kotol je v dobrom stave." if rozdiel >= 0 else f"Odporúčame výmenu kotla. Aktuálna účinnosť je o {abs(rozdiel):.1f}% nižšia ako odporúčaná."

        return {
            'odp_ucinnost': odp_ucinnost,
            'stav': stav,
            'odporucanie': odporucanie,
            'rozdiel': round(rozdiel, 1)
        }
    except:
        return {'odp_ucinnost': 0, 'stav': 'N/A', 'odporucanie': '', 'rozdiel': 0}

# ============================================================
# GENEROVANIE WORD
# ============================================================
def generuj_word(typ, data, vysledky):
    doc = Document()

    # Štýl nadpisu
    nazov = {
        'audit': 'ENERGETICKÝ AUDIT',
        'certifikat': 'ENERGETICKÝ CERTIFIKÁT',
        'vykurovanie': 'SPRÁVA O KONTROLE VYKUROVACIEHO SYSTÉMU'
    }.get(typ, 'SPRÁVA')

    h = doc.add_heading(nazov, 0)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER

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
        'vykurovanie': 'SPRÁVA O KONTROLE VYKUROVACIEHO SYSTÉMU'
    }.get(typ, 'SPRÁVA')

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
