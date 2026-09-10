# Energetické správy — Aplikácia

## Čo aplikácia obsahuje
- Verejná úvodná stránka
- Prihlásenie (meno + heslo)
- Modul: Energetický audit
- Modul: Energetický certifikát
- Modul: Kontrola vykurovacích systémov
- Generovanie Word (.docx) a PDF správ
- História všetkých záznamov
- Databáza SQLite

---

## Nahranie na GitHub

1. Prihlás sa na github.com
2. Klikni **New repository**
3. Názov: `energeticke-spravy`
4. Klikni **Create repository**
5. Nahraj všetky súbory cez **Add file → Upload files**

---

## Spustenie na Render.com

1. Prihlás sa na render.com
2. Klikni **New → Web Service**
3. Prepoj GitHub účet a vyber repozitár `energeticke-spravy`
4. Render automaticky detekuje nastavenia z `render.yaml`
5. Klikni **Create Web Service**
6. Počkaj ~3 minúty — aplikácia beží!

---

## Zmena hesla

V súbore `render.yaml` zmeň:
```
APP_USERNAME: admin          ← tvoje meno
APP_PASSWORD: zmenMeHeslo123 ← tvoje heslo
```

Alebo nastav v Render.com → Environment Variables.

---

## Zmena vzorcov

Otvor `app.py` a nájdi sekcie:
- `# VZORCE - ENERGETICKY AUDIT`
- `# VZORCE - ENERGETICKY CERTIFIKAT`
- `# VZORCE - KONTROLA VYKUROVANIA`

Všetky výpočty sú tam prehľadne napísané.

---

## Pripojenie vlastnej domény

1. V Render.com → Settings → Custom Domains
2. Zadaj svoju doménu (napr. moj-audit.sk)
3. Skopíruj DNS záznamy
4. Na Websupport.sk nastav tieto DNS záznamy
5. Počkaj 1-24 hodín na aktiváciu
